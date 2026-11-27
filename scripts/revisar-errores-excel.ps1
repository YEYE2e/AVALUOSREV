param([Parameter(Mandatory)][string]$Libro, [string]$Hallazgos, [string]$Salida)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source = (Resolve-Path -LiteralPath $Libro).ProviderPath
if ([IO.Path]::GetExtension($source) -ine '.xlsx') { throw 'Se requiere un archivo .xlsx.' }
$acA = [char]0x00E1
$acE = [char]0x00E9
$acI = [char]0x00ED
$acO = [char]0x00F3
$acU = [char]0x00FA
$acCapE = [char]0x00C9
$acCapO = [char]0x00D3
$acCapU = [char]0x00DA
$ordM = [char]0x00BA
$emDash = [char]0x2014
$reviewName = "REVISI$($acCapO)N T$($acCapE)CNICA"
$critPattern = "cr[i$acI]tico"
$archive = [IO.Compression.ZipFile]::OpenRead($source)
function Read-Part([string]$Name) {
    $entry = $archive.GetEntry($Name)
    if (!$entry -or $entry.Length -gt 32000000) { throw 'Parte XML ausente o demasiado grande.' }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Safe-Prop($Obj, [string]$Name, [string]$Default = '') {
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return [string]$prop.Value
}
function Rgb-Color([int]$R, [int]$G, [int]$B) {
    return $R + ($G * 256) + ($B * 65536)
}
$findings = @()
try {
    if ($archive.Entries.FullName -match '(?i)(vbaProject\.bin|^xl/(connections\.xml|queryTables/|activeX/|embeddings/|macrosheets/))') {
        throw 'El libro contiene componentes activos que este sistema no abre.'
    }
    $manifest = Read-Part 'xl/workbook.xml'
    if ($Hallazgos) {
        if (!$Salida -or (Test-Path -LiteralPath $Salida)) { throw 'La salida debe ser un archivo nuevo.' }
        $findings = Get-Content -LiteralPath $Hallazgos -Raw -Encoding UTF8 | ConvertFrom-Json
        $findings = @($findings)
        if ($findings.Count -gt 150) { throw 'Cantidad de hallazgos no valida.' }
        foreach ($item in $findings) {
            if ($item.Cell -cnotmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}$') { throw 'Referencia de celda no valida.' }
        }
    } else {
        $relationships = Read-Part 'xl/_rels/workbook.xml.rels'
        foreach ($sheet in $manifest.SelectNodes("//*[local-name()='sheet']")) {
            if ($sheet.name -ieq $reviewName) { throw 'Ya existe la hoja de revision; no se sobrescribe.' }
            $id = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $rel = $relationships.DocumentElement.ChildNodes | Where-Object { $_.Id -ceq $id }
            if ($rel.Type -notlike '*/worksheet') { continue }
            $part = if ($rel.Target.StartsWith('/')) { $rel.Target.TrimStart('/') } else { 'xl/' + $rel.Target }
            if ($part -cnotmatch '^xl/worksheets/sheet[0-9]+\.xml$') { throw 'Ruta de hoja no admitida.' }
            $xml = Read-Part $part
            foreach ($cell in $xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][@t='e']")) {
                $address = $cell.GetAttribute('r')
                if ($address -cnotmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}$') { throw 'Referencia de celda no valida.' }
                $errorCode = $cell.SelectSingleNode("*[local-name()='v']").InnerText
                $detail = switch ($errorCode) {
                    '#DIV/0!' { 'Comprobar si el divisor es cero o si falta el dato que lo alimenta.' }
                    '#REF!' { 'Localizar la referencia invalida y cotejarla con la plantilla original.' }
                    '#VALUE!' { 'Comprobar tipos de datos y errores en las celdas de origen.' }
                    default { 'Inspeccionar el valor de error y las celdas de origen.' }
                }
                $findings += [pscustomobject]@{
                    Sheet = $sheet.name; Cell = $address; Error = $errorCode; Recommendation = $detail;
                    Evidence = "$($sheet.name)!$address = $errorCode"; Status = 'REQUIERE VERIFICACION';
                    Severity = 'Importante'; ReviewType = "C$($acA)lculo / f$($acO)rmula"; FieldReviewed = '';
                    TechnicalExplanation = ''; CurrentValue = $errorCode; ExpectedValue = '';
                    Difference = ''; ValueImpact = 'No determinable'; NotePrefix = 'REQUIERE VERIFICACION'
                }
            }
        }
    }
} finally { $archive.Dispose() }

if ($Hallazgos) { $copy = $Salida }
else {
    $root = Join-Path (Split-Path $PSScriptRoot) 'data'
    $destination = Join-Path $root ([guid]::NewGuid().ToString())
    $null = New-Item -ItemType Directory -Path $destination
    $copy = Join-Path $destination 'avaluo-revision.xlsx'
}
$excel = New-Object -ComObject Excel.Application
$book = $null
$sentinel = $null
try {
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $sentinel = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false
    $book = $excel.Workbooks.Open($source, 0, $true)

    # Group findings by normalized (Sheet, AnchorCell) so multiple observations consolidate cleanly into one cell note
    $grouped = [ordered]@{}
    $resolvedItems = @()
    foreach ($item in $findings) {
        $sheet = $book.Worksheets.Item($item.Sheet)
        if ($Hallazgos -and $sheet.Visible -ne -1) { throw "La hoja objetivo '$($item.Sheet)' no es visible." }
        $cell = $sheet.Range($item.Cell)
        if ($cell.MergeCells) {
            $cell = $cell.MergeArea.Cells.Item(1, 1)
        }
        $anchorCoord = $cell.Address($false, $false)
        if ($Hallazgos -and ($cell.EntireRow.Hidden -or $cell.EntireColumn.Hidden -or
                (!$cell.HasFormula -and [string]::IsNullOrWhiteSpace([string]$cell.Value2)) -or
                ($cell.HasFormula -and [string]$cell.Formula -eq '=""'))) {
            throw "El destino '$($item.Sheet)!$anchorCoord' esta oculto o vacio."
        }
        $key = "$($sheet.Name.ToLowerInvariant())!$anchorCoord"
        if (!$grouped.Contains($key)) {
            $grouped[$key] = [pscustomobject]@{ Sheet = $sheet.Name; Cell = $anchorCoord; Items = @() }
        }
        $grouped[$key].Items += $item
        $resolvedItems += [pscustomobject]@{
            Sheet = $sheet.Name
            Cell = $anchorCoord
            Severity = (Safe-Prop $item 'Severity' 'Importante')
            ReviewType = (Safe-Prop $item 'ReviewType' "Revisi$($acO)n t$($acE)cnica")
            FieldReviewed = (Safe-Prop $item 'FieldReviewed' '')
            Error = (Safe-Prop $item 'Error' '')
            TechnicalExplanation = (Safe-Prop $item 'TechnicalExplanation' '')
            CurrentValue = (Safe-Prop $item 'CurrentValue' '')
            ExpectedValue = (Safe-Prop $item 'ExpectedValue' '')
            Difference = (Safe-Prop $item 'Difference' '')
            Recommendation = (Safe-Prop $item 'Recommendation' '')
            ValueImpact = (Safe-Prop $item 'ValueImpact' 'No determinable')
            Status = (Safe-Prop $item 'Status' 'REQUIERE VERIFICACION')
            Evidence = (Safe-Prop $item 'Evidence' '')
        }
    }

    foreach ($entry in $grouped.Values) {
        $sheet = $book.Worksheets.Item($entry.Sheet)
        $cell = $sheet.Range($entry.Cell)
        $blocks = @()
        foreach ($it in $entry.Items) {
            $sev = Safe-Prop $it 'Severity' 'Importante'
            $st = Safe-Prop $it 'Status' 'REQUIERE VERIFICACION'
            $prefix = Safe-Prop $it 'NotePrefix' "REVISI$($acCapO)N"
            $err = Safe-Prop $it 'Error' ''
            $ev = Safe-Prop $it 'Evidence' ''
            $rec = Safe-Prop $it 'Recommendation' ''
            $impact = Safe-Prop $it 'ValueImpact' 'No determinable'
            if ($Hallazgos) {
                $blocks += "$prefix [$($sev.ToUpperInvariant()) | $st]: $err`nEvidencia: $ev`nRecomendaci$($acO)n: $rec`nIncidencia en el valor: $impact"
            } else {
                $blocks += "REQUIERE VERIFICACION: $err guardado. $rec Incidencia en el valor no determinada."
            }
        }
        $note = ($blocks -join "`n---`n")
        $annotation = $cell.CommentThreaded
        if ($annotation) {
            $written = $annotation.AddReply($note)
            if (!$written) { throw 'No se guardo la respuesta en el comentario existente.' }
        } else {
            $existingNote = $cell.Comment
            if ($existingNote) {
                $note = "Nota previa:`n" + $existingNote.Text() + "`n`n" + $note
                $existingNote.Delete()
            }
            try {
                $written = $cell.AddCommentThreaded($note)
            } catch [Runtime.InteropServices.COMException] {
                $written = $cell.AddComment($note)
                $written.Shape.Width = 400
                $written.Shape.Height = [Math]::Min(460, 200 + (80 * $entry.Items.Count))
                $written.Visible = $true
            }
            if (!$written) { throw 'No se guardo la explicacion del hallazgo.' }
        }
        $cell.Interior.Color = 255
    }

    # Pick unique sheet name without overwriting any pre-existing sheet
    $targetSheetName = $reviewName
    $suffix = 2
    while ($true) {
        $collision = $false
        foreach ($existing in $book.Worksheets) {
            if ($existing.Name -ieq $targetSheetName) { $collision = $true; break }
        }
        if (!$collision) { break }
        if (!$Hallazgos) { throw 'Ya existe la hoja de revision; no se sobrescribe.' }
        $targetSheetName = "$reviewName ($suffix)"
        $suffix++
    }

    $report = $book.Worksheets.Add($book.Worksheets.Item(1))
    $report.Name = $targetSheetName
    $report.Range('A1').Value2 = if ($Hallazgos) { "$reviewName DE AVAL$($acCapU)O $emDash OBSERVACIONES PARA VALIDACI$($acCapO)N DEL PERITO" } else { 'Revision local de errores de Excel guardados; no es una auditoria integral.' }
    $report.Range('A1').Font.Bold = $true
    $report.Range('A1').Font.Size = 13

    $critCount = @($resolvedItems | Where-Object { $_.Severity -imatch $critPattern }).Count
    $impCount = @($resolvedItems | Where-Object { $_.Severity -imatch 'importante' }).Count
    $menCount = @($resolvedItems | Where-Object { $_.Severity -imatch 'menor' }).Count
    $report.Range('A2').Value2 = "Observaciones marcadas: $($resolvedItems.Count) (Cr$($acI)ticas: $critCount | Importantes: $impCount | Menores: $menCount) | Celdas marcadas: $($grouped.Count)."
    $report.Range('A3').Value2 = "El libro original no fue modificado ni recalculado. Este reporte apoya la revisi$($acO)n profesional y no autoriza la entrega bancaria."
    $report.Range('A3').Font.Italic = $true

    if ($Hallazgos) {
        $headers = @(
            "N.$ordM", 'Importancia', "Tipo de revisi$($acO)n", 'Hoja', 'Celda', 'Dato revisado',
            'Hallazgo', "Explicaci$($acO)n t$($acE)cnica", 'Valor actual', 'Valor referencia',
            'Diferencia', "Correcci$($acO)n recomendada", 'Incidencia en valor', 'Estado',
            'Evidencias', "Ubicaci$($acO)n"
        )
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $hc = $report.Cells.Item(5, ($i + 1))
            $hc.Value2 = $headers[$i]
            $hc.Font.Bold = $true
            $hc.Font.Color = Rgb-Color 255 255 255
            $hc.Interior.Color = Rgb-Color 30 36 43
            $hc.WrapText = $true
        }
        $row = 6
        $num = 1
        foreach ($item in $resolvedItems) {
            $values = @(
                $num, $item.Severity, $item.ReviewType, $item.Sheet, $item.Cell, $item.FieldReviewed,
                $item.Error, $item.TechnicalExplanation, $item.CurrentValue, $item.ExpectedValue,
                $item.Difference, $item.Recommendation, $item.ValueImpact, $item.Status, $item.Evidence
            )
            for ($i = 0; $i -lt $values.Count; $i++) {
                $target = $report.Cells.Item($row, ($i + 1))
                $target.NumberFormat = '@'
                $target.Value2 = [string]$values[$i]
                $target.WrapText = $true
                $target.VerticalAlignment = -4160
            }
            $sevCell = $report.Cells.Item($row, 2)
            if ($item.Severity -imatch $critPattern) {
                $sevCell.Interior.Color = Rgb-Color 255 214 214
                $sevCell.Font.Bold = $true
            } elseif ($item.Severity -imatch 'importante') {
                $sevCell.Interior.Color = Rgb-Color 255 236 204
                $sevCell.Font.Bold = $true
            } else {
                $sevCell.Interior.Color = Rgb-Color 235 235 235
            }
            $sheetReference = "'" + $item.Sheet.Replace("'", "''") + "'!" + $item.Cell
            $null = $report.Hyperlinks.Add($report.Range("P$row"), '', $sheetReference, 'Abrir celda marcada', 'Ir a la celda')
            $row++
            $num++
        }
        $widths = @{
            'A'=6; 'B'=14; 'C'=20; 'D'=22; 'E'=10; 'F'=22;
            'G'=38; 'H'=44; 'I'=18; 'J'=18; 'K'=18; 'L'=40;
            'M'=18; 'N'=22; 'O'=42; 'P'=15
        }
        foreach ($col in $widths.Keys) {
            $report.Columns.Item($col).ColumnWidth = $widths[$col]
        }
        if ($row -gt 6) {
            $report.Rows.Item("6:$($row - 1)").AutoFit() | Out-Null
        }
    } else {
        $headers = @('Hoja', 'Celda', 'Hallazgo', 'Recomendacion', 'Evidencia', 'Estado')
        for ($i = 0; $i -lt $headers.Count; $i++) { $report.Cells.Item(5, ($i + 1)).Value2 = $headers[$i] }
        $report.Range('A5:F5').Font.Bold = $true
        $report.Range('G5').Value2 = 'Ubicacion'
        $row = 6
        foreach ($item in $resolvedItems) {
            $values = @($item.Sheet, $item.Cell, $item.Error, $item.Recommendation, $item.Evidence, $item.Status)
            for ($i = 0; $i -lt $values.Count; $i++) {
                $target = $report.Cells.Item($row, ($i + 1))
                $target.NumberFormat = '@'
                $target.Value2 = [string]$values[$i]
            }
            $sheetReference = "'" + $item.Sheet.Replace("'", "''") + "'!" + $item.Cell
            $null = $report.Hyperlinks.Add($report.Range("G$row"), '', $sheetReference, 'Abrir celda marcada', 'Ir a la celda')
            $row++
        }
        $report.Columns.Item('A:F').ColumnWidth = 24
        $report.Columns.Item('D').ColumnWidth = 65
        $report.Columns.Item('G').ColumnWidth = 19
    }

    $report.Activate()
    $book.SaveAs($copy, 51)

    if ($Hallazgos) {
        $book.Close($false)
        $book = $null
        $book = $excel.Workbooks.Open($copy, 0, $true)
        foreach ($entry in $grouped.Values) {
            $cell = $book.Worksheets.Item($entry.Sheet).Range($entry.Cell)
            $annotation = $cell.CommentThreaded
            if ($annotation) {
                $messages = @($annotation.Text())
                foreach ($reply in $annotation.Replies) { $messages += $reply.Text() }
            } else {
                $annotation = $cell.Comment
                if (!$annotation -or !$annotation.Visible) { throw 'No se guardo la anotacion visible de Excel.' }
                $messages = @($annotation.Text())
            }
            foreach ($it in $entry.Items) {
                $errText = Safe-Prop $it 'Error' ''
                $evText = Safe-Prop $it 'Evidence' ''
                $recText = Safe-Prop $it 'Recommendation' ''
                $complete = $false
                foreach ($message in $messages) {
                    if ($message.Contains($errText) -and $message.Contains($evText) -and $message.Contains($recText)) {
                        $complete = $true; break
                    }
                }
                if (!$complete) { throw "No se guardo el comentario con el detalle completo en $($entry.Sheet)!$($entry.Cell)." }
            }
        }
    }
} finally {
    if ($book) { $book.Close($false) }
    if ($sentinel) { $sentinel.Close($false) }
    $excel.Quit()
    $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
if ($Hallazgos) { return }
$findings | Select-Object @{Name='Hoja';Expression={"'" + $_.Sheet}}, Cell, Error, Recommendation |
    Export-Csv -LiteralPath (Join-Path $destination 'informe.csv') -NoTypeInformation -Encoding UTF8
Write-Output "Copia e informe: $destination ($($findings.Count) celdas)"
