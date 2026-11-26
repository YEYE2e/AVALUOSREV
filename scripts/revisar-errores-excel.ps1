param([Parameter(Mandatory)][string]$Libro, [string]$Hallazgos, [string]$Salida)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source = (Resolve-Path -LiteralPath $Libro).ProviderPath
if ([IO.Path]::GetExtension($source) -ine '.xlsx') { throw 'Se requiere un archivo .xlsx.' }
$reviewName = 'REVISI' + [char]0x00D3 + 'N T' + [char]0x00C9 + 'CNICA'
$archive = [IO.Compression.ZipFile]::OpenRead($source)
function Read-Part([string]$Name) {
    $entry = $archive.GetEntry($Name)
    if (!$entry -or $entry.Length -gt 32000000) { throw 'Parte XML ausente o demasiado grande.' }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}
$findings = @()
try {
    if ($archive.Entries.FullName -match '(?i)(vbaProject\.bin|^xl/(connections\.xml|queryTables/|activeX/|embeddings/|macrosheets/))') {
        throw 'El libro contiene componentes activos que este piloto no abre.'
    }
    $manifest = Read-Part 'xl/workbook.xml'
    if ($Hallazgos) {
        if (!$Salida -or (Test-Path -LiteralPath $Salida)) { throw 'La salida debe ser un archivo nuevo.' }
        $findings = Get-Content -LiteralPath $Hallazgos -Raw -Encoding UTF8 | ConvertFrom-Json
        $findings = @($findings)
        if ($findings.Count -gt 50) { throw 'Cantidad de hallazgos no valida.' }
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
            $findings += [pscustomobject]@{ Sheet = $sheet.name; Cell = $address; Error = $errorCode; Recommendation = $detail }
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
try {
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $sentinel = $excel.Workbooks.Add()
    $excel.Calculation = -4135
    $excel.CalculateBeforeSave = $false
    $book = $excel.Workbooks.Open($source, 0, $true)
    foreach ($item in $findings) {
        $sheet = $book.Worksheets.Item($item.Sheet)
        if ($Hallazgos -and $sheet.Visible -ne -1) { throw 'La hoja objetivo no es visible.' }
        $cell = $sheet.Range($item.Cell)
        if ($Hallazgos -and $cell.MergeCells -and $cell.MergeArea.Cells.Item(1,1).Address() -ne $cell.Address()) {
            throw 'El destino no es el ancla de la celda combinada.'
        }
        if ($Hallazgos -and ($cell.EntireRow.Hidden -or $cell.EntireColumn.Hidden -or
                (!$cell.HasFormula -and [string]::IsNullOrWhiteSpace([string]$cell.Value2)) -or
                ($cell.HasFormula -and [string]$cell.Formula -eq '=""'))) {
            throw 'El destino esta oculto o vacio.'
        }
        $note = if ($Hallazgos) {
            "$($item.Status): $($item.Error)`nEvidencia: $($item.Evidence)`nRecomendacion: $($item.Recommendation)"
        } else {
            'REQUIERE VERIFICACION: ' + $item.Error + ' guardado. ' + $item.Recommendation + ' Incidencia en el valor no determinada.'
        }
        $annotation = $cell.CommentThreaded
        if ($annotation) {
            $written = $annotation.AddReply($note)
            if (!$written -or !$written.Text().Contains($item.Error)) {
                throw 'No se guardo la respuesta en el comentario existente.'
            }
        } else {
            $existingNote = $cell.Comment
            if ($existingNote) {
                $note = "Nota previa:`n" + $existingNote.Text() + "`n`n" + $note
                $existingNote.Delete()
            }
            try {
                $written = $cell.AddCommentThreaded($note)
            } catch [Runtime.InteropServices.COMException] {
                # Some desktop editions expose the method but reject threaded comments.
                $written = $cell.AddComment($note)
                $written.Shape.Width = 360
                $written.Shape.Height = 220
                $written.Visible = $true
            }
            if (!$written -or !$written.Text().Contains($item.Error)) {
                throw 'No se guardo la explicacion del hallazgo.'
            }
        }
        $cell.Interior.Color = 255
    }
    foreach ($existing in $book.Worksheets) {
        if ($existing.Name -ieq $reviewName) { throw 'Ya existe la hoja de revision; no se sobrescribe.' }
    }
    $report = $book.Worksheets.Add($book.Worksheets.Item(1))
    $report.Name = $reviewName
    $report.Range('A1').Value2 = if ($Hallazgos) { 'Hallazgos para revision humana.' } else { 'Revision local de errores de Excel guardados; no es una auditoria integral.' }
    $report.Range('A2').Value2 = "Celdas marcadas: $($findings.Count). No se recalculo el libro ni se actualizaron vinculos."
    $report.Range('A3').Value2 = 'No constituye una autorizacion de entrega.'
    $headers = @('Hoja', 'Celda', 'Hallazgo', 'Recomendacion', 'Evidencia', 'Estado')
    for ($i = 0; $i -lt $headers.Count; $i++) { $report.Cells.Item(5, ($i + 1)).Value2 = $headers[$i] }
    $report.Range('A5:F5').Font.Bold = $true
    $report.Range('G5').Value2 = 'Ubicacion'
    $row = 6
    foreach ($item in $findings) {
        $values = @($item.Sheet, $item.Cell, $item.Error, $item.Recommendation, $item.Evidence, $(if ($Hallazgos) { $item.Status } else { 'REQUIERE VERIFICACION' }))
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
    $report.Activate()
    $book.SaveAs($copy, 51)
    if ($Hallazgos) {
        $book.Close($false)
        $book = $null
        $book = $excel.Workbooks.Open($copy, 0, $true)
        foreach ($item in $findings) {
            $cell = $book.Worksheets.Item($item.Sheet).Range($item.Cell)
            $annotation = $cell.CommentThreaded
            if ($annotation) {
                $messages = @($annotation.Text())
                foreach ($reply in $annotation.Replies) { $messages += $reply.Text() }
            } else {
                $annotation = $cell.Comment
                if (!$annotation -or !$annotation.Visible) { throw 'No se guardo la anotacion visible de Excel.' }
                $messages = @($annotation.Text())
            }
            $complete = $false
            foreach ($message in $messages) {
                if ($message.Contains($item.Error) -and $message.Contains($item.Evidence) -and
                        $message.Contains($item.Recommendation)) { $complete = $true; break }
            }
            if (!$complete) { throw 'No se guardo el comentario con el detalle completo.' }
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
