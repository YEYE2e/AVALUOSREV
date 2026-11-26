# Decisiones arquitect?nicas

## Alcance y componentes

Monolito modular: React/TypeScript/Vite muestra el estado; FastAPI ser? la frontera HTTP; SQLite almacenar? expedientes, documentos, ejecuciones, eventos, hallazgos y decisiones humanas. `backend/app/models.py` define contratos iniciales; `agents.py` define el puerto de proveedor; `workspace.py` es una primitiva de almacenamiento, no un servicio de carga completo. `main.py` expone ?nicamente salud. El esquema JSON se genera con `python -m app.schema` desde el modelo, evitando duplicar la fuente de verdad. `skills/`, `prompts/` y `evals/` contienen los recursos versionados; no provienen de documentos del usuario.

## Flujo previsto

Crear expediente ? cargar originales y registrar hash, tama?o y nombre de presentaci?n ? crear ejecuci?n ? procesar en espacio aislado ? generar resultado estructurado ? verificar independientemente las referencias y su contenido ? presentar incertidumbres y revisi?n humana ? registrar decisiones e informe. Las reglas de dominio est?n pendientes de especialistas. Las categor?as distinguen confirmado, discrepancia, no verificable y revisi?n humana. Un contrato v?lido no acredita que una evidencia exista ni sustente el hallazgo.

## Ejecuci?n prevista

La interfaz `AgentProvider` separa proveedor y l?gica principal. El adaptador Codex es un marcador deshabilitado: no inicia procesos ni accede a la red. El futuro coordinador deber? persistir transiciones y eventos en SQLite, permitir cancelaci?n y supervisar procesos; SSE deber? reconectar mediante IDs y conservar eventos. Definir recuperaci?n tras reinicio, concurrencia SQLite y l?mites antes de activarlo. Registrar proveedor, modelo, errores sanitizados y consumo disponible sin inventar valores.

## Seguridad y l?mites

Los documentos y p?ginas son datos no confiables, nunca instrucciones. El futuro ejecutor deber? imponer permisos fuera del prompt, aislar procesos, limitar recursos y red, controlar SSRF, conservar originales con acceso de solo lectura y no enviar secretos al agente. La primitiva actual usa UUID, creaci?n exclusiva y comprueba la ra?z resuelta; no ofrece defensa completa frente a carreras o enlaces manipulados por otro proceso. Antes de admitir archivos se necesitan l?mites de tama?o, tipos y contenido, protecci?n contra enlaces/carreras, permisos de sistema, cuarentena y limpieza de archivos parciales. Los originales se identifican por hash; su integridad debe comprobarse antes y despu?s del procesamiento.

Sin autenticaci?n, autorizaci?n, cifrado, retenci?n, copias de seguridad ni separaci?n de usuarios implementados, esta base no est? preparada para datos reales ni despliegue compartido. Los requisitos de acceso, privacidad y retenci?n deben acordarse antes de habilitarla. La consulta de URL requiere controles de destino, redirecci?n y acceso a redes internas.

## Interfaz

Pantalla inicial con estado vac?o expl?cito. Creaci?n de expedientes, visualizaci?n de documentos, seguimiento SSE, revisi?n e historial pertenecen a etapas posteriores; no se simulan.
