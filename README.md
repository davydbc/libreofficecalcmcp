# LibreOffice Calc MCP (Rust)

## 1. Breve descripción del proyecto
Servidor MCP (Model Context Protocol) en Rust para trabajar con ficheros `.ods` de LibreOffice Calc mediante `stdio`, sin abrir LibreOffice.

Permite crear, leer y modificar hojas/celdas de forma programática para integrarlo con clientes de agentes (por ejemplo Cline).

## 2. Tools disponibles

### `create_ods`
- Descripción: crea un fichero `.ods` válido con una hoja inicial.
- Entrada:
```json
{
  "path": "string",
  "overwrite": "boolean (opcional, por defecto false)",
  "initial_sheet_name": "string (opcional, por defecto Hoja1)"
}
```
- Salida:
```json
{
  "path": "string",
  "sheets": ["string"]
}
```

### `get_sheets`
- Descripción: devuelve los nombres de hojas en orden.
- Entrada:
```json
{
  "path": "string"
}
```
- Salida:
```json
{
  "sheets": ["string"]
}
```

### `get_sheet_content`
- Descripción: devuelve el contenido de una hoja como matriz 2D.
- Entrada:
```json
{
  "path": "string",
  "sheet": { "index": 0 } | { "name": "string" },
  "mode": "matrix (opcional)",
  "max_rows": "integer (opcional)",
  "max_cols": "integer (opcional)",
  "include_empty_trailing": "boolean (opcional)"
}
```
- Salida:
```json
{
  "sheet": "string",
  "rows": "integer",
  "cols": "integer",
  "data": [["string"]]
}
```

### `set_cell_value`
- Descripción: escribe un valor en una celda A1.
- Entrada:
```json
{
  "path": "string",
  "sheet": { "index": 0 } | { "name": "string" },
  "cell": "string (A1)",
  "value": {
    "type": "string | number | boolean | empty",
    "data": "any (según type)"
  }
}
```
- Salida:
```json
{
  "updated": true,
  "sheet": "string",
  "cell": "string"
}
```

### `duplicate_sheet`
- Descripción: duplica una hoja e inserta la copia justo después.
- Entrada:
```json
{
  "path": "string",
  "source_sheet": { "index": 0 } | { "name": "string" },
  "new_sheet_name": "string"
}
```
- Salida:
```json
{
  "sheets": ["string"]
}
```

### `add_sheet`
- Descripción: añade una hoja vacía al inicio o al final.
- Entrada:
```json
{
  "path": "string",
  "sheet_name": "string",
  "position": "start | end (opcional, por defecto end)"
}
```
- Salida:
```json
{
  "sheets": ["string"]
}
```

### `get_cell_value`
- Descripción: lee el valor tipado de una celda.
- Entrada:
```json
{
  "path": "string",
  "sheet": { "index": 0 } | { "name": "string" },
  "cell": "string (A1)"
}
```
- Salida:
```json
{
  "sheet": "string",
  "cell": "string",
  "value": {
    "type": "string | number | boolean | empty",
    "data": "any (si aplica)"
  }
}
```

### `set_range_values`
- Descripción: escribe una matriz desde una celda inicial.
- Entrada:
```json
{
  "path": "string",
  "sheet": { "index": 0 } | { "name": "string" },
  "start_cell": "string (A1)",
  "data": [["string"]]
}
```
- Salida:
```json
{
  "updated": true,
  "rows_written": "integer",
  "cols_written": "integer"
}
```

## 3. Guía rápida (compilación, tests y uso)

### Compilar
```bash
cargo build
```

### Compilar release
```bash
cargo build --release
```

### Ejecutar tests
```bash
cargo test
```

### Ejecutar servidor MCP
Debug:
```bash
cargo run
```

Windows release:
```powershell
.\target\release\mcp-ods.exe
```

### Uso rápido por terminal (JSON-RPC por línea)
`initialize`:
```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}
```

`tools/list`:
```json
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}
```

`tools/call` (`create_ods`):
```json
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"create_ods","arguments":{"path":"./test.ods","overwrite":true,"initial_sheet_name":"Hoja1"}}}
```

`tools/call` (`set_cell_value`):
```json
{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"set_cell_value","arguments":{"path":"./test.ods","sheet":{"index":0},"cell":"A1","value":{"type":"string","data":"VALOR"}}}}
```

### Configuración mínima para cliente MCP (ejemplo)
```json
{
  "mcpServers": {
    "libreoffice-calc": {
      "command": "C:\\workspace\\mcp\\target\\release\\mcp-ods.exe",
      "args": [],
      "disabled": false
    }
  }
}
```
