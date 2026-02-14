# LibreOffice Calc MCP (Rust)

Servidor MCP (Model Context Protocol) en Rust para manipular ficheros `.ods` (LibreOffice Calc) en local, sin abrir LibreOffice.

## 1. Descripción general

Este proyecto implementa un servidor MCP por `stdio` que expone tools para crear, leer y modificar hojas de cálculo ODS.

Objetivos del MVP actual:
- Operar sobre rutas locales `.ods`.
- Mantener respuestas JSON simples para consumo por agentes LLM.
- Entregar una base modular, testeada y lista para integrar en clientes MCP.

Estado actual:
- MVP funcional completado.
- Build/test/clippy en verde.
- Release local generado (`target/release/mcp-ods.exe`).
- Rama principal publicada en GitHub.

## 2. Funcionalidades (tools MCP)

Tools implementadas:
1. `create_ods`
2. `get_sheets`
3. `get_sheet_content`
4. `set_cell_value`
5. `duplicate_sheet`
6. `add_sheet`
7. `get_cell_value`
8. `set_range_values`

### Resumen rápido de cada tool

- `create_ods`: crea un `.ods` válido desde cero (incluyendo estructura ZIP mínima requerida).
- `get_sheets`: devuelve nombres de hojas en orden interno.
- `get_sheet_content`: devuelve matriz 2D (`mode=matrix`) con límites y opción de trimming.
- `set_cell_value`: escribe una celda A1 concreta.
- `duplicate_sheet`: clona una hoja y la inserta tras la original.
- `add_sheet`: añade hoja vacía (inicio o final).
- `get_cell_value`: devuelve valor tipado de una celda.
- `set_range_values`: escribe bloque matricial desde una celda inicial.

## 3. Arquitectura del proyecto

Estructura principal:

```text
src/
  main.rs
  lib.rs
  common/
    errors.rs
    fs.rs
    json.rs
  mcp/
    server.rs
    protocol.rs
    dispatcher.rs
  ods/
    ods_file.rs
    ods_templates.rs
    manifest.rs
    content_xml.rs
    sheet_model.rs
    cell_address.rs
  tools/
    create_ods.rs
    get_sheets.rs
    get_sheet_content.rs
    set_cell_value.rs
    duplicate_sheet.rs
    add_sheet.rs
    get_cell_value.rs
    set_range_values.rs

test/
  mod.rs
  common/
  mcp/
  ods/
  tools/
  blackbox/
```

Diseño por capas:
- `mcp/*`: transporte stdio + JSON-RPC + dispatch de tools.
- `tools/*`: contrato de entrada/salida por tool y reglas de negocio.
- `ods/*`: lectura/escritura ODS (ZIP + XML) y modelo de workbook.
- `common/*`: errores y utilidades compartidas.

## 4. Requisitos

- Rust estable (recomendado: toolchain actualizado).
- Cargo.
- Windows, Linux o macOS (el proyecto usa stdio y filesystem local).

## 5. Cómo compilar y testear

### Build debug

```bash
cargo build
```

### Build release

```bash
cargo build --release
```

### Formato

```bash
cargo fmt
```

### Lint estricto

```bash
cargo clippy -- -D warnings
```

### Tests

```bash
cargo test
```

## 6. Ejecución manual del servidor

### Binario debug

```bash
cargo run
```

### Binario release

```bash
./target/release/mcp-ods
```

En Windows:

```powershell
.\target\release\mcp-ods.exe
```

El servidor espera una petición JSON-RPC por línea en `stdin` y escribe respuesta JSON-RPC en `stdout`.

## 7. Ejemplos JSON-RPC

### `create_ods`

Request:

```json
{"jsonrpc":"2.0","id":1,"method":"create_ods","params":{"path":"./demo.ods","overwrite":true,"initial_sheet_name":"Hoja1"}}
```

Response esperada:

```json
{"jsonrpc":"2.0","id":1,"result":{"path":".../demo.ods","sheets":["Hoja1"]}}
```

### `set_cell_value`

```json
{"jsonrpc":"2.0","id":2,"method":"set_cell_value","params":{"path":"./demo.ods","sheet":{"index":0},"cell":"B2","value":{"type":"string","data":"hola"}}}
```

### `get_cell_value`

```json
{"jsonrpc":"2.0","id":3,"method":"get_cell_value","params":{"path":"./demo.ods","sheet":{"index":0},"cell":"B2"}}
```

También se soporta la forma MCP envelope `tools/call`:

```json
{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{"name":"get_sheets","arguments":{"path":"./demo.ods"}}}
```

## 8. Configuración para enlazarlo con un cliente de agentes

Este servidor usa transporte `stdio`. Un cliente MCP debe lanzar el binario y hablar JSON-RPC por stdin/stdout.

### Config genérica (conceptual)

```json
{
  "mcpServers": {
    "libreoffice-calc": {
      "command": "C:/workspace/mcp/target/release/mcp-ods.exe",
      "args": []
    }
  }
}
```

Puntos importantes:
- Usa ruta absoluta al binario para evitar problemas de working directory.
- Si ejecutas en debug, puedes apuntar a `target/debug/mcp-ods(.exe)`.
- Asegura permisos de lectura/escritura sobre las rutas `.ods` que vaya a tocar el agente.

## 9. Tests: estrategia actual

Hay dos niveles:

- Unitarios (`test/common`, `test/mcp`, `test/ods`, `test/tools`):
  - Validación de utilidades, parseos y handlers por módulo.
- Integración/blackbox (`test/blackbox/*`):
  - Flujos end-to-end por tool y comportamiento sobre ficheros `.ods` reales temporales.

Cobertura actual incluye:
- operaciones CRUD de hojas,
- lectura/escritura de celdas y rangos,
- validación de errores,
- serialización/deserialización de valores tipados.

## 10. Errores y códigos

Errores tipados principales (`src/common/errors.rs`):
- `InvalidPath`
- `FileNotFound`
- `AlreadyExists`
- `InvalidOdsFormat`
- `SheetNotFound`
- `SheetNameAlreadyExists`
- `InvalidCellAddress`
- `XmlParseError`
- `ZipError`
- `IoError`

Cada error se mapea a código numérico estable para clientes MCP.

## 11. Limitaciones conocidas del MVP

- `get_sheet_content` soporta `mode="matrix"`.
- `set_range_values` escribe actualmente valores como `string`.
- El modelo XML cubre el flujo de datos principal; no preserva de forma exhaustiva todos los detalles avanzados de formato/fórmulas complejas fuera del modelo actual.

## 12. Roadmap sugerido

Siguientes mejoras recomendadas:
1. Soporte explícito de fórmulas y tipos mixtos en `set_range_values`.
2. Preservación ampliada de estilos/atributos avanzados por celda.
3. Tooling MCP adicional (`initialize`, `tools/list`) si se necesita compatibilidad estricta con ciertos clientes.
4. Benchmarks y tests de regresión con ODS grandes.

## 13. Historial de hitos (resumen)

- `v0.1.0` etiquetado como release MVP local.
- Commits recientes:
  - implementación MCP ODS,
  - ampliación/reorganización de tests,
  - comentarios explicativos para onboarding.

## 14. Licencia

Pendiente de definir. Añadir `LICENSE` según la política del repositorio.