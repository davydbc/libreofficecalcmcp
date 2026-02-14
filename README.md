# LibreOffice Calc MCP (Rust)

A Rust MCP (Model Context Protocol) server to manipulate local `.ods` files (LibreOffice Calc) without opening LibreOffice.

## 1. Overview

This project implements a `stdio` MCP server that exposes tools to create, read, and modify ODS spreadsheets.

Current MVP goals:
- Work with local `.ods` paths.
- Keep JSON responses simple for LLM agent consumption.
- Provide a modular, tested base ready to integrate with MCP clients.

Current status:
- Functional MVP completed.
- Build/test/clippy all green.
- Local release built (`target/release/mcp-ods.exe`).
- Main branch published on GitHub.

## 2. Features (MCP tools)

Implemented tools:
1. `create_ods`
2. `get_sheets`
3. `get_sheet_content`
4. `set_cell_value`
5. `duplicate_sheet`
6. `add_sheet`
7. `get_cell_value`
8. `set_range_values`

### Quick summary per tool

- `create_ods`: creates a valid `.ods` file from scratch (including required minimal ZIP structure).
- `get_sheets`: returns sheet names in internal order.
- `get_sheet_content`: returns a 2D matrix (`mode=matrix`) with limits and optional trimming.
- `set_cell_value`: writes a single A1 cell.
- `duplicate_sheet`: clones a sheet and inserts it right after the original.
- `add_sheet`: adds an empty sheet (start or end).
- `get_cell_value`: returns a typed value for one cell.
- `set_range_values`: writes a matrix block starting at a given cell.

## 3. Project architecture

Main structure:

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

Layered design:
- `mcp/*`: stdio transport + JSON-RPC + tool dispatch.
- `tools/*`: input/output contracts and business rules per tool.
- `ods/*`: ODS read/write (ZIP + XML) and workbook model.
- `common/*`: shared errors and utilities.

## 4. Requirements

- Stable Rust (updated toolchain recommended).
- Cargo.
- Windows, Linux, or macOS (the project uses stdio + local filesystem).

## 5. Build and test

### Debug build

```bash
cargo build
```

### Release build

```bash
cargo build --release
```

### Format

```bash
cargo fmt
```

### Strict lint

```bash
cargo clippy -- -D warnings
```

### Tests

```bash
cargo test
```

## 6. Run the server manually

### Debug binary

```bash
cargo run
```

### Release binary

```bash
./target/release/mcp-ods
```

On Windows:

```powershell
.\target\release\mcp-ods.exe
```

The server expects one JSON-RPC request per line on `stdin` and writes JSON-RPC responses to `stdout`.

## 7. JSON-RPC examples

### `create_ods`

Request:

```json
{"jsonrpc":"2.0","id":1,"method":"create_ods","params":{"path":"./demo.ods","overwrite":true,"initial_sheet_name":"Sheet1"}}
```

Expected response:

```json
{"jsonrpc":"2.0","id":1,"result":{"path":".../demo.ods","sheets":["Sheet1"]}}
```

### `set_cell_value`

```json
{"jsonrpc":"2.0","id":2,"method":"set_cell_value","params":{"path":"./demo.ods","sheet":{"index":0},"cell":"B2","value":{"type":"string","data":"hello"}}}
```

### `get_cell_value`

```json
{"jsonrpc":"2.0","id":3,"method":"get_cell_value","params":{"path":"./demo.ods","sheet":{"index":0},"cell":"B2"}}
```

The MCP envelope `tools/call` is also supported:

```json
{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{"name":"get_sheets","arguments":{"path":"./demo.ods"}}}
```

## 8. Configuration for agent clients

This server uses `stdio` transport. An MCP client must launch the binary and communicate JSON-RPC via stdin/stdout.

### Generic configuration (conceptual)

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

Important notes:
- Use an absolute binary path to avoid working-directory issues.
- For debug runs, you can point to `target/debug/mcp-ods(.exe)`.
- Ensure read/write permissions for the `.ods` paths your agent will touch.

## 9. Test strategy

There are two levels:

- Unit tests (`test/common`, `test/mcp`, `test/ods`, `test/tools`):
  - Validation of utilities, parsing, and module handlers.
- Integration/blackbox tests (`test/blackbox/*`):
  - End-to-end tool flows on real temporary `.ods` files.

Current coverage includes:
- sheet CRUD operations,
- cell/range read and write,
- error validation,
- typed value serialization/deserialization.

## 10. Errors and codes

Main typed errors (`src/common/errors.rs`):
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

Each error maps to a stable numeric code for MCP clients.

## 11. Known MVP limitations

- `get_sheet_content` supports `mode="matrix"`.
- `set_range_values` currently writes values as `string`.
- The XML model covers the main data path, but does not fully preserve all advanced formatting/formula details outside the current model.

## 12. Suggested roadmap

Recommended next improvements:
1. Explicit formula and mixed-type support in `set_range_values`.
2. Wider preservation of advanced cell style/attributes.
3. Additional MCP tooling (`initialize`, `tools/list`) if strict compatibility is required for specific clients.
4. Benchmarks and regression tests with large ODS files.

## 13. Milestones summary

- `v0.1.0` tagged as local MVP release.
- Recent commits include:
  - MCP ODS implementation,
  - test expansion/reorganization,
  - explanatory comments for onboarding.

## 14. License

Pending definition. Add a `LICENSE` file according to repository policy.