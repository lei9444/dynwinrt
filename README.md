# dynwinrt

Dynamic WinRT API invocation — call any Windows Runtime method at runtime without native code generation.

## Overview

`dynwinrt` is a Rust library that uses runtime metadata (.winmd files) and FFI (libffi) to call arbitrary WinRT methods dynamically. It provides a foundation for JavaScript and Python bindings that don't require MSVC compilation or version-specific generated code.

## Repository Structure

```
dynwinrt/
├── crates/dynwinrt/       # Core Rust library
├── bindings/
│   ├── js/                # JavaScript/TypeScript bindings (napi-rs)
│   └── py/                # Python bindings (PyO3)
└── tools/
    └── winrt-meta/        # → d:\work\winrt-meta (code generator)
```

## Build

```bash
# Build the core library
cargo build -p dynwinrt

# Run tests
cargo test -p dynwinrt

# Build JS bindings
cd bindings/js && npm install && npx napi build --no-const-enum --platform --release -o dist

# Build Python bindings
cd bindings/py && maturin develop
```

## Code Generation with winrt-meta

`winrt-meta` reads Windows metadata (.winmd) files and generates typed TypeScript bindings that use `dynwinrt-js` at runtime.

### Step 1: Build winrt-meta

```bash
cd d:\work\winrt-meta
cargo build --release
```

### Step 2: Generate Bindings

```bash
# Generate for a specific class
cargo run --release -- generate \
  --winmd "C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd" \
  --namespace "Windows.Foundation" \
  --class-name "Uri" \
  --lang ts \
  --output ./generated/Windows.Foundation

# Generate for an entire namespace
cargo run --release -- generate \
  --winmd "C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd" \
  --namespace "Windows.Web.Http" \
  --lang ts \
  --output ./generated/Windows.Web.Http
```

**Arguments:**

| Argument | Required | Description |
|---|---|---|
| `--winmd` | Yes | Path to .winmd file(s), separated by `;` |
| `--namespace` | Yes | WinRT namespace to generate |
| `--class-name` | No | Specific class (generates dependencies too) |
| `--lang` | No | Target language: `ts` (default) |
| `--output` | No | Output directory |

### Step 3: Fix Import Paths (local development)

Generated files import from `'dynwinrt-js'`. For local development, fix to relative path:

```bash
# Replace package import with relative path to built bindings
find generated -name "*.ts" -exec sed -i "s|from 'dynwinrt-js'|from '../../dist/index.js'|g" {} +
```

### Step 4: Use Generated Bindings

```typescript
import { roInitialize } from 'dynwinrt-js'
import { Uri } from './generated/Windows.Foundation/Uri'

roInitialize(1) // Initialize WinRT (MTA)

const uri = Uri.createUri('https://example.com/path?q=1')
console.log(uri.host)       // "example.com"
console.log(uri.port)       // 443
console.log(uri.schemeName) // "https"
```

### What Gets Generated

For each WinRT class, winrt-meta generates:

- **Interface registration** — `DynWinRtType.registerInterface()` with all methods and type signatures
- **Wrapper class** — TypeScript class with typed properties and methods
- **Factory methods** — Static methods for object creation (via activation factory)
- **Collection types** — `IVector<T>`, `IMap<K,V>`, etc. wrappers in `_collections.ts`
- **Enums** — TypeScript `enum` declarations

### Running Tests

```bash
# Core library tests
cargo test -p dynwinrt

# JS binding test (uses generated pattern)
cd bindings/js && npx tsx __test__/test_generated.ts

# Python binding tests
cd bindings/py && pytest
```

## Use WinAppSDK Bootstrap

The path to the WinAppSDK Bootstrap DLL is retrieved from the `WINAPPSDK_BOOTSTRAP_DLL_PATH` environment variable. Only needed for unpackaged apps using WinAppSDK APIs.

```typescript
import { initWinappsdk } from 'dynwinrt-js'
initWinappsdk(1, 8) // Initialize WinAppSDK 1.8
```

## Architecture

### Dynamic Call Strategies

dynwinrt selects an optimized call strategy at method build time:

| Strategy | When | Uses libffi? |
|---|---|---|
| `Direct0In1Out` | Property getter (0 in, 1 out) | No |
| `Direct1In1Out` | Factory/method (1 in, 1 out) | No |
| `Direct1In0Out` | Property setter (1 in, 0 out) | No |
| `Libffi(Cif)` | General case (2+ params, structs) | Yes (cached Cif) |

### Type System

Three-layer mapping: `TypeKind` (compile-time descriptor) → `AbiValue` (ABI representation) → `WinRTValue` (runtime value container).

Supports: primitives, HString, Object, GUID, structs, arrays, async operations, parameterized generics (`IVector<T>`, `IReference<T>`, etc.).
