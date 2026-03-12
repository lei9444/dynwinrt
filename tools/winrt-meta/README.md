# winrt-meta

Read Windows metadata (`.winmd`) files and generate typed TypeScript bindings that use `dynwinrt-js` at runtime.

## Build

```bash
cargo build -p winrt-meta --release
```

## Usage

```bash
cargo run -p winrt-meta --release -- generate [OPTIONS]
```

### Arguments

| Argument | Required | Description |
|---|---|---|
| `--winmd` | Yes | Path to `.winmd` file(s), separated by `;` |
| `--namespace` | Yes | WinRT namespace to generate |
| `--class` | No | Specific class name (generates its dependencies too) |
| `--lang` | No | Target language: `ts` (default). Only `ts` is supported currently |
| `--output` | No | Output directory (default: `./generated`) |

### Examples

Generate bindings for a specific class:

```bash
cargo run -p winrt-meta --release -- generate \
  --winmd "C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd" \
  --namespace "Windows.Storage.Pickers" \
  --class "FileOpenPicker" \
  --output ./generated/Windows.Storage.Pickers
```

Generate bindings for an entire namespace:

```bash
cargo run -p winrt-meta --release -- generate \
  --winmd "C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd" \
  --namespace "Windows.Web.Http" \
  --output ./generated/Windows.Web.Http
```

Generate with multiple `.winmd` files (e.g. Windows SDK + WinAppSDK):

```bash
cargo run -p winrt-meta --release -- generate \
  --winmd "C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd;path\to\Microsoft.WindowsAppSDK.winmd" \
  --namespace "Microsoft.UI.Xaml" \
  --output ./generated/Microsoft.UI.Xaml
```

## Output

For each WinRT class, the tool generates:

- **Interface registration** -- `DynWinRtType.registerInterface()` with all methods and type signatures
- **Wrapper class** -- TypeScript class with typed properties and methods
- **Factory methods** -- Static methods for object creation via activation factory
- **Enums** -- TypeScript `enum` declarations
- **Collection types** -- `IVector<T>`, `IVectorView<T>`, `IMap<K,V>`, etc. in `_collections.ts`
- **Index file** -- `index.ts` re-exporting all generated types

Dependencies are resolved automatically -- specifying `--class FileOpenPicker` will also generate `StorageFile`, `StorageFolder`, `Uri`, enums, and other referenced types.

## Using Generated Bindings

```typescript
import { roInitialize } from 'dynwinrt-js';
import { FileOpenPicker } from './generated/Windows.Storage.Pickers/FileOpenPicker';
import { PickerViewMode } from './generated/Windows.Storage.Pickers/PickerViewMode';

roInitialize(1); // Initialize WinRT (MTA)

const picker = FileOpenPicker.create();
picker.viewMode = PickerViewMode.Thumbnail;
picker.fileTypeFilter.append('.png');
picker.fileTypeFilter.append('.jpg');

const file = await picker.pickSingleFileAsync();
```

### Local Development

Generated files import from `'dynwinrt-js'`. For local development without the npm package, fix the import path:

```bash
find generated -name "*.ts" -exec sed -i "s|from 'dynwinrt-js'|from '../../dist/index.js'|g" {} +
```
