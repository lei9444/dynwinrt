# dynwinrt Sample

A self-contained test kit for dynwinrt-js. Download from GitHub Actions artifacts and run.

## Contents

```
sample/
├── winrt-meta.exe       # codegen tool
├── dynwinrt-js/         # runtime (index.js + index.d.ts + .node files)
├── package.json
├── test_uri.ts          # Windows SDK test (no WinAppSDK needed)
├── test_picker.ts       # WinAppSDK test (requires WinAppSDK 1.8)
└── README.md
```

## Test 1: Uri (Windows SDK only)

No WinAppSDK needed. Works on any Windows 10/11 with Windows SDK.

```bash
# Generate
./winrt-meta.exe generate --namespace "Windows.Foundation" --class-name "Uri" --output ./generated-uri

# Run
npm install
npx tsx test_uri.ts
```

## Test 2: FileOpenPicker (WinAppSDK)

Requires WinAppSDK 1.8 runtime and the `WINAPPSDK_BOOTSTRAP_DLL_PATH` environment variable.

```bash
# Generate
./winrt-meta.exe generate \
  --winmd "<path-to>/Microsoft.Windows.Storage.Pickers.winmd" \
  --namespace "Microsoft.Windows.Storage.Pickers" \
  --class-name "FileOpenPicker" \
  --output ./generated

# Run
npx tsx test_picker.ts
```

A file picker dialog will open. Select a file to complete the test.

## Prerequisites

- Windows 10/11 with Windows SDK installed
- Node.js 20+
- For Test 2 only: WinAppSDK 1.8 runtime + `WINAPPSDK_BOOTSTRAP_DLL_PATH`

### WINAPPSDK_BOOTSTRAP_DLL_PATH

This environment variable points to the WinAppSDK Bootstrap DLL, which dynwinrt-js loads to initialize WinAppSDK APIs (e.g. `Microsoft.Windows.Storage.Pickers.FileOpenPicker`). Without it, `initWinappsdk()` will fail.

Typical value:
```
C:\Users\<user>\.winapp\packages\Microsoft.WindowsAppSDK.Foundation.1.8.251104000\runtimes\win-x64\native\Microsoft.WindowsAppRuntime.Bootstrap.dll
```

If you use `@microsoft/winappcli`, run `winapp restore` to download the WinAppSDK package, then set the variable:

```powershell
$env:WINAPPSDK_BOOTSTRAP_DLL_PATH = "$HOME\.winapp\packages\Microsoft.WindowsAppSDK.Foundation.1.8.251104000\runtimes\win-x64\native\Microsoft.WindowsAppRuntime.Bootstrap.dll"
```

Windows SDK APIs (Test 1) do not require this variable.

### Windows.winmd

`winrt-meta` auto-detects `Windows.winmd` from `C:\Program Files (x86)\Windows Kits\10\UnionMetadata\`. This file is installed with the Windows SDK and contains type definitions for all `Windows.*` APIs.
