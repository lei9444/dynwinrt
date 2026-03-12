# dynwinrt Sample

A self-contained test kit for dynwinrt-js. Download this folder from GitHub Actions artifacts and run.

## Contents

```
sample/
├── winrt-meta.exe      # codegen tool
├── dynwinrt-js/        # runtime (index.js + index.d.ts + .node files)
├── package.json
├── test_picker.ts      # test script (requires WinAppSDK 1.8)
└── README.md
```

## Quick Start

### 1. Generate bindings

```bash
./winrt-meta.exe generate \
  --winmd "<path-to>/Microsoft.Windows.Storage.Pickers.winmd" \
  --namespace "Microsoft.Windows.Storage.Pickers" \
  --class-name "FileOpenPicker" \
  --output ./generated
```

### 2. Install and run

```bash
npm install
npx tsx test_picker.ts
```

A file picker dialog will open. Select a file to complete the test.

## Prerequisites

- Windows 10/11
- Node.js 20+
- WinAppSDK 1.8 runtime
- `WINAPPSDK_BOOTSTRAP_DLL_PATH` environment variable set

## Using Windows SDK APIs (no WinAppSDK needed)

For Windows SDK APIs like Uri, HttpClient, etc., no WinAppSDK is needed:

```bash
./winrt-meta.exe generate \
  --namespace "Windows.Foundation" \
  --class-name "Uri" \
  --output ./generated-uri
```

Windows.winmd is auto-detected from the Windows SDK install path.
