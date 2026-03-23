# dynwinrt SDK Release TODO

## P0 - Release Blockers

- [x] **JS binding error handling**: All 13 `.unwrap()` calls in `bindings/js/src/lib.rs` replaced with `napi::Result` + `.map_err()` / `.ok_or_else()` — errors now surface as JS exceptions instead of crashing the Node.js process
- [ ] **Package metadata**: All Cargo.toml files missing `authors`, `license`, `description`, `repository`, `keywords`
  - `bindings/js/package.json` repository URL points to napi-rs template, needs update
  - `bindings/py/pyproject.toml` missing `authors`, `license`, `homepage`
- [x] **CI/CD**: `.github/workflows/build.yml` — winrt-meta builds (x64 + arm64), dynwinrt-js (x64 + arm64), publishing, and sample generation
- [x] **Remove debug eprintln**: `[resolve]` debug prints removed from `meta.rs`
- [ ] **Auto-detect WinAppSDK Bootstrap DLL**: `initWinappsdk(major, minor)` should auto-find Bootstrap DLL from `~/.winapp/packages/` or known install paths, with `WINAPPSDK_BOOTSTRAP_DLL_PATH` as override. Currently requires manual env var setup which is a friction point for unpackaged app developers.

## P1 - Quality

- [ ] **Clippy cleanup**: 74+ warnings (dead code, unused imports, style issues)
  - `strip_generic_arity()` (winrt-meta) never used
  - `query_interface()`, `find_winappsdk_package()` unused
  - Redundant closures (`.map(|a| f(a))` -> `.map(f)`)
- [ ] **Update CLAUDE.md**: Known Limitations section outdated -- generics fully supported, codegen tool exists, parameterized interfaces from winmd
- [ ] **Python .pyi type stubs**: No Python type hint files generated
- [ ] **JSDoc comments**: napi binding `.d.ts` has no parameter descriptions
- [x] **Remove unused `_collections.ts`**: Removed — parameterized interfaces now generated from winmd (IVector_String.ts etc.)
- [x] **Remove unused JS binding methods**: `call_0`, `callSingleOut0`, `callSingleOut1` removed — `method().invoke()` is the sole invoke path

## P2 - Feature Completeness

- [x] **Delegate / Event support**: Full implementation in `delegate.rs` — Rust-side COM vtable + napi ThreadsafeFunction callback. `DynWinRtDelegate.create(iid, paramTypes, callback)` creates delegate COM objects from JS callbacks. Supports Object, HString, Bool, I32/U32/I64/U64, Enum parameter types
- [ ] **Struct auto-marshaling**: Users must manually `DynWinRtStruct.create()` + `setF64(index, value)` per field; support auto-conversion from JS objects
- [ ] **IAsyncOperationWithProgress IID computation**: Struct fields containing enums produce `i4` instead of `enum(Name;i4)` in type signature → wrong IID → QI fails
  - Root cause: enum fields in struct signature not using named format
  - Also: `StructEntry.name` uses `Option<String>` but WinRT structs are always named — should be `String`, deprecate `define_struct` in favor of `define_named_struct`
- [ ] **Nullable / IReference\<T\> return handling**: Null COM pointer returns `Null` variant; JS side needs better null-check patterns
- [ ] **Struct codegen deduplication**: `DynWinRtType.registerStruct(...)` is inlined in every method signature that uses the struct; should generate a shared struct definition file and import it (runtime is idempotent, but codegen is verbose)
- [ ] **Exclusive interface codegen**: Methods on exclusive interfaces (e.g. `IXmlDocumentIO.LoadXml`) are not generated; need to resolve all interfaces a class implements, not just the default one
- [x] **Codegen missing dependency warning**: `resolve_named_type` now emits warnings when types are not found in loaded .winmd files, plus `assert!(!iid.is_empty(), ...)` catches empty GUIDs at generation time

## P3 - Developer Experience

- [ ] **Error message enrichment**: COM HRESULT errors should include WinRT error message (`IRestrictedErrorInfo`)
- [ ] **Performance**:
  - `call()` / `callVoid()` create a temporary InterfaceSignature + build Method per call; should cache or remove in favor of `invoke()`
  - `invoke()` should accept raw JS values (number, string, bool) instead of requiring `DynWinRtValue` wrappers — saves ~0.6-1.6µs per argument (one fewer napi boundary crossing). Needs `in_param_types()` on MethodHandle + type-directed conversion in `bindings/js/src/lib.rs`
  - Rust core: `invoke_method` takes RwLock read on every call (~15-20 ns); store `Arc<Method>` in MethodHandle directly to bypass lock
  - Rust core: `Ok(vec![out])` heap-allocates per call; switch to `SmallVec<[WinRTValue; 2]>` for stack return
- [x] **Multi-platform builds**: npm prebuild for `win32-x64-msvc` and `aarch64-pc-windows-msvc` (ARM64). CI builds both architectures via `.github/workflows/build.yml`
- [ ] **Python binding parity**: Python binding missing `callVoid`, collection wrappers, struct access, and other APIs added to JS
- [ ] **Troubleshooting docs**: README missing common error resolution (WinAppSDK init failure, WINAPPSDK_BOOTSTRAP_DLL_PATH not set, etc.)

## Done (this session)

- [x] **IUnknown::from_raw(null) UB fix**: COM pointer out-params now use `RawPtr(*mut c_void)` instead of `IUnknown::from_raw(null)` which was UB under release optimization
- [x] **Parameterized type panic fix**: `default_winrt_value` for Parameterized types no longer panics
- [x] **JS binding type coverage**: Added all DynWinRtValue constructors (bool, i8-u64, f32, f64, guid, null), extractors (toBool, toI64, toF64, toGuid, isNull), and DynWinRtType factories (guid, char16, hresult, delegate, fillArray, iid)
- [x] **toNumber() expanded**: Now supports Bool, I8, U8, I16, U16, I32, U32, HResult
- [x] **Collection methods**: IVector/IVectorView/IMap/IMapView/IKeyValuePair/IIterable/IIterator with full methods
- [x] **Collection codegen from winmd**: Parameterized interfaces (IVector\<String\>, IReference\<UInt32\>) read from winmd and generated as concrete types (IVector_String.ts)
- [x] **Auto value wrapping**: `filter.append('.png')` works directly — generated IVector_String accepts `string` not `DynWinRtValue`
- [x] **Auto-detect Windows SDK**: winrt-meta automatically finds `Windows.winmd` from `C:\Program Files (x86)\Windows Kits\10\UnionMetadata\`
- [x] **callVoid()**: Added for void WinRT method calls
- [x] **DynWinRtType.iid()**: Compute parameterized IID from JS
- [x] **WinGuid.toString()**: For cache keys
- [x] **JS binding error handling**: All 13 `.unwrap()` in `bindings/js/src/lib.rs` → `napi::Result` with contextual error messages
- [x] **Electron benchmark app**: `bench-electron/` — full IPC round-trip benchmark (renderer → main → WinRT → main → renderer), static vs dynamic comparison. IPC baseline ~80µs dominates, ratio ~1.0x across all operations
- [x] **Enum as independent runtime type**: `WinRTValue::Enum { value, type_handle }` fully implemented across value.rs, call.rs, delegate.rs, array.rs, type_handle.rs, and JS bindings
- [x] **CI/CD**: `.github/workflows/build.yml` — winrt-meta builds (x64 + arm64), dynwinrt-js (x64 + arm64), publishing, and sample generation
- [x] **Remove debug eprintln**: `[resolve]` debug prints removed from `meta.rs`
- [x] **Remove unused `_collections.ts`**: Parameterized interfaces now generated from winmd
- [x] **Remove unused JS binding methods**: `call_0`, `callSingleOut0`, `callSingleOut1` removed
- [x] **Delegate / Event support**: Full `delegate.rs` — COM vtable + napi ThreadsafeFunction. `DynWinRtDelegate.create()` creates delegate COM objects from JS callbacks
- [x] **Codegen missing dependency warning**: Warnings emitted for missing types + assert on empty GUIDs
- [x] **Multi-platform builds**: ARM64 (`aarch64-pc-windows-msvc`) support in npm prebuild and CI
