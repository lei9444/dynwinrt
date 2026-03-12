# dynwinrt SDK Release TODO

## P0 - Release Blockers

- [ ] **JS binding error handling**: `bindings/js/src/lib.rs` still has `.unwrap()` calls that crash the Node.js process; replace remaining ones with `napi::Result` + `.map_err()` (some already fixed: `callSingleOut0/1`, `call_0`)
- [ ] **Package metadata**: All Cargo.toml files missing `authors`, `license`, `description`, `repository`, `keywords`
  - `bindings/js/package.json` repository URL points to napi-rs template, needs update
  - `bindings/py/pyproject.toml` missing `authors`, `license`, `homepage`
- [ ] **CI/CD**: No GitHub Actions; add `.github/workflows/test.yml` (cargo test + npm test + pytest)
- [ ] **Remove debug eprintln**: `meta.rs` has `[resolve]` debug prints that should be removed before release

## P1 - Quality

- [ ] **Clippy cleanup**: 74+ warnings (dead code, unused imports, style issues)
  - `strip_generic_arity()` (winrt-meta) never used
  - `query_interface()`, `find_winappsdk_package()` unused
  - Redundant closures (`.map(|a| f(a))` -> `.map(f)`)
- [ ] **Update CLAUDE.md**: Known Limitations section outdated -- generics fully supported, codegen tool exists, parameterized interfaces from winmd
- [ ] **Python .pyi type stubs**: No Python type hint files generated
- [ ] **JSDoc comments**: napi binding `.d.ts` has no parameter descriptions
- [ ] **Remove unused `_collections.ts`**: Now that parameterized interfaces are generated from winmd (IVector_String.ts etc.), the hardcoded `_collections.ts` fallback can be removed
- [ ] **Remove unused JS binding methods**: `call_0`, `callSingleOut0`, `callSingleOut1` are superseded by `method().invoke()` — deprecate or remove

## P2 - Feature Completeness

- [ ] **Delegate / Event support**: TypeKind::Delegate exists but cannot create callbacks from JS/Python
  - Needs: Rust-side COM vtable implementation + napi ThreadsafeFunction callback
  - Blocks: event subscription (IObservableVector.VectorChanged, etc.), all methods accepting delegate parameters
- [ ] **Struct auto-marshaling**: Users must manually `DynWinRtStruct.create()` + `setF64(index, value)` per field; support auto-conversion from JS objects
- [ ] **IAsyncOperationWithProgress IID computation**: Named struct signature issue (anonymous structs + lost Parameterized types -> wrong IID -> QI fails)
  - Needs: `StructEntry.name: Option<String>`, `enum(Name;i4)` signature, preserve Parameterized in struct fields
- [ ] **Nullable / IReference\<T\> return handling**: Null COM pointer returns `Null` variant; JS side needs better null-check patterns
- [ ] **Struct codegen deduplication**: `DynWinRtType.registerStruct(...)` is inlined in every method signature that uses the struct; should generate a shared struct definition file and import it (runtime is idempotent, but codegen is verbose)
- [ ] **Exclusive interface codegen**: Methods on exclusive interfaces (e.g. `IXmlDocumentIO.LoadXml`) are not generated; need to resolve all interfaces a class implements, not just the default one

## P3 - Developer Experience

- [ ] **Error message enrichment**: COM HRESULT errors should include WinRT error message (`IRestrictedErrorInfo`)
- [ ] **Performance**:
  - `call()` / `callVoid()` create a temporary InterfaceSignature + build Method per call; should cache or remove in favor of `invoke()`
- [ ] **Multi-platform builds**:
  - npm prebuild support (currently only win32-x64-msvc)
  - ARM64 Windows validation
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
