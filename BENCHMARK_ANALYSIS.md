# Performance Analysis: N-API Layer Overhead

Deep dive into where time is spent across different N-API binding strategies.

Environment: Snapdragon X Elite, Windows 11, Node.js v22. All Rust builds: release + LTO + codegen-units=1.

---

## 4-Way Static Comparison

Four JS → native → WinRT paths benchmarked side by side (50,000 iterations):

| Path | Stack |
|------|-------|
| C++ static | JS → node-addon-api → C++/WinRT → COM vtable |
| Rust raw N-API | JS → napi-sys (raw) → windows crate → COM vtable |
| Rust napi-rs | JS → napi-rs (macro) → windows crate → COM vtable |
| Dynamic (standard) | JS → napi-rs → dynwinrt (Rust + libffi) → COM vtable |

### Getter (0 in → 1 out)

| API | C++ static | Rust raw | Rust napi-rs | Dynamic |
|-----|-----------|----------|-------------|---------|
| `get_Host()` → hstring | 70 ns | 77 ns | 205 ns | 1.10 µs |
| `get_Port()` → i32 | 40 ns | 48 ns | 69 ns | 1.04 µs |
| `get_Suspicious()` → bool | 35 ns | 47 ns | 69 ns | - |

### Factory (1 in → 1 out object)

| API | C++ static | Rust raw | Rust napi-rs | Dynamic |
|-----|-----------|----------|-------------|---------|
| `Uri.CreateUri(hstring)` | 897 ns | 945 ns | 1.56 µs | 3.16 µs |
| `PV.CreateInt32(i32)` | 429 ns | 705 ns | 1.28 µs | 2.31 µs |
| `PV.CreateDouble(f64)` | 325 ns | 732 ns | 1.77 µs | 3.06 µs |
| `PV.CreateBoolean(bool)` | 648 ns | 296 ns | 1.30 µs | 2.52 µs |
| `PV.CreateString(hstring)` | 709 ns | 896 ns | 1.10 µs | 3.09 µs |
| `Geopoint.Create(struct 3×f64)` | 890 ns | 921 ns | 2.09 µs | 5.82 µs |

### Ratio vs C++ static

| Layer | Getter | Factory | Struct |
|-------|--------|---------|--------|
| C++ static | 1.0x | 1.0x | 1.0x |
| Rust raw N-API | **1.0-1.3x** | 0.5-1.6x | 1.0x |
| Rust napi-rs | 1.7-2.9x | 1.6-2.4x | 2.3x |
| Dynamic (standard) | 15-26x | 3.5-5.3x | 6.5x |

---

## Fast Path Optimization

Specialized methods that combine invoke + result extraction into a single napi call, using zero-allocation direct vtable calls.

### What each level eliminates

| Level | Description | Savings |
|-------|-------------|---------|
| Standard `invoke([]).toString()` | 2-3 napi calls, Vec alloc, WinRTValue wrapping | baseline |
| **Fast path v2 (zero-alloc)** | 1 napi call, `call_getter_i32()` writes to stack | skip Vec + WinRTValue + extra napi call |
| **Raw napi-sys + dynwinrt** | 0 napi-rs overhead, raw `napi_unwrap` + direct call | skip napi-rs macro layer entirely |

### Results

| Operation | Standard | Fast path (napi-rs) | Raw napi-sys + dynwinrt | C++ static |
|-----------|----------|--------------------|-----------------------|-----------|
| `get_Host()` → string | 1.37 µs | 255 ns | **68 ns** | 60 ns |
| `get_Port()` → i32 | 1.18 µs | 124 ns | **48 ns** | 38 ns |
| `CreateUri(hstring)` → obj | 3.11 µs | 2.03 µs | - | 897 ns |
| `PV.CreateInt32(i32)` → obj | 2.41 µs | 1.27 µs | - | 429 ns |

### Speedup summary

| Optimization | get_Host | get_Port |
|-------------|---------|---------|
| Standard → napi-rs fast path | **5.4x** | **9.5x** |
| Standard → raw napi-sys + dynwinrt | **20x** | **25x** |
| Raw napi-sys + dynwinrt vs C++ | **1.1x** | **1.3x** |

**Raw napi-sys + dynwinrt dynamic dispatch is within 10-30% of C++ static projection.**

### Fast Path API Reference

Methods on `DynWinRtMethodHandle` (napi-rs binding):

| Method | JS Signature | Returns | Rust fast path |
|--------|-------------|---------|----------------|
| `getString(obj)` | `(obj: DynWinRtValue) → string` | JS string directly | `call_getter_hstring()` → HSTRING on stack |
| `getI32(obj)` | `(obj: DynWinRtValue) → number` | JS number directly | `call_getter_i32()` → i32 on stack |
| `getBool(obj)` | `(obj: DynWinRtValue) → boolean` | JS boolean directly | `call_getter_bool()` → bool on stack |
| `getObj(obj)` | `(obj: DynWinRtValue) → DynWinRtValue` | Wrapped COM object | `call_getter_object()` → raw ptr on stack |
| `invokeHstring(obj, s)` | `(obj: DynWinRtValue, s: string) → DynWinRtValue` | Wrapped COM object | Skips `DynWinRtValue.hstring()` napi call |
| `invokeI32(obj, n)` | `(obj: DynWinRtValue, n: number) → DynWinRtValue` | Wrapped COM object | Skips `DynWinRtValue.i32()` napi call |

Usage in generated code:
```typescript
// Standard (2-3 napi calls, ~1.2 µs)
const host = method.invoke(obj, []).toString();

// Fast path (1 napi call, ~255 ns)
const host = method.getString(obj);
```

---

## Electron IPC Round-trip

Every call: **Renderer → IPC → Main Process → native addon → WinRT → IPC → Renderer**.

IPC baseline (noop): **~79 µs**.

| API | Static (C++/WinRT) | Dynamic (dynwinrt) | Ratio |
|-----|--------------------|--------------------|-------|
| `Uri.get_Host()` → hstring | 83.8 µs | 82.2 µs | 1.0x |
| `Uri.get_Port()` → i32 | 83.9 µs | 83.2 µs | 1.0x |
| `Uri.CreateUri(hstring)` | 76.7 µs | 90.1 µs | 1.2x |
| `PV.CreateInt32(i32)` | 73.0 µs | 81.7 µs | 1.1x |
| `Geopoint.Create(3×f64 struct)` | 77.7 µs | 81.9 µs | 1.1x |

**IPC overhead (~80 µs) completely dominates everything else.**

---

## Overhead by Layer

| Layer | Per-call | Source |
|-------|----------|--------|
| COM vtable dispatch | <1 ns | Function pointer indirection |
| dynwinrt Rust core | ~10 ns | RwLock read + vtable ptr extraction |
| Raw N-API boundary | ~30-70 ns | `napi_get_value_external` + `napi_create_string_utf16` |
| napi-rs macro layer | +75-170 ns | Type checking, class unwrap, UTF-8 string roundtrip |
| DynWinRtValue wrapping | +100-200 ns | External creation, Vec allocation |
| Extra napi call (`.toString()`) | +500-1000 ns | Second napi boundary crossing for result extraction |
| Electron IPC | ~80 µs | Renderer ↔ Main process serialization |

### Complete call path breakdown (get_Port → i32)

```
38 ns   C++ static (node-addon-api + C++/WinRT)
48 ns   Rust raw N-API + dynwinrt          = 38 + 10 ns dynwinrt dispatch
124 ns  napi-rs fast path (getI32)         = 48 + 76 ns napi-rs macro overhead
1180 ns Standard (invoke + toNumber)       = 124 + 1056 ns (Vec + WinRTValue + 2nd napi call)
80 µs   Electron IPC round-trip            = 1180 ns + 79 µs IPC
```

---

## How to Run

```bash
# 4-way benchmark (C++ / Rust raw / Rust napi-rs / Dynamic)
cd bindings/js && npm run build
cd bindings/js/static-bench-cpp && npm install && npm run build
cargo build -p static-bench-rust --release
cp target/release/static_bench_rust.dll bindings/js/static-bench-rust/static_bench_rust.node
cd bindings/js && npx tsx samples/bench_3way.ts

# Electron IPC benchmark
cd bench-electron && npm install && npm run dev

# Rust core engine (no JS overhead)
cargo bench -p dynwinrt --bench bench
```

---

## Summary

```
                    get_Port → i32 (per-call latency)

C++ static            ██ 38 ns
Rust raw static       ██ 48 ns
Raw napi + dynwinrt   ██ 48 ns          ← dynamic dispatch, C++ parity!
napi-rs fast path     █████ 124 ns
napi-rs standard      █████████████████████████████████████████████████ 1180 ns
Electron IPC          ██████████████████████████████████████████ ... █████ 80,000 ns
```

### Key Insights

1. **Dynamic dispatch can match C++ static** — Raw napi-sys + dynwinrt: 48 ns vs C++ 38 ns for i32 getter. The 10 ns gap is purely dynwinrt's RwLock read.

2. **napi-rs macro layer = main bottleneck** — Adds 76-170 ns per call (type checking, UTF-8 string roundtrip, class unwrap). This is the #1 optimization target.

3. **Fast path API = 5-10x speedup** — `getString()`/`getI32()` eliminate the second napi call and Vec allocation. Available now, codegen can adopt immediately.

4. **Standard invoke penalty = extra napi calls** — Each `DynWinRtValue.hstring()` and `.toString()` is a full napi round-trip (~500-1000 ns). The "invoke with raw JS values" optimization (TODO) would eliminate these.

5. **Electron IPC dwarfs everything** — 80 µs makes all native-layer differences irrelevant from the UI.
