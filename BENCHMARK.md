# dynwinrt Benchmark Results

Static projection (windows-rs) vs dynamic invocation (dynwinrt).

Environment: Intel i7, Windows 11, Rust 1.87 release + LTO, Node.js v24.

---

## JavaScript End-to-End

Full round-trip: JS → napi → Rust → COM vtable → Rust → napi → JS.
All method handles and objects pre-created before measurement.

### By parameter count

Fixed return type, varying number of input parameters:

| Params | API | Static | Dynamic | Ratio | Overhead |
|--------|-----|--------|---------|-------|----------|
| 0 in → 1 out | `Uri.get_Host()` | 279 ns | 1.17 µs | 4.2x | +891 ns |
| 1 in → 1 out | `Uri.CombineUri(hstring)` | 2.41 µs | 4.02 µs | 1.7x | +1.61 µs |
| 2 in → 1 out | `Uri.CreateWithRelativeUri(hstring, hstring)` | 2.53 µs | 6.70 µs | 2.6x | +4.17 µs |

### By input type

Fixed call shape (1 in → 1 out object), varying input type:

| Input type | API | Static | Dynamic | Ratio | Overhead |
|------------|-----|--------|---------|-------|----------|
| i32 | `PropertyValue.CreateInt32(42)` | 797 ns | 2.63 µs | 3.3x | +1.83 µs |
| f64 | `PropertyValue.CreateDouble(3.14)` | 820 ns | 2.26 µs | 2.8x | +1.44 µs |
| bool | `PropertyValue.CreateBoolean(true)` | 1.17 µs | 2.58 µs | 2.2x | +1.41 µs |
| hstring | `PropertyValue.CreateString("hello")` | 2.49 µs | 2.69 µs | 1.1x | +200 ns |
| struct (3×f64) | `Geopoint.Create(BasicGeoposition)` | 2.42 µs | 3.97 µs | 1.6x | +1.55 µs |

### By return type

Fixed call shape (0 in → 1 out getter on pre-created Uri), varying return type.
Dynamic includes `.toNumber()` / `.toBool()` / `.toString()` to match static's direct JS value return.

| Return type | API | Static | Dynamic | Ratio | Overhead |
|-------------|-----|--------|---------|-------|----------|
| i32 | `Uri.get_Port()` | 92 ns | 1.74 µs | 18.9x | +1.65 µs |
| bool | `Uri.get_Suspicious()` | 93 ns | 1.96 µs | 21.1x | +1.87 µs |
| hstring | `Uri.get_Host()` | 270 ns | 1.31 µs | 4.9x | +1.04 µs |

High ratios on i32/bool are because static getters are extremely fast (~93ns). Absolute overhead is consistent (~1-2µs) across all return types.

### By parameter count (with cached args)

Same as above, but args pre-created outside the loop — isolates pure invoke overhead:

| Params | API | Static | Dynamic (cached) | Ratio | Overhead |
|--------|-----|--------|-------------------|-------|----------|
| 0 in → 1 out | `Uri.get_Host()` | 279 ns | 1.17 µs | 4.2x | +891 ns |
| 1 in → 1 out | `Uri.CombineUri(hstring)` | 2.41 µs | 3.01 µs | 1.2x | +600 ns |
| 2 in → 1 out | `Uri.CreateWithRelativeUri(hstring, hstring)` | 2.53 µs | 3.07 µs | 1.2x | +540 ns |

With cached args, overhead drops to **~600ns fixed** regardless of param count. The extra overhead in the uncached group comes from `DynWinRtValue` construction crossing the napi boundary (~1µs per arg).

### Arg caching (napi call savings)

Same `PropertyValue.Create*(v)`, comparing cached arg (1 napi call) vs uncached (2 napi calls):

| Type | Static | Cached (1 napi) | Uncached (2 napi) | Cached ratio | Uncached ratio |
|------|--------|----------------|-------------------|--------------|----------------|
| i32 | 797 ns | 1.31 µs | 2.73 µs | 1.6x | 3.4x |
| f64 | 820 ns | 1.63 µs | 2.60 µs | 2.0x | 3.2x |
| bool | 1.17 µs | 1.58 µs | 2.38 µs | 1.4x | 2.0x |
| hstring | 2.49 µs | 1.93 µs | 3.02 µs | ~1x | 1.2x |

Note: JS benchmarks have ~30% run-to-run variance due to V8 JIT/GC. The key finding is consistent: **caching saves ~0.5-1.4µs per argument** (one fewer napi boundary crossing). If `invoke()` could accept raw JS values directly (planned optimization), this saving would be automatic.

### Method handle caching

| Approach | Time |
|----------|------|
| Cached | 1.34 µs |
| Uncached (lookup by name each call) | 2.48 µs |
| **Savings** | **~1.1 µs/call** |

### Batch workload

| Workload | Static | Dynamic | Ratio | Overhead |
|----------|--------|---------|-------|----------|
| 200× create Uri + read Host | 430 µs | 1.66 ms | 3.9x | +1.23 ms |

### How to run

```bash
cd bindings/js && npm run build && npx tsx samples/benchmark.ts
```

---

## Rust Core Engine

Same WinRT operations, measured in pure Rust with criterion.rs.

### By parameter count

| Params | API | Static | Dynamic | Ratio | Overhead |
|--------|-----|--------|---------|-------|----------|
| 0 in → 1 out | `Uri.get_Host()` | 6.7 ns | 65 ns | 9.7x | +58 ns |
| 1 in → 1 out | `Uri.CombineUri(hstring)` | 2.5 µs | 2.8 µs | 1.1x | +300 ns |
| 2 in → 1 out | `Uri.CreateWithRelativeUri(hstring, hstring)` | 684 ns | 953 ns | 1.4x | +269 ns |

### By input type

| Input type | API | Static | Dynamic | Ratio | Overhead |
|------------|-----|--------|---------|-------|----------|
| i32 | `PropertyValue.CreateInt32(42)` | 105 ns | 155 ns | 1.5x | +50 ns |
| f64 | `PropertyValue.CreateDouble(3.14)` | 105 ns | 152 ns | 1.4x | +47 ns |
| bool | `PropertyValue.CreateBoolean(true)` | 103 ns | 151 ns | 1.5x | +48 ns |
| hstring | `PropertyValue.CreateString("hello")` | 107 ns | 322 ns | 3.0x | +215 ns |
| object | `PropertyValue.CreateInspectable(obj)` | 7.1 ns | 57 ns | 8.0x | +50 ns |
| struct (3×f64) | `Geopoint.Create(BasicGeoposition)` | 116 ns | 618 ns | 5.3x | +502 ns |

### By return type

| Return type | API | Static | Dynamic | Ratio | Overhead |
|-------------|-----|--------|---------|-------|----------|
| i32 | `Uri.get_Port()` | 2.8 ns | 50 ns | 17.9x | +47 ns |
| bool | `Uri.get_Suspicious()` | 3.6 ns | 135 ns | 37.5x | +131 ns |
| hstring | `Uri.get_Host()` | 7.6 ns | 72 ns | 9.5x | +64 ns |
| object | `Uri.CombineUri(hstring)` | 735 ns | 1.04 µs | 1.4x | +303 ns |

### Overhead isolation

| Path | hstring getter | i32 getter |
|------|---------------|------------|
| Static (compiler-inlined) | 6.6 ns | 2.8 ns |
| Raw vtable (no framework) | 7.4 ns | 24 ns |
| Dynamic (dynwinrt) | 51 ns | 72 ns |
| **Framework overhead** | **+43 ns** | **+48 ns** |

### Batch workload

| Workload | Static | Dynamic | Ratio | Overhead |
|----------|--------|---------|-------|----------|
| 100× create Uri + read property | 225 µs | 241 µs | 1.07x | +16 µs |

### How to run

```bash
cargo bench -p dynwinrt --bench bench
```

---

## Overhead by layer

| Layer | Per-call overhead | Source |
|-------|-------------------|--------|
| COM vtable dispatch | <1 ns | Function pointer indirection |
| dynwinrt Rust core | ~50 ns | RwLock + Vec alloc + value marshaling |
| JS ↔ Rust napi boundary | ~500-1000 ns | V8 value conversion, argument wrapping |
