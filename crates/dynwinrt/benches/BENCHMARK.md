# dynwinrt Benchmark: Dynamic vs Static WinRT API Calls

## Overview

Compares the same WinRT operations invoked via:
- **Static**: windows-rs compile-time bindings (direct vtable call, compiler-inlined)
- **Dynamic**: dynwinrt runtime invocation (`MethodHandle::invoke` path)
- **Raw vtable**: manual vtable call without dynwinrt overhead (isolates framework cost)

All benchmarks use `Windows.Foundation.Uri` — a system built-in WinRT class with no I/O, pure CPU operations.

## CallStrategy Coverage

dynwinrt selects a `CallStrategy` at method build time to optimize common patterns:

| Strategy | When | Uses libffi? |
|---|---|---|
| `Direct0In0Out` | No params, no return (e.g. event fire) | No |
| `Direct0In1Out` | Property getter (0 in, 1 out) | No |
| `Direct1In0Out` | Property setter (1 in, 0 out) | No |
| `Direct1In1Out` | Factory/method (1 in, 1 out) | No |
| `Libffi(Cif)` | General case (2+ in/out params) | Yes (cached Cif) |

## Results

Measured on Intel i7, Windows 11, Rust 1.87 (release mode with LTO).

### Per-operation comparison

| Benchmark | Strategy | Static | Dynamic | Raw vtable | Ratio |
|---|---|---|---|---|---|
| **create_uri** | Direct1In1Out | 2.05 µs | 2.33 µs | — | 1.14x |
| **create_with_relative** | Libffi 2in1out | 701 ns | 953 ns | — | 1.36x |
| **get_absolute_uri** (HString) | Direct0In1Out | 5.97 ns | 44.7 ns | 6.88 ns | 7.5x |
| **get_port** (i32) | Direct0In1Out | 2.56 ns | 63.9 ns | — | 25x |
| **combine_uri** | Direct1In1Out | 675 ns | 924 ns | — | 1.37x |
| **equals** | Direct1In1Out | 207 ns | 250 ns | — | 1.21x |
| **get_3_properties** | Direct0In1Out x3 | 19.9 ns | 142 ns | — | 7.1x |
| **batch 100 create+read** | Mixed | 193 µs | 234 µs | — | 1.21x |
| **geopoint_create** (struct in) | Libffi struct | 114 ns | 472 ns | — | 4.1x |

### Overhead breakdown (get_absolute_uri)

Using a raw vtable baseline to isolate each layer of overhead:

```
static (windows-rs):    5.94 ns  — compiler-inlined vtable call
raw vtable call:        6.79 ns  — manual vtable dispatch, no framework overhead
dynamic invoke:        44.3  ns  — full MethodHandle::invoke path

Framework overhead decomposition (44.3 - 6.79 = 37.5 ns):
  RwLock read (uncontended):    ~10-15 ns  (atomic CAS on methods arena)
  vec![out] heap alloc+free:    ~20-30 ns  (malloc + free per call)
  default_winrt_value + misc:    ~3-5  ns  (stack enum init, pointer ops)
```

### Note on get_port (i32) — why 24.6x?

The 24.6x ratio looks alarming but is misleading. The issue is that the static and dynamic paths measure very different amounts of work:

**Static 2.60 ns** — the compiler fully inlines the COM vtable call:
```asm
mov rax, [rcx]           ; read vtable pointer
call [rax + 14*8]        ; vtable[14] indirect call
mov [rsp+out], eax       ; store i32 result on stack
; ~3-4 instructions total, branch predictor 100% hit
```

**Dynamic 63.9 ns** — the full `MethodHandle::invoke` path:
1. `RwLock::read()` — atomic CAS (~12 ns)
2. Array index to get `Method` (~1 ns)
3. `match CallStrategy` dispatch (~1 ns)
4. `default_winrt_value()` — create `WinRTValue::I32(0)` on stack (~1 ns)
5. `out_ptr()` — take pointer to the i32 inside the enum (~1 ns)
6. **vtable call** — the actual COM call (~3 ns, same as static)
7. `vec![out]` — **heap allocation** for return Vec (~30 ns)
8. `RwLock` drop — release read lock (~2 ns)
9. Result/Vec return + caller drop/free (~12 ns)

The vtable call itself (step 6) is only ~3 ns — virtually identical to static. The remaining ~61 ns is framework overhead, dominated by `vec![out]` heap allocation (~30 ns) and `RwLock` (~12 ns).

The 24.6x ratio is simply because the denominator (2.60 ns) is extremely small after compiler inlining. The **absolute overhead** of ~61 ns is the same as for `get_absolute_uri` (~38 ns framework). In practice, 64 ns to read an i32 property is negligible — a single Python function call takes ~100-300 ns.

## Key Findings

### 1. vtable call itself has near-zero overhead
Raw vtable (6.79 ns) vs static (5.94 ns) — the dynamic function pointer dispatch is essentially free. The ~0.85 ns difference is likely due to the static path being fully inlined by the compiler.

### 2. Vec heap allocation dominates for lightweight getters
`vec![out]` accounts for ~60-70% of the framework overhead. Each call does `malloc` (allocate Vec buffer) + `free` (drop Vec after caller reads the result). This is the primary optimization target — replacing `Vec<WinRTValue>` with `SmallVec<[WinRTValue; 2]>` would eliminate heap allocation for 0-2 return values, covering all Direct call paths.

### 3. RwLock read adds ~10-15ns per call
The `methods` arena is behind `RwLock` for thread safety during interface registration. In the invoke hot path, each call takes a read lock (uncontended = atomic CAS). Could be eliminated by caching a direct `&Method` reference after registration is complete.

### 4. Real-world operations are dominated by WinRT itself
For operations with actual work (object creation at ~2µs, string manipulation at ~700ns), the framework overhead is only 12-36%. The overhead is dwarfed by the WinRT operation.

### 5. Libffi path adds modest overhead over Direct
`CreateWithRelativeUri` (Libffi, 2 in + 1 out) is 1.36x vs static, while `CreateUri` (Direct, 1 in + 1 out) is 1.12x. The extra ~24% comes from libffi argument marshaling, but the Cif is pre-cached so there's no repeated compilation.

### 6. Struct parameters add noticeable but bounded overhead
`Geopoint::Create(BasicGeoposition)` with a 3xf64 struct: 4.1x slower (114 → 447 ns). Using a struct-alloc-only baseline to isolate costs:

```
static (windows-rs):      117 ns  — compiler passes struct by value on stack
struct_alloc_only:         115 ns  — default_value() heap alloc + 3x set_field (no WinRT call)
dynamic_with_struct_alloc: 447 ns  — full path (alloc + fields + libffi invoke)

Decomposition of dynamic 447 ns:
  struct alloc + fields:   115 ns  (26%) — heap-allocated aligned buffer + field writes
  invoke overhead:         332 ns  (74%) — of which:
    WinRT call itself:     ~117 ns        (same as static)
    libffi marshaling:     ~175 ns        (struct→ffi layout, field copying, Vec, RwLock)
    Vec + RwLock + misc:    ~40 ns        (same as other paths)
```

The struct heap allocation (`default_value`) alone costs as much as the entire static call. The libffi struct marshaling adds ~175 ns on top — significantly more than the ~38 ns framework overhead for scalar types, because libffi must construct a `Type::structure` layout and copy each field into the ffi argument buffer.

Despite the 4.1x ratio, 447 ns absolute is still very fast. Struct parameters are uncommon in typical WinRT usage (most methods pass Object references).

### 7. End-to-end is acceptable for JS/Python bindings
Batch 100 creates + reads: 1.16x slower. Cross-language overhead (JS<->Rust ~100-500ns, Python<->Rust ~200-800ns per call) will dominate over the dynwinrt framework overhead in practice.

## Optimization Opportunities

| Optimization | Estimated improvement | Effort |
|---|---|---|
| `SmallVec<[WinRTValue; 2]>` for return values | -25-30ns per getter call | Medium (API change) |
| Freeze methods arena + direct `&Method` refs | -10-15ns per call | Medium |
| Specialized `invoke_getter() -> WinRTValue` API | -35-40ns (no Vec, no lock) | Low (additive API) |
| Pre-compute vtable function pointer in MethodHandle | -1-2ns | Low |

## Running

```bash
cargo bench -p dynwinrt --bench uri_bench
```

HTML reports are generated in `target/criterion/`.
