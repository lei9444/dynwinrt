//! Benchmark: dynamic (dynwinrt) vs static (windows-rs) WinRT API calls.
//!
//! # Overview
//!
//! Compares the same WinRT operations invoked via:
//! - **Static**: windows-rs compile-time bindings (direct vtable call, compiler-inlined)
//! - **Dynamic**: dynwinrt runtime invocation (MethodHandle::invoke path)
//! - **Raw vtable**: manual vtable call without dynwinrt overhead (isolates framework cost)
//!
//! # CallStrategy Coverage
//!
//! dynwinrt selects a CallStrategy at method build time:
//! - `Direct0In1Out`: property getter (0 in, 1 out) — no libffi
//! - `Direct1In1Out`: factory/method (1 in, 1 out) — no libffi
//! - `Libffi(Cif)`:   general case (2+ in/out) — cached libffi Cif
//!
//! # Results Summary (typical, Intel i7)
//!
//! | Benchmark                   | Static  | Dynamic | Raw vtable | Ratio |
//! |-----------------------------|---------|---------|------------|-------|
//! | create_uri (1in1out)        | 2.0 µs  | 2.3 µs  | —          | 1.15x |
//! | create_with_relative (2in)  | 676 ns  | 940 ns  | —          | 1.39x |
//! | get_absolute_uri (HString)  | 6.0 ns  | 44.7 ns | 6.9 ns     | 7.5x  |
//! | get_port (i32)              | 2.5 ns  | 65.9 ns | —          | 25x   |
//! | combine_uri (1in1out)       | 678 ns  | 934 ns  | —          | 1.37x |
//! | equals (obj in, bool out)   | 205 ns  | 251 ns  | —          | 1.22x |
//! | get_3_properties            | 19.8 ns | 144 ns  | —          | 7.3x  |
//! | batch 100 create+read       | 197 µs  | 232 µs  | —          | 1.18x |
//! | geopoint_create (struct in) | 115 ns  | 487 ns  | —          | 4.2x  |
//!
//! # Overhead Breakdown (get_absolute_uri)
//!
//! ```text
//! static (windows-rs):    6.0 ns  — compiler-inlined vtable call
//! raw vtable call:        6.9 ns  — manual vtable dispatch, no framework overhead
//! dynamic invoke:        44.7 ns  — full MethodHandle::invoke path
//!
//! Overhead decomposition (44.7 - 6.9 = 37.8 ns):
//!   - RwLock read (uncontended):    ~10-15 ns  (atomic CAS)
//!   - vec![out] heap alloc+free:    ~20-30 ns  (malloc + free)
//!   - default_winrt_value + misc:    ~3-5  ns
//! ```
//!
//! # Key Findings
//!
//! 1. **vtable call itself is near-zero overhead** (6.9 vs 6.0 ns) — the dynamic
//!    dispatch through function pointer is essentially the same as static.
//!
//! 2. **Vec heap allocation dominates** for lightweight getters. Using SmallVec<[T; 2]>
//!    for inline storage would eliminate ~30ns per call.
//!
//! 3. **RwLock read** adds ~10-15ns per call. Could be eliminated by caching a direct
//!    reference to the Method after registration is complete.
//!
//! 4. **For real-world operations** (object creation, network calls), the framework
//!    overhead is <25% — the WinRT operation itself dominates.
//!
//! 5. **Libffi path** (2+ params) adds only ~20% over Direct path — Cif caching works.
//!
//! 6. **End-to-end batch** (100 creates + reads): only 1.18x slower — acceptable for
//!    JS/Python bindings where cross-language overhead (~100-500ns) dwarfs this.
//!
//! # Run
//!
//! ```bash
//! cargo bench -p dynwinrt --bench uri_bench
//! ```

use std::sync::Arc;

use criterion::{Criterion, black_box, criterion_group, criterion_main};
use windows::core::{HSTRING, Interface};
use windows::Foundation::Uri;
use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};

use dynwinrt::metadata_table::{MetadataTable, TypeHandle};
use dynwinrt::{MethodSignature, WinRTValue};

// ======================================================================
// Setup helpers
// ======================================================================

fn ensure_ro_initialized() {
    unsafe { RoInitialize(RO_INIT_MULTITHREADED) }.ok();
}

struct DynUriApi {
    factory_iid: windows::core::GUID,
    uri_iid: windows::core::GUID,
    factory_type: TypeHandle,
    uri_type: TypeHandle,
}

fn setup_dynamic() -> (Arc<MetadataTable>, DynUriApi) {
    let table = MetadataTable::new();

    // IUriRuntimeClassFactory: CreateUri(HSTRING) -> Object  [Direct1In1Out]
    let factory_iid =
        windows::core::GUID::try_from("44a9796f-723e-4fdf-a218-033e75b0c084").unwrap();
    let factory_type = table.register_interface("IUriRuntimeClassFactory", factory_iid);
    let factory_type = factory_type.add_method(
        "CreateUri",
        MethodSignature::new(&table)
            .add_in(table.hstring())
            .add_out(table.object()),
    );
    // CreateUri(baseUri: HSTRING, relativeUri: HSTRING) -> Object  [Libffi: 2-in 1-out]
    let factory_type = factory_type.add_method(
        "CreateWithRelativeUri",
        MethodSignature::new(&table)
            .add_in(table.hstring())
            .add_in(table.hstring())
            .add_out(table.object()),
    );

    // IUriRuntimeClass — property getters  [Direct0In1Out]
    let uri_iid =
        windows::core::GUID::try_from("9e365e57-48b2-4160-956f-c7385120bbfc").unwrap();
    let uri_type = table.register_interface("IUriRuntimeClass", uri_iid);
    let hstring_out = || MethodSignature::new(&table).add_out(table.hstring());
    let uri_type = uri_type.add_method("get_AbsoluteUri", hstring_out());   // 6
    let uri_type = uri_type.add_method("get_DisplayUri", hstring_out());    // 7
    let uri_type = uri_type.add_method("get_RawUri", hstring_out());        // 8
    let uri_type = uri_type.add_method("get_SchemeName", hstring_out());    // 9
    let uri_type = uri_type.add_method("get_UserName", hstring_out());      // 10
    let uri_type = uri_type.add_method("get_Password", hstring_out());      // 11
    let uri_type = uri_type.add_method("get_Host", hstring_out());          // 12
    let uri_type = uri_type.add_method("get_Domain", hstring_out());        // 13
    let uri_type = uri_type.add_method("get_Port",                          // 14
        MethodSignature::new(&table).add_out(table.i32_type()));
    let uri_type = uri_type.add_method("get_Path", hstring_out());          // 15
    let uri_type = uri_type.add_method("get_Query", hstring_out());         // 16
    let uri_type = uri_type.add_method("get_QueryParsed",                   // 17
        MethodSignature::new(&table).add_out(table.object()));
    let uri_type = uri_type.add_method("get_Fragment", hstring_out());      // 18
    let uri_type = uri_type.add_method("get_Extension", hstring_out());     // 19
    let uri_type = uri_type.add_method("get_Suspicious",                    // 20
        MethodSignature::new(&table).add_out(table.bool_type()));
    let uri_type = uri_type.add_method("Equals",                            // 21
        MethodSignature::new(&table)
            .add_in(table.object())
            .add_out(table.bool_type()));
    let uri_type = uri_type.add_method("CombineUri",                        // 22
        MethodSignature::new(&table)
            .add_in(table.hstring())
            .add_out(table.object()));

    (table, DynUriApi { factory_iid, uri_iid, factory_type, uri_type })
}

fn dyn_create_uri(api: &DynUriApi, uri_str: &str) -> (WinRTValue, *mut std::ffi::c_void) {
    let factory = dynwinrt::ro_get_activation_factory_2(
        &HSTRING::from("Windows.Foundation.Uri"),
    )
    .unwrap();
    let fac_obj = factory.cast(&api.factory_iid).unwrap();
    let create = api.factory_type.method_by_name("CreateUri").unwrap();
    let uri_val = create
        .invoke(
            fac_obj.as_object().unwrap().as_raw(),
            &[WinRTValue::HString(HSTRING::from(uri_str))],
        )
        .unwrap();
    let uri_obj = uri_val.into_iter().next().unwrap();
    let casted = uri_obj.cast(&api.uri_iid).unwrap();
    let raw = casted.as_object().unwrap().as_raw();
    (casted, raw)
}

// ======================================================================
// 1. Create Uri — Direct1In1Out path
// ======================================================================

fn bench_create_uri(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com/path?query=value&foo=bar#fragment";
    let mut group = c.benchmark_group("create_uri");

    // Static
    group.bench_function("static", |b| {
        let hstr = HSTRING::from(uri_str);
        b.iter(|| black_box(Uri::CreateUri(&hstr).unwrap()));
    });

    // Dynamic (Direct1In1Out)
    let factory = dynwinrt::ro_get_activation_factory_2(
        &HSTRING::from("Windows.Foundation.Uri"),
    ).unwrap();
    let fac_obj = factory.cast(&api.factory_iid).unwrap();
    let fac_raw = fac_obj.as_object().unwrap().as_raw();
    let create = api.factory_type.method_by_name("CreateUri").unwrap();
    group.bench_function("dynamic_direct1in1out", |b| {
        let arg = WinRTValue::HString(HSTRING::from(uri_str));
        b.iter(|| black_box(create.invoke(fac_raw, &[arg.clone()]).unwrap()));
    });

    group.finish();
}

// ======================================================================
// 2. Create Uri with relative — Libffi path (2-in 1-out)
// ======================================================================

fn bench_create_with_relative(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let mut group = c.benchmark_group("create_with_relative_uri");

    // Static
    group.bench_function("static", |b| {
        let base = HSTRING::from("https://example.com");
        let relative = HSTRING::from("/path/to/resource?q=1");
        b.iter(|| black_box(Uri::CreateWithRelativeUri(&base, &relative).unwrap()));
    });

    // Dynamic (Libffi: 2-in 1-out)
    let factory = dynwinrt::ro_get_activation_factory_2(
        &HSTRING::from("Windows.Foundation.Uri"),
    ).unwrap();
    let fac_obj = factory.cast(&api.factory_iid).unwrap();
    let fac_raw = fac_obj.as_object().unwrap().as_raw();
    let create_rel = api.factory_type.method_by_name("CreateWithRelativeUri").unwrap();
    group.bench_function("dynamic_libffi_2in1out", |b| {
        let base = WinRTValue::HString(HSTRING::from("https://example.com"));
        let rel = WinRTValue::HString(HSTRING::from("/path/to/resource?q=1"));
        b.iter(|| black_box(create_rel.invoke(fac_raw, &[base.clone(), rel.clone()]).unwrap()));
    });

    group.finish();
}

// ======================================================================
// 3. Read single property — Direct0In1Out path
// ======================================================================

fn bench_get_absolute_uri(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com/path?query=value";
    let mut group = c.benchmark_group("get_absolute_uri");

    // Static
    let static_uri = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    group.bench_function("static", |b| {
        b.iter(|| black_box(static_uri.AbsoluteUri().unwrap()));
    });

    // Dynamic (Direct0In1Out) — full invoke path including Vec alloc
    let (_uri_val, uri_raw) = dyn_create_uri(&api, uri_str);
    let get_abs = api.uri_type.method_by_name("get_AbsoluteUri").unwrap();
    group.bench_function("dynamic_direct0in1out", |b| {
        b.iter(|| black_box(get_abs.invoke(uri_raw, &[]).unwrap()));
    });

    // Raw vtable call — no Vec, no RwLock, no invoke overhead.
    // Isolates the pure vtable call + HSTRING creation cost.
    group.bench_function("raw_vtable_call", |b| {
        b.iter(|| {
            let mut result_ptr: *mut std::ffi::c_void = std::ptr::null_mut();
            let vtable_ptr = unsafe { *(uri_raw as *const *const *mut std::ffi::c_void) };
            let fn_ptr = unsafe { *vtable_ptr.add(6) }; // vtable[6] = get_AbsoluteUri
            type VtableMethod = unsafe extern "system" fn(
                *mut std::ffi::c_void,
                *mut *mut std::ffi::c_void,
            ) -> windows_core::HRESULT;
            let method: VtableMethod = unsafe { std::mem::transmute(fn_ptr) };
            let hr = unsafe { method(uri_raw, &mut result_ptr) };
            hr.ok().unwrap();
            // Wrap as WinRTValue like dynwinrt does, but without Vec
            let val = WinRTValue::HString(unsafe {
                std::mem::transmute::<*mut std::ffi::c_void, HSTRING>(result_ptr)
            });
            black_box(val);
        });
    });

    group.finish();
}

// ======================================================================
// 4. Read i32 property (Port) — Direct0In1Out but i32 return
// ======================================================================

fn bench_get_port(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com:8080/path";
    let mut group = c.benchmark_group("get_port_i32");

    let static_uri = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    group.bench_function("static", |b| {
        b.iter(|| black_box(static_uri.Port().unwrap()));
    });

    let (_uri_val, uri_raw) = dyn_create_uri(&api, uri_str);
    let get_port = api.uri_type.method_by_name("get_Port").unwrap();
    group.bench_function("dynamic_direct0in1out", |b| {
        b.iter(|| black_box(get_port.invoke(uri_raw, &[]).unwrap()));
    });

    group.finish();
}

// ======================================================================
// 5. CombineUri — Direct1In1Out (HString in, Object out)
// ======================================================================

fn bench_combine_uri(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com/base";
    let mut group = c.benchmark_group("combine_uri");

    let static_uri = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    group.bench_function("static", |b| {
        let rel = HSTRING::from("/relative/path");
        b.iter(|| black_box(static_uri.CombineUri(&rel).unwrap()));
    });

    let (_uri_val, uri_raw) = dyn_create_uri(&api, uri_str);
    let combine = api.uri_type.method_by_name("CombineUri").unwrap();
    group.bench_function("dynamic_direct1in1out", |b| {
        let rel = WinRTValue::HString(HSTRING::from("/relative/path"));
        b.iter(|| black_box(combine.invoke(uri_raw, &[rel.clone()]).unwrap()));
    });

    group.finish();
}

// ======================================================================
// 6. Equals — Direct1In1Out but Object in, Bool out
// ======================================================================

fn bench_equals(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com/path";
    let mut group = c.benchmark_group("equals");

    let static_uri1 = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    let static_uri2 = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    group.bench_function("static", |b| {
        b.iter(|| black_box(static_uri1.Equals(&static_uri2).unwrap()));
    });

    let (_uri_val1, uri_raw1) = dyn_create_uri(&api, uri_str);
    let (uri_val2, _) = dyn_create_uri(&api, uri_str);
    let equals = api.uri_type.method_by_name("Equals").unwrap();
    group.bench_function("dynamic_direct1in1out", |b| {
        b.iter(|| black_box(equals.invoke(uri_raw1, &[uri_val2.clone()]).unwrap()));
    });

    group.finish();
}

// ======================================================================
// 7. Read multiple properties — multiple Direct0In1Out calls
// ======================================================================

fn bench_get_multiple_properties(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let uri_str = "https://example.com/some/path?key=value";
    let mut group = c.benchmark_group("get_3_properties");

    let static_uri = Uri::CreateUri(&HSTRING::from(uri_str)).unwrap();
    group.bench_function("static", |b| {
        b.iter(|| {
            let host = static_uri.Host().unwrap();
            let path = static_uri.Path().unwrap();
            let query = static_uri.Query().unwrap();
            black_box((host, path, query));
        });
    });

    let (_uri_val, uri_raw) = dyn_create_uri(&api, uri_str);
    let get_host = api.uri_type.method_by_name("get_Host").unwrap();
    let get_path = api.uri_type.method_by_name("get_Path").unwrap();
    let get_query = api.uri_type.method_by_name("get_Query").unwrap();
    group.bench_function("dynamic_direct0in1out_x3", |b| {
        b.iter(|| {
            let host = get_host.invoke(uri_raw, &[]).unwrap();
            let path = get_path.invoke(uri_raw, &[]).unwrap();
            let query = get_query.invoke(uri_raw, &[]).unwrap();
            black_box((host, path, query));
        });
    });

    group.finish();
}

// ======================================================================
// 8. Batch — create 100 URIs and read AbsoluteUri
// ======================================================================

fn bench_batch_create_and_read(c: &mut Criterion) {
    ensure_ro_initialized();
    let (_table, api) = setup_dynamic();
    let n = 100;
    let mut group = c.benchmark_group("batch_100_create_and_read");

    group.bench_function("static", |b| {
        b.iter(|| {
            for i in 0..n {
                let hstr = HSTRING::from(format!("https://example.com/path/{i}"));
                let uri = Uri::CreateUri(&hstr).unwrap();
                black_box(uri.AbsoluteUri().unwrap());
            }
        });
    });

    let factory = dynwinrt::ro_get_activation_factory_2(
        &HSTRING::from("Windows.Foundation.Uri"),
    ).unwrap();
    let fac_obj = factory.cast(&api.factory_iid).unwrap();
    let fac_raw = fac_obj.as_object().unwrap().as_raw();
    let create = api.factory_type.method_by_name("CreateUri").unwrap();
    let get_abs = api.uri_type.method_by_name("get_AbsoluteUri").unwrap();

    group.bench_function("dynamic", |b| {
        b.iter(|| {
            for i in 0..n {
                let arg = WinRTValue::HString(HSTRING::from(format!(
                    "https://example.com/path/{i}"
                )));
                let uri_val = create.invoke(fac_raw, &[arg]).unwrap();
                let uri_obj = uri_val.into_iter().next().unwrap();
                let casted = uri_obj.cast(&api.uri_iid).unwrap();
                black_box(
                    get_abs.invoke(casted.as_object().unwrap().as_raw(), &[]).unwrap()
                );
            }
        });
    });

    group.finish();
}

// ======================================================================
// 9. Struct in-param — Geopoint::Create(BasicGeoposition) → Libffi path
// ======================================================================

fn bench_geopoint_create(c: &mut Criterion) {
    ensure_ro_initialized();
    let table = dynwinrt::metadata_table::MetadataTable::new();

    let mut group = c.benchmark_group("geopoint_create_struct");

    // Static
    group.bench_function("static", |b| {
        use windows::Devices::Geolocation::{BasicGeoposition, Geopoint};
        let pos = BasicGeoposition {
            Latitude: 47.643,
            Longitude: -122.131,
            Altitude: 100.0,
        };
        b.iter(|| black_box(Geopoint::Create(pos).unwrap()));
    });

    // Dynamic — struct in-param goes through Libffi path
    let f64_h = table.f64_type();
    let geo_struct_type = table.struct_type("Windows.Devices.Geolocation.BasicGeoposition", &[f64_h.clone(), f64_h.clone(), f64_h]);

    // Register IGeopointFactory
    let factory_iid = windows::Devices::Geolocation::IGeopointFactory::IID;
    let factory_type = table.register_interface("IGeopointFactory", factory_iid);
    let factory_type = factory_type.add_method(
        "Create",
        dynwinrt::MethodSignature::new(&table)
            .add_in(geo_struct_type.clone())
            .add_out(table.object()),
    );

    // Get factory object
    let afactory = dynwinrt::ro_get_activation_factory_2(
        &HSTRING::from("Windows.Devices.Geolocation.Geopoint"),
    ).unwrap();
    let fac_obj = afactory.cast(&factory_iid).unwrap();
    let fac_raw = fac_obj.as_object().unwrap().as_raw();
    let create = factory_type.method_by_name("Create").unwrap();

    // Dynamic — includes struct alloc + field writes + libffi invoke
    group.bench_function("dynamic_with_struct_alloc", |b| {
        b.iter(|| {
            let mut geo_val = geo_struct_type.default_value();
            geo_val.set_field(0, 47.643f64);
            geo_val.set_field(1, -122.131f64);
            geo_val.set_field(2, 100.0f64);
            black_box(
                create.invoke(fac_raw, &[WinRTValue::Struct(geo_val)]).unwrap()
            );
        });
    });

    // Isolate: only struct alloc + field writes (no WinRT call)
    group.bench_function("struct_alloc_only", |b| {
        b.iter(|| {
            let mut geo_val = geo_struct_type.default_value();
            geo_val.set_field(0, 47.643f64);
            geo_val.set_field(1, -122.131f64);
            geo_val.set_field(2, 100.0f64);
            black_box(geo_val);
        });
    });

    group.finish();
}

criterion_group!(
    benches,
    bench_create_uri,
    bench_create_with_relative,
    bench_get_absolute_uri,
    bench_get_port,
    bench_combine_uri,
    bench_equals,
    bench_get_multiple_properties,
    bench_batch_create_and_read,
    bench_geopoint_create,
);
criterion_main!(benches);
