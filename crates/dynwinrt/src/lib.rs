use windows::Foundation::IStringable_Impl;
use windows::Storage::{FileAccessMode, IStorageFileStatics};
use windows::{Data::Xml::Dom::*, core::*};
use windows_future::IAsyncOperation;

mod abi;
mod call;
mod interfaces;
mod result;
mod roapi;
mod signature;
mod value;
mod winapp;

mod array;
mod bindings;
mod dasync;
mod meta;
pub mod metadata_table;

pub struct IIds;
impl IIds {
    pub const IFileOpenPickerFactory: windows_core::GUID = bindings::IFileOpenPickerFactory::IID;
    pub const IAsyncOperationPickFileResult: windows_core::GUID =
        IAsyncOperation::<bindings::PickFileResult>::IID;
    pub const ITextRecognizerStatics: windows_core::GUID = bindings::ITextRecognizerStatics::IID;
    pub const IImageBufferStatics: windows_core::GUID = bindings::IImageBufferStatics::IID;
    pub const ISoftwareBitmap: windows_core::GUID = bindings::SoftwareBitmap::IID;
    pub const IAsyncOperationTextRecognizer: windows_core::GUID =
        IAsyncOperation::<bindings::TextRecognizer>::IID;
    pub const IAsyncOperationRecognizedText: windows_core::GUID =
        IAsyncOperation::<bindings::RecognizedText>::IID;
    pub const TextRecognizer: windows_core::GUID = bindings::TextRecognizer::IID;
    pub const RecognizedText: windows_core::GUID = bindings::RecognizedText::IID;
}

pub fn export_add(x: f64, y: &f64) -> f64 {
    println!("export_add called with x = {}, y = {}", x, y);
    return x + y;
}

use crate::call::get_vtable_function_ptr;
pub use crate::result::Result;
use crate::roapi::query_interface;
pub use crate::roapi::ro_get_activation_factory_2;
pub use crate::signature::{InterfaceSignature, MethodSignature};
pub use crate::metadata_table::{TypeHandle, TypeKind, MetadataTable, MethodHandle, ValueTypeData};
pub use crate::array::ArrayData;
pub use crate::value::WinRTValue;
use crate::winapp::pick_path;
pub use crate::winapp::test_pick_open_picker_full_dynamic;
pub use crate::winapp::{WinAppSdkContext, initialize_winappsdk};
pub use interfaces::uri_vtable;

#[implement(windows::Foundation::IStringable)]
struct MyComObject {
    pub data: i32,
}

impl IStringable_Impl for MyComObject_Impl {
    fn ToString(&self) -> windows_core::Result<windows_core::HSTRING> {
        Ok(HSTRING::from(format!(
            "MyComObject ToString with data: {}",
            self.data
        )))
    }
}

impl Drop for MyComObject {
    fn drop(&mut self) {
        println!("MyComObject is being dropped");
    }
}

pub async fn get_async_string(op_string: IAsyncOperation<HSTRING>) -> windows_core::Result<String> {
    let s = op_string.await?;
    Ok(s.to_string())
}

pub async fn http_get_string(url: &str) -> windows_core::Result<String> {
    use windows::Foundation::Uri;
    use windows::Web::Http::HttpClient;
    let uri = Uri::CreateUri(&HSTRING::from(url))?;
    let http_client = HttpClient::new()?;
    let op = http_client.GetStringAsync(&uri)?;
    let s = op.await?;
    Ok(s.to_string())
}

#[cfg(test)]
mod tests {
    use windows::Foundation::{IStringable, Uri};

    use crate::value::WinRTValue;

    use super::*;

    #[test]
    fn from_raw_borrow_test() -> Result<()> {
        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment")).unwrap();
        let raw = uri.as_raw();
        let uri2 = unsafe { Uri::from_raw_borrowed(&raw) }.unwrap();
        let _ = uri2.clone();

        assert_eq!(uri.SchemeName().unwrap(), uri2.SchemeName().unwrap());
        Ok(())
    }

    #[test]
    fn http_call() -> Result<()> {
        futures::executor::block_on(async {
            use std::future::IntoFuture;

            let uri = Uri::CreateUri(h!("https://www.microsoft.com"))?;
            uri.SchemeName()?; // "http"
            uri.Host();
            uri.Port();
            let op = windows::Web::Http::HttpClient::new()?.GetStringAsync(&uri)?;
            let s = op.into_future().await?;
            // println!("Response: {}", s);

            Ok(())
        })
    }

    #[test]
    fn customCom() -> Result<()> {
        let my_obj = MyComObject { data: 42 };
        let sta: IStringable = my_obj.into();
        assert_eq!(
            sta.ToString()?.to_os_string(),
            "MyComObject ToString with data: 42"
        );
        Ok(())
    }

    #[test]
    fn simple_windows_api_should_work() -> Result<()> {
        let doc = XmlDocument::new()?;
        doc.LoadXml(h!("<html>hello world</html>"))?;

        let root = doc.DocumentElement()?;
        assert!(root.NodeName()? == "html");
        assert!(root.InnerText()? == "hello world");

        Ok(())
    }

    #[test]
    fn test_winrt_uri() -> Result<()> {
        use windows::Foundation::Uri;
        println!("URI guid = {:?} ...", Uri::IID);
        println!("URI name = {:?} ...", Uri::NAME);
        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;
        assert_eq!(uri.SchemeName()?, "https");
        assert_eq!(uri.Host()?, "www.example.com");
        assert_eq!(uri.Path()?, "/path");
        assert_eq!(uri.Query()?, "?query=1");
        assert_eq!(uri.Fragment()?, "#fragment");

        Ok(())
    }

    #[test]
    fn test_winrt_uri_interop_using_libffi() -> Result<()> {
        use libffi::middle::*;
        use std::ffi::c_void;
        use windows::Foundation::Uri;

        let uri = Uri::CreateUri(h!("http://www.example.com/path?query=1#fragment"))?;
        let mut hstr = HSTRING::new();
        let out_ptr: *mut c_void = std::ptr::from_mut(&mut hstr).cast();
        let obj = uri.as_raw();
        let fptr = call::get_vtable_function_ptr(obj.cast(), 17);
        let pars = vec![Type::pointer(), Type::pointer()];
        let cif = Cif::new(pars, Type::i32());
        let args = &[arg(&obj), arg(&out_ptr)];
        let hr: HRESULT = unsafe { cif.call(CodePtr(fptr), args) };
        hr.ok()?;
        assert_eq!(hstr, "http");

        Ok(())
    }

    #[test]
    fn test_winrt_uri_interop_using_signature() -> Result<()> {
        use windows::Foundation::Uri;

        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;

        let reg = metadata_table::MetadataTable::new();
        let vtable = uri_vtable(&reg);

        let get_runtime_classname = &vtable.methods[4];
        assert_eq!(
            get_runtime_classname.call_dynamic(uri.as_raw(), &[])?[0]
                .as_hstring()
                .unwrap(),
            "Windows.Foundation.Uri"
        );

        let get_scheme = &vtable.methods[17];
        let scheme = get_scheme.call_dynamic(uri.as_raw(), &[])?;
        assert_eq!(scheme[0].as_hstring().unwrap(), "https");
        let get_path = &vtable.methods[13];
        let path = get_path.call_dynamic(uri.as_raw(), &[])?;
        assert_eq!(path[0].as_hstring().unwrap(), "/path");
        let get_port = &vtable.methods[19];
        let port = get_port.call_dynamic(uri.as_raw(), &[])?;
        assert_eq!(port[0].as_i32().unwrap(), 443);

        Ok(())
    }

    #[test]
    fn test_uri_call_dynamic() -> Result<()> {
        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment")).unwrap();
        let reg = metadata_table::MetadataTable::new();
        let vtable = interfaces::uri_vtable(&reg);
        let res = vtable.methods[17].call_dynamic(uri.as_raw(), &[])?;
        assert_eq!(res[0].as_hstring().unwrap(), "https");
        Ok(())
    }

    #[test]
    fn test_calling_simple_add_extern_c() {
        extern "C" fn add(x: f64, y: &f64) -> f64 {
            return x + y;
        }

        use libffi::middle::*;

        let args = vec![Type::f64(), Type::pointer()];
        let cif = Cif::new(args.into_iter(), Type::f64());

        let n: f64 = unsafe { cif.call(CodePtr(add as *mut _), &[arg(&5f64), arg(&&6f64)]) };
        assert_eq!(11f64, n);
    }

    #[test]
    fn test_winmd_read() {
        use windows_metadata::*;

        let index = reader::Index::read(
            r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd",
        )
        .unwrap();

        let def = index.expect("Windows.Foundation", "Point");
        assert_eq!(def.namespace(), "Windows.Foundation");
        assert_eq!(def.name(), "Point");

        let extends = def.extends().unwrap();
        assert_eq!(extends.namespace(), "System");
        assert_eq!(extends.name(), "ValueType");

        let fields: Vec<_> = def.fields().collect();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].name(), "X");
        assert_eq!(fields[1].name(), "Y");
        assert_eq!(fields[0].ty(), Type::F32);
        assert_eq!(fields[1].ty(), Type::F32);
        println!("{:?}", fields);
    }

    #[test]
    fn test_winmd_read_uri() {
        use windows_metadata::*;
        let index = reader::Index::read(
            r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd",
        )
        .unwrap();
        let def = index.expect("Windows.Foundation", "Uri");
        // list all methods and print their signatures
        for method in def.methods() {
            let params: Vec<String> = method
                .params()
                .enumerate()
                .map(|(i, p)| format!("{}: {:?}", p.name(), method.signature(&[]).types))
                .collect();
            println!(
                "fn {:?} {}({}) -> {:?}",
                method.flags(),
                method.name(),
                params.join(", "),
                method.signature(&[]).return_type
            );
        }
    }

    #[test]
    fn test_winmd_read_http_client() {
        use windows_metadata::*;
        let index = reader::Index::read(
            r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd",
        )
        .unwrap();
        let def = index.expect("Windows.Web.Http", "HttpClient");
        // list all methods and print their signatures
        for (i, method) in def.methods().enumerate() {
            let params: Vec<String> = method
                .params()
                .enumerate()
                .map(|(i, p)| format!("{}: {:?}", p.name(), ""))
                .collect();
            println!(
                "#{}, fn {:?} {}({}) -> {:?}",
                i,
                // method.flags(),
                "",
                method.name(),
                params.join(", "),
                method.signature(&[]).return_type
            );
        }
    }

    #[test]
    fn test_struct_in_param_geopoint_create() -> Result<()> {
        use windows::Devices::Geolocation::Geopoint;
        use windows::Win32::System::WinRT::{
            IActivationFactory, RO_INIT_MULTITHREADED, RoGetActivationFactory, RoInitialize,
        };

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        // 1. Define BasicGeoposition { Latitude: f64, Longitude: f64, Altitude: f64 }
        let reg = metadata_table::MetadataTable::new();
        let f64_h = reg.f64_type();
        let geo_type = reg.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);

        // 2. Create and populate struct value
        let mut geo_val = geo_type.default_value();
        geo_val.set_field(0, 47.643f64);
        geo_val.set_field(1, -122.131f64);
        geo_val.set_field(2, 100.0f64);

        // 3. Get IGeopointFactory
        let afactory = unsafe {
            RoGetActivationFactory::<IActivationFactory>(
                h!("Windows.Devices.Geolocation.Geopoint"),
            )
        }?;
        let geopoint_factory =
            afactory.cast::<windows::Devices::Geolocation::IGeopointFactory>()?;

        // 4. Define IGeopointFactory with Create method at vtable index 6
        // ABI: fn(this, BasicGeoposition, *out) -> HRESULT
        let mut factory_sig = InterfaceSignature::define_from_iinspectable(
            "IGeopointFactory",
            windows::Devices::Geolocation::IGeopointFactory::IID,
            &reg,
        );
        factory_sig.add_method(
            MethodSignature::new(&reg)
                .add_in(geo_type.clone())
                .add_out(reg.object()),
        );

        // 5. Call Create via MethodSignature
        let create_method = &factory_sig.methods[6];
        let results = create_method.call_dynamic(
            geopoint_factory.as_raw(),
            &[WinRTValue::Struct(geo_val)],
        )?;

        // 6. Verify
        let geopoint_obj = results[0].as_object().unwrap();
        let geopoint: Geopoint = geopoint_obj.cast()?;
        let pos = geopoint.Position()?;
        assert!((pos.Latitude - 47.643).abs() < 1e-6);
        assert!((pos.Longitude - (-122.131)).abs() < 1e-6);
        assert!((pos.Altitude - 100.0).abs() < 1e-6);

        Ok(())
    }

    #[test]
    fn test_struct_out_param_geopoint_position() -> Result<()> {
        use windows::Devices::Geolocation::{BasicGeoposition, Geopoint, IGeopoint};
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        // 1. Create a Geopoint using static projection (known-good)
        let position = BasicGeoposition {
            Latitude: 47.643,
            Longitude: -122.131,
            Altitude: 100.0,
        };
        let geopoint = Geopoint::Create(position)?;

        // 2. Define BasicGeoposition struct type
        let reg = metadata_table::MetadataTable::new();
        let f64_h = reg.f64_type();
        let geo_type = reg.define_struct(&[f64_h.clone(), f64_h.clone(), f64_h]);

        // 3. Define IGeopoint with get_Position at vtable index 6
        // ABI: fn(this, *out_BasicGeoposition) -> HRESULT
        let mut igeopoint_sig = InterfaceSignature::define_from_iinspectable(
            "IGeopoint",
            IGeopoint::IID,
            &reg,
        );
        igeopoint_sig.add_method(
            MethodSignature::new(&reg)
                .add_out(geo_type),
        );

        // 4. Call get_Position via MethodSignature
        let igeopoint: IGeopoint = geopoint.cast()?;
        let get_position = &igeopoint_sig.methods[6];
        let results = get_position.call_dynamic(igeopoint.as_raw(), &[])?;

        // 5. Verify
        let data = results[0].as_struct().expect("Expected WinRTValue::Struct");
        let lat: f64 = data.get_field(0);
        let lon: f64 = data.get_field(1);
        let alt: f64 = data.get_field(2);
        assert!((lat - 47.643).abs() < 1e-6);
        assert!((lon - (-122.131)).abs() < 1e-6);
        assert!((alt - 100.0).abs() < 1e-6);

        Ok(())
    }

    #[test]
    fn test_pass_array_create_int32() -> Result<()> {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        // IPropertyValueStatics: {629BDBC8-D932-4FF4-96B9-8D96C5C1E858}
        let statics_iid = windows_core::GUID::from_u128(0x629BDBC8_D932_4FF4_96B9_8D96C5C1E858);

        // Get factory and QI to IPropertyValueStatics
        let factory = WinRTValue::from_activation_factory(h!("Windows.Foundation.PropertyValue")).unwrap();
        let statics = factory.cast(&statics_iid).unwrap();

        // Build IPropertyValueStatics interface signature
        // vtable[6..28] = scalar Create methods (23 methods)
        // vtable[29] = CreateInt32Array(UINT32 length, INT32* data, IInspectable** result)
        let reg = metadata_table::MetadataTable::new();
        let mut iface = InterfaceSignature::define_from_iinspectable(
            "IPropertyValueStatics",
            statics_iid,
            &reg,
        );
        for _ in 0..23 {
            iface.add_method(MethodSignature::new(&reg)); // placeholders for vtable[6..28]
        }
        iface.add_method(
            MethodSignature::new(&reg)
                .add_in(reg.array(&reg.i32_type()))
                .add_out(reg.object()),
        );

        // Create array data
        let array_reg = metadata_table::MetadataTable::new();
        let array_arg = WinRTValue::Array(value::ArrayData::from_values(
            array_reg.i32_type(),
            &[WinRTValue::I32(10), WinRTValue::I32(20), WinRTValue::I32(30)],
        ));

        // Call CreateInt32Array at vtable index 29
        let results = iface.methods[29].call_dynamic(
            statics.as_object().unwrap().as_raw(),
            &[array_arg],
        )?;

        assert_eq!(results.len(), 1);
        let result_obj = results[0].as_object().unwrap();

        // Verify by reading back via static projection
        let inspectable: windows_core::IInspectable = result_obj.cast()?;
        let pv: windows::Foundation::IPropertyValue = inspectable.cast()?;
        let mut readback = windows::core::Array::<i32>::new();
        pv.GetInt32Array(&mut readback)?;
        assert_eq!(&readback[..], &[10, 20, 30]);

        Ok(())
    }

    #[test]
    fn test_receive_array_get_int32() -> Result<()> {
        use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};

        let _ = unsafe { RoInitialize(RO_INIT_MULTITHREADED) };

        // Create a PropertyValue with known data via static projection
        let test_data = vec![100i32, 200, 300, 400, 500];
        let prop = windows::Foundation::PropertyValue::CreateInt32Array(&test_data)?;

        // IPropertyValue: {4BD682DD-7554-40E9-9A9B-82654EDE7E62}
        let ipv_iid = windows_core::GUID::from_u128(0x4BD682DD_7554_40E9_9A9B_82654EDE7E62);

        // QI to IPropertyValue
        let prop_unk: IUnknown = prop.cast()?;
        let mut ipv_ptr = std::ptr::null_mut();
        unsafe { prop_unk.query(&ipv_iid, &mut ipv_ptr) }.ok()?;
        let ipv = unsafe { IUnknown::from_raw(ipv_ptr) };

        // Build IPropertyValue interface signature
        // vtable[6..25] = scalar Get methods (20 methods)
        // vtable[26..28] = GetUInt8Array, GetInt16Array, GetUInt16Array
        // vtable[29] = GetInt32Array(UINT32* length, INT32** data)
        let reg = metadata_table::MetadataTable::new();
        let mut iface = InterfaceSignature::define_from_iinspectable(
            "IPropertyValue",
            ipv_iid,
            &reg,
        );
        for _ in 0..23 {
            iface.add_method(MethodSignature::new(&reg)); // placeholders for vtable[6..28]
        }
        iface.add_method(
            MethodSignature::new(&reg)
                .add_out(reg.array(&reg.i32_type())),
        );

        // Call GetInt32Array at vtable index 29
        let results = iface.methods[29].call_dynamic(ipv.as_raw(), &[])?;

        // Verify
        let array = results[0].as_array().expect("Expected WinRTValue::Array");
        assert_eq!(array.len(), 5);
        assert_eq!(array.get_i32(0), 100);
        assert_eq!(array.get_i32(1), 200);
        assert_eq!(array.get_i32(2), 300);
        assert_eq!(array.get_i32(3), 400);
        assert_eq!(array.get_i32(4), 500);

        Ok(())
    }
}

mod from_api_test {
    use windows::{Win32::System::Com::IUri, core::*};
    link!("urlmon.dll" "system" fn CreateUri(
    uri: PCWSTR, 
    flags: u32, 
    reserved: usize, 
    result: *mut *mut core::ffi::c_void) -> HRESULT);
    fn main() -> Result<()> {
        unsafe {
            let mut uri = core::mem::zeroed();
            let uri: IUri = CreateUri(w!("http://kennykerr.ca"), 0, 0, &mut uri)
                .and_then(|| Type::from_abi(uri))?;
            let domain = uri.GetDomain()?;
            let port = uri.GetPort()?;
            println!("{:?} ({port})", domain);
        }
        Ok(())
    }
}

pub fn windows_ai_ocr_api_call() {
    println!("windows_ai_ocr_api_call called");
    use bindings::*;
    use winapp::*;
    let options = WinAppSdkBootstrapOptions {
        major_version: 1,
        minor_version: 8,
        build_version: 0,
        revision_version: 0,
        bootstrap_dll_path: None,
    };
    // let result = initialize(options);
    // assert!(result.is_ok(), "message {}", result.err().unwrap().message());
    let factory =
        roapi::ro_get_activation_factory(h!("Microsoft.Windows.AI.Imaging.TextRecognizer"));
    assert!(
        factory.is_ok(),
        "message {}",
        factory.err().unwrap().message()
    );
    println!("WinAppSdk initialized");
    let ready_state = TextRecognizer::GetReadyState().unwrap();
    println!("TextRecognizer ready state: {:?}", ready_state);
}

pub use crate::winapp::get_bitmap_from_file;

pub async fn windows_ai_ocr_api_call_projected(path: &str) -> windows::core::Result<()> {
    use bindings::*;
    use winapp::*;
    println!("windows_ai_ocr_api_call_projected called");

    let options = WinAppSdkBootstrapOptions {
        major_version: 1,
        minor_version: 8,
        build_version: 0,
        revision_version: 0,
        bootstrap_dll_path: None,
    };
    // initialize(options)?;
    println!("WinAppSdk initialized");
    let ready_state = TextRecognizer::GetReadyState().unwrap();
    println!("TextRecognizer ready state: {:?}", ready_state);

    if ready_state == AIFeatureReadyState::NotReady {
        let operation = TextRecognizer::EnsureReadyAsync()?.await?;
        if operation.Status()? != AIFeatureReadyResultState::Success {
            return Err(Error::from_hresult(HRESULT(1).into() /* E_FAIL */));
        }
    }

    let recognizer = TextRecognizer::CreateAsync()?.await?;
    println!("TextRecognizer created successfully");

    // Load SoftwareBitmap from file path (similar to C# BitmapDecoder pattern)
    // use windows::Storage::{StorageFile, FileAccessMode};
    use windows::Storage::StorageFile;

    let file = StorageFile::GetFileFromPathAsync(&HSTRING::from(path))?.await?;
    println!("File loaded successfully from path: {}", path);
    let stream = file.OpenAsync(FileAccessMode::Read)?.await?;
    println!("File stream opened successfully");
    let decoder = windows::Graphics::Imaging::BitmapDecoder::CreateAsync(&stream)?.await?;
    println!("BitmapDecoder created successfully");
    let bitmap = decoder.GetSoftwareBitmapAsync()?.await?;
    println!("SoftwareBitmap obtained successfully");
    let raw = bitmap.as_raw();
    let bitmapb = unsafe { bindings::SoftwareBitmap::from_raw_borrowed(&raw) };
    println!("SoftwareBitmap wrapped successfully");
    let image_buffer = ImageBuffer::CreateForSoftwareBitmap(bitmapb).unwrap();
    println!("ImageBuffer created from file successfully");
    let result = recognizer
        .RecognizeTextFromImageAsync(&image_buffer)?
        .await?;
    println!("Text recognition completed successfully");
    // TODO: Call recognizer.RecognizeTextFromImage(image_buffer) when bitmap is available
    let lines = result.Lines()?;
    for i in 0..lines.len() {
        let line = lines[i].as_ref().unwrap();
        println!("Recognized line: {}", line.Text()?);
    }
    for line in result.Lines()?.into_iter() {
        println!("Recognized line: {}", line.as_ref().unwrap().Text()?);
    }
    Ok(())
}

pub fn print_ocr_paths(orc_result: WinRTValue) {
    let result = orc_result
        .as_object()
        .unwrap().cast::<bindings::RecognizedText>()
        .unwrap();
    let lines = result.Lines().unwrap();
    for i in 0..lines.len() {
        let text = lines[i].as_ref().unwrap().Text().unwrap();
        println!("Recognized line {}: {}", i, text);
    }
}

pub async fn ocr_demo(bitmap: WinRTValue) -> WinRTValue {
    // use bindings::*;
    use winapp::*;
    let reg = metadata_table::MetadataTable::new();
    let factory =
        WinRTValue::from_activation_factory(h!("Microsoft.Windows.AI.Imaging.TextRecognizer"))
            .unwrap();
    let text_recongizer_factory = factory.cast(&IIds::ITextRecognizerStatics).unwrap();
    let get_ready_state = signature::MethodSignature::new(&reg).add_out(reg.i32_type()).build(6);
    let ready_state = get_ready_state
        .call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap()[0]
        .as_i32()
        .unwrap();
    println!("TextRecognizer ready state: {:?}", ready_state);

    // if ready_state != AIFeatureReadyState::Ready.0 {
    if (ready_state != 0) {
        panic!("TextRecognizer is not ready");
    }
    println!("Creating TextRecognizer asynchronously...");

    let async_op_type = reg.async_operation(&reg.runtime_class(
        "Microsoft.Windows.AI.Imaging.TextRecognizer".into(),
        bindings::TextRecognizer::IID,
    ));
    let create_async = signature::MethodSignature::new(&reg).add_out(async_op_type).build(8);
    let recognizer_v = create_async
        .call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap().into_iter().next().unwrap();
    println!("TextRecognizer created successfully");
    let recognizer = recognizer_v.await.unwrap();
    println!("SoftwareBitmap wrapped successfully");
    let bitmapt = bitmap.cast(&IIds::ISoftwareBitmap).unwrap();
    let image_buffer_af = WinRTValue::from_activation_factory(h!("Microsoft.Graphics.Imaging.ImageBuffer")).unwrap();
    let image_buffer_static = image_buffer_af.cast(&IIds::IImageBufferStatics).unwrap();
    let create_for_bitmap = signature::MethodSignature::new(&reg).add_in(reg.object()).add_out(reg.object()).build(7);
    let image_buffer = create_for_bitmap
        .call_dynamic(image_buffer_static.as_object().unwrap().as_raw(), &[bitmapt])
        .unwrap().into_iter().next().unwrap()
        .as_object()
        .unwrap();
    println!("ImageBuffer created from file successfully");
    let recognize = signature::MethodSignature::new(&reg).add_in(reg.i64_type()).add_out(reg.object()).build(7);
    let res = recognize
        .call_dynamic(recognizer.as_object().unwrap().as_raw(), &[WinRTValue::I64(image_buffer.as_raw() as i64)])
        .unwrap().into_iter().next().unwrap();
        // .as_object()
        // .unwrap();
    let result = res.cast(&IIds::RecognizedText).unwrap();
    println!("Text recognition completed successfully");
    // TODO: Call recognizer.RecognizeTextFromImage(image_buffer) when bitmap is available
    // let lines = result.Lines().unwrap();
    // for i in 0..lines.len() {
    //     let line = lines[i].as_ref().unwrap();
    //     println!("Recognized line: {}", line.Text().unwrap());
    // }
    // print_ocr_paths(result);
    result
}

pub async fn windows_ai_ocr_api_call_dynamic(path: &str) -> result::Result<()> {
    use bindings::*;
    use winapp::*;
    let reg = metadata_table::MetadataTable::new();
    println!("WinAppSdk initialized");
    let factory =
        WinRTValue::from_activation_factory(h!("Microsoft.Windows.AI.Imaging.TextRecognizer"))
            .unwrap();
    let text_recongizer_factory = factory.cast(&IIds::ITextRecognizerStatics).unwrap();
    let get_ready_state = signature::MethodSignature::new(&reg).add_out(reg.i32_type()).build(6);
    let ready_state = get_ready_state
        .call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap()[0]
        .as_i32()
        .unwrap();
    println!("TextRecognizer ready state: {:?}", ready_state);

    if ready_state == AIFeatureReadyState::NotReady.0 {
        println!("TextRecognizer is not ready, calling EnsureReadyAsync...");
        let operation = TextRecognizer::EnsureReadyAsync().unwrap().await.unwrap();
        if operation.Status()? != AIFeatureReadyResultState::Success {
            return Err(result::Error::WindowsError(Error::from_hresult(
                HRESULT(1).into(), /* E_FAIL */
            )));
        }
    }

    println!("Creating TextRecognizer asynchronously...");

    // let f2 = factory.as_object().unwrap().cast::<ITextRecognizerStatics>().unwrap();
    let async_op_type = reg.async_operation(&reg.runtime_class(
        "Microsoft.Windows.AI.Imaging.TextRecognizer".into(),
        bindings::TextRecognizer::IID,
    ));
    let create_async = signature::MethodSignature::new(&reg).add_out(async_op_type).build(8);
    let recognizer_v = create_async
        .call_dynamic(text_recongizer_factory.as_object().unwrap().as_raw(), &[])
        .unwrap().into_iter().next().unwrap();
    println!("TextRecognizer created successfully");
    let recognizer = recognizer_v.await?;

    // let recognizer = recognizer_o.as_object().unwrap();
    // .cast::<TextRecognizer>()?;

    // let recognizer = TextRecognizer::CreateAsync().unwrap().await?;

    // Load SoftwareBitmap from file path (similar to C# BitmapDecoder pattern)
    // use windows::Storage::{StorageFile, FileAccessMode};
    use windows::Storage::StorageFile;

    let file = StorageFile::GetFileFromPathAsync(&HSTRING::from(path))?.await?;
    println!("File loaded successfully from path: {}", path);
    let stream = file.OpenAsync(FileAccessMode::Read)?.await?;
    println!("File stream opened successfully");
    let decoder = windows::Graphics::Imaging::BitmapDecoder::CreateAsync(&stream)?.await?;
    println!("BitmapDecoder created successfully");
    let bitmap = decoder.GetSoftwareBitmapAsync()?.await?;
    println!("SoftwareBitmap obtained successfully");
    let raw = bitmap.as_raw();
    let bitmapb = unsafe { bindings::SoftwareBitmap::from_raw_borrowed(&raw) };
    println!("SoftwareBitmap wrapped successfully");
    let image_buffer = ImageBuffer::CreateForSoftwareBitmap(bitmapb).unwrap();
    println!("ImageBuffer created from file successfully");

    // let rukn = unsafe { IUnknown::from_raw(recognizer.as_raw())};
    // let val = WinRTValue::Object(rukn);
    // std::mem::forget(recognizer);
    let recognize = signature::MethodSignature::new(&reg).add_in(reg.i64_type()).add_out(reg.object()).build(7);
    let res = recognize
        .call_dynamic(recognizer.as_object().unwrap().as_raw(), &[WinRTValue::I64(image_buffer.as_raw() as i64)])
        .unwrap().into_iter().next().unwrap();
        // .as_object()
        // .unwrap();
    // let result: bindings::RecognizedText = res.cast().unwrap();

    // let fptr = get_vtable_function_ptr(recognizer.as_raw(), 7);
    // let mut out = std::ptr::null_mut();
    // unsafe {
    //     let method: extern "system" fn(
    //         *mut std::ffi::c_void,
    //         *mut std::ffi::c_void,
    //         *mut *mut std::ffi::c_void,
    //     ) -> HRESULT = std::mem::transmute(fptr);
    //     let hr = method(
    //         recognizer.as_raw(),
    //         image_buffer.as_raw() as *const _ as *mut _,
    //         // paramAbi.abi() as *mut _,
    //         &mut out,
    //     );
    //     println!("Text recognition dynamic call out HRESULT {:?}", hr);
    //     hr.ok()
    //         .map_err(|e| {
    //             println!("Error calling method: {:?}", e);
    //             e
    //         })
    //         .unwrap();
    // }
    // println!("Text recognition dynamic call out result {:?}", out);

    // let result = unsafe { bindings::RecognizedText::from_raw(out) };

    // let result = recognizer.RecognizeTextFromImage(&image_buffer).unwrap();

    // std::mem::forget(val);
    println!("Text recognition completed successfully");
    // TODO: Call recognizer.RecognizeTextFromImage(image_buffer) when bitmap is available
    // let lines = result.Lines()?;
    // for i in 0..lines.len() {
    //     let line = lines[i].as_ref().unwrap();
    //     println!("Recognized line: {}", line.Text()?);
    // }
    // for line in result.Lines()?.into_iter() {
    //     println!("Recognized line: {}", line.as_ref().unwrap().Text()?);
    // }
    print_ocr_paths(res);
    Ok(())
}

#[cfg(test)]
mod tests2 {
    use windows_core::IUnknown;

    #[test]
    fn size_of_iunknown() {
        println!(
            "Size of IUnknown: {} bytes",
            core::mem::size_of::<IUnknown>()
        );
    }
}
