use windows::{
    Data::Xml::Dom::*, Win32::Foundation::*, Win32::System::Threading::*,
    Win32::UI::WindowsAndMessaging::*, core::*,
};

mod call;
mod value;

pub fn add(left: u64, right: u64) -> u64 {
    left + right
}

#[cfg(test)]
mod tests {
    use libffi::raw::ffi_abi_FFI_WIN64;

    use crate::call::{DWinRTHRESULTValue, DWinRTPointerValue, DWinRTValueUnion};

    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }

    #[test]
    fn simple_windows_api_should_work() -> Result<()> {
        let doc = XmlDocument::new()?;
        doc.LoadXml(h!("<html>hello world</html>"))?;

        let root = doc.DocumentElement()?;
        assert!(root.NodeName()? == "html");
        assert!(root.InnerText()? == "hello world");

        unsafe {
            let event = CreateEventW(None, true, false, None)?;
            SetEvent(event)?;
            WaitForSingleObject(event, 0);
            CloseHandle(event)?;

            MessageBoxA(None, s!("Ansi"), s!("Caption"), MB_OK);
            MessageBoxW(None, w!("Wide测试"), w!("Caption"), MB_OK);
        }

        Ok(())
    }

    #[test]
    fn test_winrt_uri() -> Result<()> {
        use windows::Foundation::Uri;

        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;
        assert_eq!(uri.SchemeName()?, "https");
        assert_eq!(uri.Host()?, "www.example.com");
        assert_eq!(uri.Path()?, "/path");
        assert_eq!(uri.Query()?, "?query=1");
        assert_eq!(uri.Fragment()?, "#fragment");

        Ok(())
    }

    #[test]
    fn test_winrt_uri_interop() -> Result<()> {
        use windows::Foundation::Uri;
        println!("URI guid = {:?} ...", Uri::IID);
        println!("URI name = {:?} ...", Uri::NAME);

        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;
        unsafe {
            let raw = uri.as_raw();
            let ukn: IUnknown = uri.cast().expect("Cast to IUnknown should always work");

            // IUriRuntimeClass layout:
            // 0-2: IUnknown
            // 3-5: IInspectable
            // 6: get_AbsoluteUri
            // ...
            // 17: get_SchemeName

            // type GetSchemeName = extern "system" fn(*mut std::ffi::c_void, *mut HSTRING) -> HRESULT;
            // let get_scheme_name: GetSchemeName = std::mem::transmute(get_scheme_name_ptr);

            // let mut scheme = HSTRING::new();
            // // get_scheme_name(raw, &mut scheme).ok()?;
            // let result = interface::call_method_ptr_ptr_ret_hresult(
            //     interface::DynWinRTValue::Pointer(raw as *mut std::ffi::c_void),
            //     get_scheme_name_ptr,
            //     &[interface::DynWinRTValue::Pointer(
            //         &mut scheme as *mut _ as *mut std::ffi::c_void,
            //     )],
            // );
            let mut scheme = HSTRING::new();
            let obj = DWinRTPointerValue::from_com_object(&uri);
            let outPtr = DWinRTPointerValue::from_out_ptr(&mut scheme);
            call::call_method_1(17, obj, outPtr);

            assert_eq!(scheme, "https");
        }

        Ok(())
    }

    #[test]
    fn test_winrt_uri_interop_using_libffi() -> Result<()> {
        use libffi::middle::*;
        use std::ffi::c_void;
        use windows::Foundation::Uri;
        println!("URI guid = {:?} ...", Uri::IID);
        println!("URI name = {:?} ...", Uri::NAME);

        let uri = Uri::CreateUri(h!("http://www.example.com/path?query=1#fragment"))?;
        let mut hstr = HSTRING::new();
        let outPtr: *mut c_void = std::ptr::from_mut(&mut hstr).cast();
        let obj = uri.as_raw();
        let fptr = call::get_vtable_function_ptr(obj.cast(), 17);
        unsafe {
            let pars = vec![Type::pointer(), Type::pointer()];
            let cif = Cif::new(pars, Type::i32());
            let args = &[arg(&obj), arg(&outPtr)];
            let hr: HRESULT = cif.call(CodePtr(fptr), args);
            hr.ok()?;
            assert_eq!(hstr, "http");
        }

        Ok(())
    }

    #[test]
    fn test_winrt_uri_interop_using_argument() -> Result<()> {
        use windows::Foundation::Uri;
        println!("URI guid = {:?} ...", Uri::IID);
        println!("URI name = {:?} ...", Uri::NAME);

        let s = HSTRING::new();

        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;
        let mut out_hstr_arg = value::Argument::out_hstring();
        let hr = call::call_method_with_values(17, uri.as_raw(), &mut [&mut out_hstr_arg]);
        hr.ok()?;
        assert_eq!(out_hstr_arg.as_hstring(), "https");

        let mut out_host = value::Argument::out_hstring();
        let hr = call::call_method_with_values(11, uri.as_raw(), &mut [&mut out_host]);
        hr.ok()?;

        assert_eq!(out_host.as_hstring(), "www.example.com");
        Ok(())
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
    fn test_calling_simple_add_extern_c() {
        use windows_core::HRESULT;
        extern "C" fn add(x: f64, y: &f64) -> f64 {
            return x + y;
        }

        use libffi::middle::*;
        let ok = HRESULT(0);

        let args = vec![Type::f64(), Type::pointer()];
        let cif = Cif::new(args.into_iter(), Type::f64());

        let arhr = arg(&ok);

        let n = unsafe { cif.call(CodePtr(add as *mut _), &[arg(&5f64), arg(&&6f64)]) };
        assert_eq!(11f64, n);
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
}
