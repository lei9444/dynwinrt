use windows::{Data::Xml::Dom::*, Win32::UI::WindowsAndMessaging::*, core::*};

mod call;
mod value;

pub fn export_add(x: f64, y: &f64) -> f64 {
    return x + y;
}

pub fn uri_vtable() -> value::VTableSignature {
    let mut vtable = value::VTableSignature::new();
    vtable
        .add_method(|m| m) // 0 QueryInterface
        .add_method(|m| m) // 1 AddRef
        .add_method(|m| m) // 2 Release
        .add_method(|m| m) // 3 GetIids
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 4 GetRuntimeClassName
        .add_method(|m| m) // 5 GetTrustLevel
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 6 get_AbsoluteUri
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 7 get_DisplayUri
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 8 get_Domain
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 9 get_Extension
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 10 get_Fragment
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 11 get_Host
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 12 get_Password
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 13 get_Path
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 14 get_Query
        .add_method(|m| m) // 15 get_QueryParsed
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 16 get_RawUri
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 17 get_SchemeName
        .add_method(|m| m.add_out(value::WinRTType::HString)) // 18 get_UserName
        .add_method(|m| m.add_out(value::WinRTType::I32)) // 19 get_Port
        .add_method(|m| m); // 20 get_Suspicious;
    vtable
}

pub use crate::value::VTableSignature;
pub use crate::value::WinRTType;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn simple_windows_api_should_work() -> Result<()> {
        let doc = XmlDocument::new()?;
        doc.LoadXml(h!("<html>hello world</html>"))?;

        let root = doc.DocumentElement()?;
        assert!(root.NodeName()? == "html");
        assert!(root.InnerText()? == "hello world");

        // unsafe {
        //     let event = CreateEventW(None, true, false, None)?;
        //     SetEvent(event)?;
        //     WaitForSingleObject(event, 0);
        //     CloseHandle(event)?;

        //     MessageBoxA(None, s!("Ansi"), s!("Caption"), MB_OK);
        //     MessageBoxW(None, w!("Wide测试"), w!("Caption"), MB_OK);
        // }

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

    // #[test]
    // fn test_winrt_uri_interop_using_argument() -> Result<()> {
    //     use windows::Foundation::Uri;
    //     println!("URI guid = {:?} ...", Uri::IID);
    //     println!("URI name = {:?} ...", Uri::NAME);

    //     let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;
    //     let mut out_hstr_arg = value::Argument::out_hstring();
    //     let hr = call::call_method_dynamic(17, uri.as_raw(), &mut [&mut out_hstr_arg]);
    //     hr.ok()?;
    //     assert_eq!(out_hstr_arg.as_hstring(), "https");

    //     let mut out_host = value::Argument::out_hstring();
    //     let hr = call::call_method_dynamic(11, uri.as_raw(), &mut [&mut out_host]);
    //     hr.ok()?;

    //     assert_eq!(out_host.as_hstring(), "www.example.com");
    //     Ok(())
    // }

    #[test]
    fn test_winrt_uri_interop_using_signature() -> Result<()> {
        use crate::call;
        use crate::value;
        use windows::Foundation::Uri;

        let uri = Uri::CreateUri(h!("https://www.example.com/path?query=1#fragment"))?;

        // let mut vtable = value::VTableSignature::new();
        // vtable
        //     .add_method(|m| m) // 0 QueryInterface
        //     .add_method(|m| m) // 1 AddRef
        //     .add_method(|m| m) // 2 Release
        //     .add_method(|m| m) // 3 GetIids
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 4 GetRuntimeClassName
        //     .add_method(|m| m) // 5 GetTrustLevel
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 6 get_AbsoluteUri
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 7 get_DisplayUri
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 8 get_Domain
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 9 get_Extension
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 10 get_Fragment
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 11 get_Host
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 12 get_Password
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 13 get_Path
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 14 get_Query
        //     .add_method(|m| m) // 15 get_QueryParsed
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 16 get_RawUri
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 17 get_SchemeName
        //     .add_method(|m| m.add_out(value::WinRTType::HString)) // 18 get_UserName
        //     .add_method(|m| m.add_out(value::WinRTType::I32)) // 19 get_Port
        //     .add_method(|m| m); // 20 get_Suspicious;
        let vtable = uri_vtable();

        let get_runtime_classname = &vtable.methods[4];
        assert_eq!(
            get_runtime_classname.call(uri.as_raw(), &[])?[0].as_hstring(),
            "Windows.Foundation.Uri"
        );

        let get_scheme = &vtable.methods[17];
        let scheme = get_scheme.call(uri.as_raw(), &[])?;
        assert_eq!(scheme[0].as_hstring(), "https");
        let get_path = &vtable.methods[13];
        let path = get_path.call(uri.as_raw(), &[])?;
        assert_eq!(path[0].as_hstring(), "/path");
        let get_port = &vtable.methods[19];
        let port = get_port.call(uri.as_raw(), &[])?;
        assert_eq!(port[0].as_i32(), 443);

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
