use std::ffi::{CString, OsStr};
use std::os::windows::ffi::OsStrExt;
use std::path::PathBuf;

use windows::ApplicationModel::PackageVersion;
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
use windows::core::{PCSTR, PCWSTR};
use windows_core::{HRESULT, HSTRING, HStringBuilder, IUnknown, h};
use windows_future::IAsyncOperation;

use crate::{WinRTValue, bindings};

pub struct WinAppSdkContext;

#[derive(Debug, Clone)]
pub struct WinAppSdkBootstrapOptions {
    pub major_version: u32,
    pub minor_version: u32,
    pub build_version: u32,
    pub revision_version: u32,
    pub bootstrap_dll_path: Option<String>,
}

pub fn initialize_winappsdk(major: u32, minor: u32) -> crate::result::Result<WinAppSdkContext> {
    let options = WinAppSdkBootstrapOptions {
        major_version: major,
        minor_version: minor,
        build_version: 0,
        revision_version: 0,
        bootstrap_dll_path: None,
    };
    initialize(options).map_err(|e| e.into())
}

// use MddBootstrapInitialize from WinAppSDK to initialize the WinAppSDK
pub fn initialize(options: WinAppSdkBootstrapOptions) -> windows::core::Result<WinAppSdkContext> {
    const WINAPPSDK_BOOTSTRAP_DLL_PATH_ENV: &str = "WINAPPSDK_BOOTSTRAP_DLL_PATH";

    let dll_path = HSTRING::from(
        options
            .bootstrap_dll_path
            .or_else(|| std::env::var(WINAPPSDK_BOOTSTRAP_DLL_PATH_ENV).ok())
            .expect("WinAppSDK Bootstrap dll path is requires, set WINAPPSDK_BOOTSTRAP_DLL_PATH env variable or provide in options")
            .to_string(),
    );

    let dp = PCWSTR::from_raw(dll_path.as_ptr());

    let module = unsafe { LoadLibraryW(dp) }?;

    let methodName = CString::new(h!("MddBootstrapInitialize2").to_string()).unwrap();
    let proc = unsafe { GetProcAddress(module, PCSTR::from_raw(methodName.as_ptr() as _)) };
    if proc.is_none() {
        panic!("MddBootstrapInitialize2 not found in bootstrap DLL");
    }

    let init: MddBootstrapInitialize2 = unsafe { std::mem::transmute(proc) };

    // Initialize WinRT first (ok if already initialized in a different mode)
    unsafe { RoInitialize(RO_INIT_MULTITHREADED) }?;

    let major_minor_version = (options.major_version << 16) | options.minor_version;
    let min_version = PackageVersion {
        Major: options.major_version as u16,
        Minor: options.minor_version as u16,
        Build: options.build_version as u16,
        Revision: options.revision_version as u16,
    };

    let hr = unsafe {
        init(
            major_minor_version,
            PCWSTR::from_raw(h!("").as_ptr()),
            min_version,
            0,
        )
    };
    hr.ok()?;
    Ok(WinAppSdkContext {})
}

pub fn find_winappsdk_package(
    major: u32,
    minor: u32,
) -> windows::core::Result<Vec<windows::ApplicationModel::Package>> {
    use windows::Management::Deployment::{PackageManager, PackageTypes};
    use windows_core::HSTRING;

    let manager = PackageManager::new()?;
    let family = format!("Microsoft.WindowsAppRuntime.{major}.{minor}_8wekyb3d8bbwe");
    let packages = manager.FindPackagesByUserSecurityIdPackageFamilyNameWithPackageTypes(
        &HSTRING::default(),
        &HSTRING::from(family),
        PackageTypes::Framework,
    )?;

    let packages: Vec<windows::ApplicationModel::Package> = packages.into_iter().collect();
    Ok(packages)
}

type MddBootstrapInitialize2 =
    unsafe extern "system" fn(u32, PCWSTR, PackageVersion, u32) -> HRESULT;

mod IID {
    use windows_core::{GUID, IUnknown, Interface};

    pub static IID_IUnknown: GUID = GUID::from_u128(0x00000000_0000_0000_c000_000000000046);
    pub static IID_IUnknown2: GUID = IUnknown::IID;
}

// async fn require_send_sync<T: Send + Sync>(_: &T) {}
// fn foo(x: &IUnknown) {
//     // let _ = require_send_sync(x); // compile error :
//     // `NonNull<c_void>` cannot be sent between threads safely
//     // within `IUnknown`, the trait `Send` is not implemented for `NonNull<c_void>`
// }
// fn bar(x: &IAsyncOperation<HSTRING>) {
//     let _ = require_send_sync(x);
// }

pub async fn pick_path() -> crate::result::Result<WinRTValue> {
    use crate::bindings;
    use windows::Web::Http::HttpClient;
    use windows_core::{GUID, IInspectable, Interface};
    use windows_future::{IAsyncInfo, IAsyncOperation};
    let factory = crate::roapi::ro_get_activation_factory_2(h!(
        "Microsoft.Windows.Storage.Pickers.FileOpenPicker"
    ))
    .unwrap();
    let iid = bindings::IFileOpenPickerFactory::IID;
    let fac = factory.cast(&iid).unwrap();
    let picker = fac
        .call_single_out(
            6,
            &crate::WinRTType::Object,
            &[crate::value::WinRTValue::I64(0)],
        )
        .unwrap();
    let picked_file = picker
        .call_single_out(
            13,
            &crate::WinRTType::IAsyncOperation(
                Box::new(crate::WinRTType::RuntimeClass(
                    "Microsoft.Windows.Storage.Pickers.PickFileResult".into(),
                    IAsyncOperation::<bindings::PickFileResult>::IID,
                )),
            ),
            &[],
        )
        .unwrap();

    let res = (&picked_file).await?;
    println!("Picked file result: {:?}", res);
    let path = res
        .call_single_out(6, &crate::WinRTType::HString, &[])
        .unwrap();
    println!("Picked file: {:?}", path);
    Ok(path)
}

pub async fn get_bitmap_from_file() -> WinRTValue {
    use crate::WinRTType;
    use windows::Storage::FileAccessMode;
    use windows::core::Interface;
    let picked = pick_path().await.unwrap().as_hstring().unwrap();
    // let picker = bindings::FileOpenPicker::CreateInstance(Default::default()).unwrap();
    // let pick_result = picker.PickSingleFileAsync().unwrap().await.unwrap();
    // let StorageFileActivationFactory = WinRTValue::from_activation_factory(h!("Windows.Storage.StorageFile")).unwrap();
    // let StorageFile = StorageFileActivationFactory.cast(&windows::Storage::IStorageFileStatics::IID).unwrap();
    // let file_o = StorageFile.call_single_out(6,
    //     &WinRTType::IAsyncOperation(IAsyncOperation::<windows::Storage::StorageFile>::IID),
    //     &[WinRTValue::I64(&picked as *const _ as i64)]
    //     ).unwrap();
    // println!("Getting file from path async...");
    let file = windows::Storage::StorageFile::GetFileFromPathAsync(&picked)
        .unwrap()
        .await
        .unwrap();
    // let file = (&file_o).await.unwrap().as_object().unwrap().cast::<windows::Storage::StorageFile>().unwrap();
    println!("File loaded successfully from path: {}", picked);
    let stream = file.OpenAsync(FileAccessMode::Read).unwrap().await.unwrap();
    println!("File stream opened successfully");
    let decoder = windows::Graphics::Imaging::BitmapDecoder::CreateAsync(&stream)
        .unwrap()
        .await
        .unwrap();
    println!("BitmapDecoder created successfully");
    let bitmap = decoder.GetSoftwareBitmapAsync().unwrap().await.unwrap();
    println!("SoftwareBitmap obtained successfully");
    let raw = bitmap.as_raw();
    let result = WinRTValue::Object(unsafe { IUnknown::from_raw(raw) });
    std::mem::forget(bitmap);
    result
}



pub async fn test_pick_open_picker_full_dynamic() -> crate::result::Result<WinRTValue> {
    let _ = initialize_winappsdk(1, 8).unwrap();
    pick_path().await
}

#[cfg(test)]
mod tests {
    use windows::{Foundation::Uri, Web::Http::HttpClient};
    use windows_core::{GUID, IInspectable, Interface};
    use windows_future::{IAsyncInfo, IAsyncOperation};

    use crate::{bindings, interfaces};

    use super::*;

    #[test]
    fn use_windows_rs_find_package() -> windows::core::Result<()> {
        let packages = find_winappsdk_package(1, 8)?;
        assert!(!packages.is_empty());
        packages.into_iter().for_each(|pkg| {
            let name = pkg.DisplayName().unwrap().to_string();
            println!("Found package: {}", name);
        });
        Ok(())
    }

    #[test]
    fn test_initialize() {
        let options = WinAppSdkBootstrapOptions {
            major_version: 1,
            minor_version: 8,
            build_version: 0,
            revision_version: 0,
            bootstrap_dll_path: None,
        };
        let result = initialize(options);
        assert!(result.is_ok());
    }

    #[repr(transparent)]
    #[derive(Clone, PartialEq, Eq, Debug)]
    struct PickFileResult(IUnknown);

    #[test]
    fn use_bindings() {}

    struct IAsyncOperationFilePickResultWrapper(IUnknown);

    #[tokio::test]
    async fn test_http_get_string() -> windows_core::Result<()> {
        let s = crate::http_get_string("https://www.microsoft.com").await?;
        println!("HTTP GET result: {}", s);
        Ok(())
    }

    #[tokio::test]
    async fn test_pick_open_picker() -> windows_core::Result<()> {
        let options = WinAppSdkBootstrapOptions {
            major_version: 1,
            minor_version: 8,
            build_version: 0,
            revision_version: 0,
            bootstrap_dll_path: None,
        };
        let result = initialize(options);
        assert!(
            result.is_ok(),
            "message {}",
            result.err().unwrap().message()
        );
        let factory = crate::roapi::ro_get_activation_factory(h!(
            "Microsoft.Windows.Storage.Pickers.FileOpenPicker"
        ));
        assert!(factory.is_ok());
        let iid = GUID::from_values(
            828278487,
            55202,
            23937,
            [179, 121, 122, 247, 130, 7, 177, 175],
        );
        let mut fPtr = std::ptr::null_mut();
        unsafe { factory?.query(&iid, &mut fPtr) }.ok()?;
        let fac = unsafe { windows_core::IUnknown::from_raw(fPtr) };

        // IAsyncOperation<PickFileResult>
        // IAsyncOperation<HSTRING> //
        // IVector<PickFileResult>
        // let a: IAsyncOperation<HSTRING> = unimplemented!();
        // let c: HttpClient = HttpClient::new()?;
        // let u: windows::Foundation::Uri = unimplemented!();

        let factoryInterface = interfaces::FileOpenPickerFactory();
        let result = factoryInterface.methods[6]
            .call_dynamic(fac.as_raw(), &[crate::value::WinRTValue::I64(0)])?;
        let rv1 = &result[0].as_object().unwrap();
        let pickerInterface = interfaces::FileOpenPicker();
        let result = pickerInterface.methods[13].call_dynamic(
            rv1.as_raw(),
            &[], // No parameters
        )?;
        let rv2 = &result[0].as_object().unwrap();

        // let asv: IAsyncOperation<bindings::PickFileResult> = rv2.cast().unwrap();
        // let rr = asv.join()?;
        let rva: IAsyncInfo = rv2.cast().unwrap();
        let iid = IAsyncOperation::<bindings::PickFileResult>::IID;
        let op = crate::dasync::DynWinRTAsyncOperationIUnknown(rva, iid);
        let res = op.await?;
        println!("Picked file result: {:?}", res);
        let pfrvtbl = interfaces::PickFileResult();
        let path_results = pfrvtbl.methods[6].call_dynamic(res.as_raw(), &[])?;
        let path = path_results[0].as_hstring().unwrap();
        // let mut ptr = std::ptr::null_mut();
        // unsafe { res.query(&bindings::PickFileResult::IID, &mut ptr) }.unwrap();
        // // let r: bindings::PickFileResult = unsafe { bindings::PickFileResult::from_raw(ptr as _) };
        // let r: bindings::PickFileResult = res.cast()?;
        // assert_eq!(r.as_raw(), unsafe { ptr });
        println!("Picked file: {:?}", path);
        // println!("Picked file: {:?}", r.Path());
        Ok(())
    }

    #[tokio::test]
    async fn test_pick_open_picker_full_dynamic_wrapper() -> crate::result::Result<()> {
        initialize_winappsdk(1, 8).unwrap();
        let path = test_pick_open_picker_full_dynamic().await.unwrap();
        println!("Picked path: {:?}", path.as_hstring().unwrap());
        Ok(())
    }
}
