use std::ffi::CString;

use windows::ApplicationModel::PackageVersion;
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
use windows::Win32::System::WinRT::{RO_INIT_MULTITHREADED, RoInitialize};
use windows::core::PCSTR;
use windows::core::PCWSTR;
use windows_core::{HRESULT, HSTRING, h};

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

    let method_name = CString::new(h!("MddBootstrapInitialize2").to_string()).unwrap();
    let proc = unsafe { GetProcAddress(module, PCSTR::from_raw(method_name.as_ptr() as _)) };
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

#[allow(dead_code)]
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
