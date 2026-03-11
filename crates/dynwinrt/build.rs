#[test]
fn bind_gen_test() {}

fn main() {
    // println!("building build.rs...");
    // let packages_folder = std::env::var("NUGET_PACKAGES").ok().map(std::path::PathBuf::from).unwrap_or_else(|| {
    //     let userprofile = std::env::var("USERPROFILE").unwrap_or_else(|_| String::new());
    //     if userprofile.is_empty() {
    //         std::path::PathBuf::from(r"C:\\Users\\Default\\.nuget\\packages")
    //     } else {
    //         std::path::Path::new(&userprofile).join(r".nuget\\packages")
    //     }
    // });

    // let winappsdk_version = std::env::var("WINAPPSDK_VERSION").ok();

    // fn version_key(s: &str) -> Vec<u64> {
    //     s.split('.')
    //         .map(|part| {
    //             let digits: String = part.chars().take_while(|c| c.is_ascii_digit()).collect();
    //             digits.parse::<u64>().unwrap_or(0)
    //         })
    //         .collect()
    // }

    // fn list_versions(packages_folder: &std::path::Path, package_id: &str) -> Vec<String> {
    //     let root = packages_folder.join(package_id);
    //     let mut versions = Vec::new();
    //     let Ok(entries) = std::fs::read_dir(root) else {
    //         return versions;
    //     };
    //     for entry in entries.flatten() {
    //         if let Ok(ft) = entry.file_type() {
    //             if ft.is_dir() {
    //                 versions.push(entry.file_name().to_string_lossy().into_owned());
    //             }
    //         }
    //     }
    //     versions.sort_by(|a, b| version_key(b).cmp(&version_key(a)));
    //     versions
    // }

    // fn pick_version(packages_folder: &std::path::Path, package_id: &str, preferred: Option<&str>) -> (String, bool) {
    //     if let Some(preferred) = preferred {
    //         if packages_folder.join(package_id).join(preferred).exists() {
    //             return (preferred.to_string(), false);
    //         }
    //     }

    //     let versions = list_versions(packages_folder, package_id);
    //     let Some(best) = versions.first() else {
    //         panic!(
    //             "NuGet package '{}' not found under {}",
    //             package_id,
    //             packages_folder.to_string_lossy()
    //         );
    //     };
    //     (best.clone(), true)
    // }

    // let pkg = |relative: &str| packages_folder.join(relative).to_string_lossy().into_owned();

    // let (foundation_version, foundation_fallback) = pick_version(
    //     &packages_folder,
    //     "microsoft.windowsappsdk.foundation",
    //     winappsdk_version.as_deref(),
    // );
    // if foundation_fallback {
    //     if let Some(requested) = winappsdk_version.as_deref() {
    //         println!(
    //             "cargo:warning=WinAppSDK metadata package microsoft.windowsappsdk.foundation does not contain version {}; using {} instead.",
    //             requested, foundation_version
    //         );
    //     }
    // }

    // let (interactive_version, interactive_fallback) = pick_version(
    //     &packages_folder,
    //     "microsoft.windowsappsdk.interactiveexperiences",
    //     winappsdk_version.as_deref(),
    // );
    // if interactive_fallback {
    //     if let Some(requested) = winappsdk_version.as_deref() {
    //         println!(
    //             "cargo:warning=WinAppSDK metadata package microsoft.windowsappsdk.interactiveexperiences does not contain version {}; using {} instead.",
    //             requested, interactive_version
    //         );
    //     }
    // }

    // let metadata_sdk_version = std::env::var("WINAPPSDK_METADATA_SDK_VERSION").unwrap_or_else(|_| "10.0.18362.0".to_string());

    // let pickers_winmd = pkg(&format!(
    //     "microsoft.windowsappsdk.foundation/{}/metadata/Microsoft.Windows.Storage.Pickers.winmd",
    //     foundation_version
    // ));

    // let ui_winmd_primary = pkg(&format!(
    //     "microsoft.windowsappsdk.interactiveexperiences/{}/metadata/{}/Microsoft.UI.winmd",
    //     interactive_version, metadata_sdk_version
    // ));
    // let foundation_winmd_primary = pkg(&format!(
    //     "microsoft.windowsappsdk.interactiveexperiences/{}/metadata/{}/Microsoft.Foundation.winmd",
    //     interactive_version, metadata_sdk_version
    // ));

    // let ui_winmd_fallback = pkg(&format!(
    //     "microsoft.windowsappsdk.interactiveexperiences/{}/metadata/10.0.17763.0/Microsoft.UI.winmd",
    //     interactive_version
    // ));
    // let foundation_winmd_fallback = pkg(&format!(
    //     "microsoft.windowsappsdk.interactiveexperiences/{}/metadata/10.0.17763.0/Microsoft.Foundation.winmd",
    //     interactive_version
    // ));

    // let ui_winmd = if std::path::Path::new(&ui_winmd_primary).exists() {
    //     ui_winmd_primary
    // } else {
    //     ui_winmd_fallback
    // };
    // let foundation_winmd = if std::path::Path::new(&foundation_winmd_primary).exists() {
    //     foundation_winmd_primary
    // } else {
    //     foundation_winmd_fallback
    // };

    // let ai_winmd = pkg("microsoft.windowsappsdk.ai/1.8.44/metadata/Microsoft.Windows.AI.winmd");

    // let args: Vec<String> = vec![
       
    //     "--out".into(),
    //     "src/bindings.rs".into(),
    //     "--in".into(),
    //     r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd".into(),
    //     // "--in".into(),
    //     // ui_winmd,
    //     // "--in".into(),
    //     // foundation_winmd,
    //     "--in".into(),
    //     ai_winmd,
    //      "--in".into(),
    //     pickers_winmd,
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.interactiveexperiences\\1.8.251217001\\metadata\\10.0.18362.0\\Microsoft.UI.winmd".into(),
    //     "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.interactiveexperiences\\1.8.251217001\\metadata\\10.0.18362.0\\Microsoft.Foundation.winmd".into(),
    //     "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.interactiveexperiences\\1.8.251217001\\metadata\\10.0.18362.0\\Microsoft.Graphics.winmd".into(),
    //     "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.ai\\1.8.44\\metadata\\Microsoft.Windows.AI.winmd".into(),
    //     "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.ai\\1.8.44\\metadata\\Microsoft.Windows.AI.Imaging.winmd".into(),
    //     "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.ai\\1.8.44\\metadata\\Microsoft.Graphics.Imaging.winmd".into(),
    //      "--in".into(),
    //     "C:\\Users\\xianghong\\.nuget\\packages\\microsoft.windowsappsdk.ai\\1.8.44\\metadata\\Microsoft.Windows.AI.ContentSafety.winmd".into(),       
    //     "--flat".into(),
    //     "--filter".into(),
    //     "Microsoft.Windows.Storage.Pickers.*".into(),
    //     "--filter".into(),
    //     "Microsoft.Windows.AI.*".into(),
    //     "--filter".into(),
    //     "Microsoft.Windows.AI.Imaging.*".into(),
    //     "--filter".into(),
    //     "Microsoft.UI.WindowId".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.*".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.Imaging.*".into(),
    //     "--filter".into(),
    //     "Microsoft.Graphics.Imaging.*".into(),
    //     "--filter".into(),
    //     "Windows.Storage.Streams.InputStreamOptions".into(),
    //     "--filter".into(),
    //     "Windows.Storage.Streams.IBuffer".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.DirectX.Direct3D11.IDirect3DSurface".into(),
    //     "--filter".into(),
    //     "Windows.Foundation.IMemoryBufferReference".into(),
    //     "--filter".into(),
    //     "Windows.Foundation.PropertyType".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.DirectX.DirectXPixelFormat".into(),
    //     "--filter".into(),
    //     "Microsoft.Windows.AI.ContentSafety.ContentFilterOptions".into(),
    //     "--filter".into(),
    //     "Microsoft.Windows.AI.ContentSafety.*".into(),
    //     "--filter".into(),
    //     "Windows.Foundation.TypedEventHandler".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.DirectX.Direct3D11.Direct3DSurfaceDescription".into(),
    //     "--filter".into(),
    //     "Windows.Graphics.DirectX.Direct3D11.Direct3DMultisampleDescription".into(),
    // ];

    // windows_bindgen::bindgen(args).unwrap();
}
