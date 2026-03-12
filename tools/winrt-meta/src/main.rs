use std::collections::HashSet;
use std::fs;
use std::path::Path;

use clap::{Parser, Subcommand};

use winrt_meta::codegen::typescript;
use winrt_meta::meta;
use winrt_meta::types::TypeMeta;

#[derive(Parser)]
#[command(name = "winrt-meta")]
#[command(about = "Generate language bindings from WinRT metadata (.winmd) files")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Generate TypeScript or Python bindings from .winmd files
    Generate {
        /// Path(s) to .winmd metadata files, separated by ';'
        #[arg(long)]
        winmd: String,

        /// Filter by namespace (e.g. "Windows.Foundation")
        #[arg(long)]
        namespace: Option<String>,

        /// Generate bindings for a specific class only
        #[arg(long, name = "class")]
        class_name: Option<String>,

        /// Target language: "ts" (TypeScript) or "py" (Python)
        #[arg(long, default_value = "ts")]
        lang: String,

        /// Output directory for generated files
        #[arg(long, default_value = "./generated")]
        output: String,
    },
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Commands::Generate {
            winmd,
            namespace,
            class_name,
            lang,
            output,
        } => {
            if lang != "ts" {
                eprintln!("Only TypeScript (ts) generation is supported currently.");
                std::process::exit(1);
            }

            // Auto-detect Windows SDK Windows.winmd and append if not already included
            let winmd = if winmd.contains("Windows.winmd") {
                winmd
            } else if let Some(sdk_winmd) = find_windows_sdk_winmd() {
                eprintln!("Auto-detected Windows SDK: {}", sdk_winmd);
                format!("{};{}", winmd, sdk_winmd)
            } else {
                winmd
            };

            let output_dir = Path::new(&output);
            fs::create_dir_all(output_dir).expect("Failed to create output directory");

            let classes = if let Some(ref cls) = class_name {
                let ns = namespace
                    .as_deref()
                    .expect("--namespace is required when --class is specified");
                match meta::parse_class(&winmd, ns, cls) {
                    Some(c) => vec![c],
                    None => {
                        eprintln!("Class {}.{} not found in {}", ns, cls, winmd);
                        std::process::exit(1);
                    }
                }
            } else if let Some(ref ns) = namespace {
                meta::parse_namespace(&winmd, ns)
            } else {
                eprintln!("Either --namespace or --class (with --namespace) must be specified.");
                std::process::exit(1);
            };

            let interfaces = if let Some(ref ns) = namespace {
                meta::parse_interfaces(&winmd, ns)
            } else {
                Vec::new()
            };

            let enums = if let Some(ref ns) = namespace {
                meta::parse_enums(&winmd, ns)
            } else {
                Vec::new()
            };

            // Resolve transitive dependencies (recursively discovers all referenced types)
            let deps = meta::resolve_dependencies(&winmd, &classes, &interfaces, &enums);
            let mut all_classes = classes;
            let mut all_interfaces = interfaces;
            let mut all_enums = enums;
            all_classes.extend(deps.classes);
            all_interfaces.extend(deps.interfaces);
            all_enums.extend(deps.enums);

            // Build set of known type names (types that will have generated .ts files)
            let mut known_types: HashSet<String> = HashSet::new();
            for c in &all_classes { known_types.insert(c.name.clone()); }
            for i in &all_interfaces { known_types.insert(i.name.clone()); }
            for e in &all_enums {
                if let TypeMeta::Enum { name, .. } = e { known_types.insert(name.clone()); }
            }

            // Generate interface files
            for iface in &all_interfaces {
                let ts_code = typescript::generate_interface(iface, &known_types);
                let filename = format!("{}.ts", iface.name);
                let filepath = output_dir.join(&filename);
                fs::write(&filepath, &ts_code).expect("Failed to write generated file");
                println!("Generated {}", filepath.display());
            }

            // Generate enum files
            for en in &all_enums {
                if let TypeMeta::Enum { name, .. } = en {
                    if let Some(ts_code) = typescript::generate_enum(en) {
                        let filename = format!("{}.ts", name);
                        let filepath = output_dir.join(&filename);
                        fs::write(&filepath, &ts_code).expect("Failed to write generated file");
                        println!("Generated {}", filepath.display());
                    }
                }
            }

            // Generate class files
            for class in &all_classes {
                let ts_code = typescript::generate_class(class, &known_types);
                let filename = format!("{}.ts", class.name);
                let filepath = output_dir.join(&filename);
                fs::write(&filepath, &ts_code).expect("Failed to write generated file");
                println!("Generated {}", filepath.display());
            }



            // Generate index.ts
            let total = all_classes.len() + all_interfaces.len() + all_enums.len();
            if namespace.is_some() && total > 1 {
                let index_code = typescript::generate_index(&all_classes, &all_interfaces, &all_enums);
                let index_path = output_dir.join("index.ts");
                fs::write(&index_path, &index_code).expect("Failed to write index.ts");
                println!("Generated {}", index_path.display());
            }

            println!(
                "Done. {} class(es) + {} interface(s) + {} enum(s) generated in {}",
                all_classes.len(),
                all_interfaces.len(),
                all_enums.len(),
                output_dir.display()
            );
        }
    }
}

/// Find Windows SDK Windows.winmd by scanning the standard install location.
/// Returns the path to the latest version's Windows.winmd, if found.
fn find_windows_sdk_winmd() -> Option<String> {
    let base = Path::new(r"C:\Program Files (x86)\Windows Kits\10\UnionMetadata");
    if !base.exists() {
        return None;
    }
    // Find all version directories and pick the latest (lexicographic sort works for version strings)
    let mut versions: Vec<_> = fs::read_dir(base)
        .ok()?
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_dir())
        .map(|e| e.file_name().to_string_lossy().to_string())
        .filter(|name| name.starts_with("10.")) // Only version directories like "10.0.26100.0"
        .collect();
    versions.sort();
    // Check from latest to oldest
    for version in versions.iter().rev() {
        let winmd_path = base.join(version).join("Windows.winmd");
        if winmd_path.exists() {
            return Some(winmd_path.to_string_lossy().to_string());
        }
    }
    None
}
