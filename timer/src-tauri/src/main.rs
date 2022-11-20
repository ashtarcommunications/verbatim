#![cfg_attr(
  all(not(debug_assertions), target_os = "windows"),
  windows_subsystem = "windows"
)]

use tauri_plugin_store::{PluginBuilder, StoreBuilder};

fn main() {
  let settings = StoreBuilder::new(".settings".parse().unwrap())
    .default("the-key".to_string(), "wooooot".into())
    .build();

  tauri::Builder::default()
    .plugin(PluginBuilder::default().stores([settings]).freeze().build())
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}
