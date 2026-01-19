fn main() {
    // Only run Windows resource compilation on Windows
    #[cfg(target_os = "windows")]
    {
        let mut res = winresource::WindowsResource::new();
        res.set_icon("windows/visigrid.ico");
        // Don't set manifest - gpui already provides one
        res.set("FileDescription", "VisiGrid - A fast, keyboard-driven spreadsheet");
        res.set("ProductName", "VisiGrid");
        res.set("CompanyName", "RegAtlas, LLC");
        res.set("LegalCopyright", "Copyright © 2025 RegAtlas, LLC");
        res.compile().expect("Failed to compile Windows resources");
    }
}
