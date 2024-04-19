use std::ffi::CString;
use windows::{
    Win32::{
        Foundation::{
            HANDLE, HGLOBAL
        },
        System::{
            Memory::{
                GMEM_MOVEABLE,
                GlobalAlloc, GlobalLock, GlobalUnlock
            },
        }
    }
};
use anyhow::Context;

// Set HTML to the clipboard on Windows.
//
// https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
//
// Uses 50 characters for the offsets.
#[allow(dead_code)]
pub fn set_clipboard_html(html: String) -> anyhow::Result<()>{
    let fragment = html;

    let start_html = 391;
    let start_fragment = 454;
    let start_selection = Some(start_fragment);

    let end_fragment = start_fragment + fragment.len() - 1;
    let end_selection = Some(end_fragment);
    // 37 is the `<!-- EndFragment -->\n...`
    let end_html = end_fragment + 37;

    let mut document = String::new();

    // Description

    // Version
    document.push_str(format!("Version:{}\n", "0.9").as_str());

    // StartHTML
    document.push_str(format!("StartHTML:{:0>50}\n", start_html).as_str());
    // EndHTML
    document.push_str(format!("EndHTML:{:0>50}\n", end_html).as_str());

    // StartFragment
    document.push_str(format!("StartFragment:{:0>50}\n", start_fragment).as_str());
    // EndFragment
    document.push_str(format!("EndFragment:{:0>50}\n", end_fragment).as_str());

    if let (Some(start_selection), Some(end_selection)) = (start_selection, end_selection) {
        // StartSelection
        document.push_str(format!("StartSelection:{:0>50}\n", start_selection).as_str());
        // EndSelection
        document.push_str(format!("EndSelection:{:0>50}\n", end_selection).as_str());
    }

    // Context

    document.push_str(
        r#"<!DOCTYPE>
<HTML>
<HEAD>
</HEAD>
<BODY>
<!-- StartFragment -->
"#,
    );
    document.push_str(&fragment);
    document.push_str(
        r#"
<!-- EndFragment -->
</BODY>
</HTML>"#,
    );

    let cstring = CString::new(document).unwrap();
    let cstring = cstring.as_bytes_with_nul();

    // 1. Open Clipboard
    // 2. Empty Clipboard
    // 3. Set Clipboard
    // 4. Close Clipboard

    unsafe {
        // Open Clipboard
        let _open_clipboard = OpenClipboard(None).with_context(|| "open clipboard")?;

        // Empty Clipboard
        let _empty_clipboard = EmptyClipboard().with_context(|| "empty clipboard")?;

        // Register Format
        #[allow(non_snake_case)]
        let CF_HTML;
        let format_name: Vec<u16> = "HTML Format\0".encode_utf16().collect();
        let pcwstr = windows::core::PCWSTR(format_name.as_ptr() as *const u16);
        let register_clipboard_format = RegisterClipboardFormatW(pcwstr);
        CF_HTML = register_clipboard_format;

        // Set Clipboard
        let mem_alloc: HGLOBAL = GlobalAlloc(GMEM_MOVEABLE, cstring.len() * std::mem::size_of::<u16>())?;
        let mem_lock = GlobalLock(mem_alloc);
        std::ptr::copy_nonoverlapping(cstring.as_ptr(), mem_lock as *mut u8, cstring.len());
        let _ = GlobalUnlock(mem_alloc);
        let handle = HANDLE(mem_alloc.0 as isize);

        let _set_clipboard_data = SetClipboardData(CF_HTML, handle).with_context(|| "set clipboard")?;

        // Close Clipboard
        let _close_clipboard = CloseClipboard().with_context(|| "close clipboard")?;
    }

    Ok(())
}
