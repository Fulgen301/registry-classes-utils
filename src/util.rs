use windows::{
    Win32::{
        Foundation::HMODULE,
        System::LibraryLoader::{
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            GetModuleFileNameW, GetModuleHandleExW,
        },
    },
    core::PCWSTR,
};

#[inline(never)]
pub unsafe fn get_this_module_handle() -> windows::core::Result<HMODULE> {
    let mut module = HMODULE::default();

    unsafe {
        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            PCWSTR::from_raw(get_this_module_path as *const () as *const _),
            &raw mut module,
        )?;
    }

    Ok(module)
}

pub fn get_module_path(module: Option<HMODULE>) -> windows::core::Result<Vec<u16>> {
    let mut path = vec![0; 1024];

    loop {
        let size = unsafe { GetModuleFileNameW(module, &mut path) };

        if size == 0 {
            return Err(windows::core::Error::from_win32());
        } else if size as usize != path.len() {
            path.truncate((size as usize) + 1);
            return Ok(path);
        } else {
            path.resize(path.len() * 2, 0);
        }
    }
}

pub unsafe fn get_this_module_path() -> windows::core::Result<Vec<u16>> {
    get_module_path(Some(unsafe { get_this_module_handle()? }))
}
