use std::ffi::c_void;
use std::sync::atomic::{AtomicUsize, Ordering};

use windows::Win32::Foundation::{CLASS_E_NOAGGREGATION, E_POINTER};
use windows::{
    Win32::System::Com::{IClassFactory, IClassFactory_Impl},
    core::{BOOL, GUID, Ref, implement},
};

static LOCK_COUNT: AtomicUsize = AtomicUsize::new(0);

#[implement(IClassFactory)]
pub struct ClassFactory {
    constructor: fn(*const GUID, *mut *mut c_void) -> windows::core::Result<()>,
}

impl ClassFactory {
    pub fn new(
        constructor: fn(*const GUID, *mut *mut c_void) -> windows::core::Result<()>,
    ) -> Self {
        Self { constructor }
    }

    pub fn can_unload_now() -> bool {
        LOCK_COUNT.load(Ordering::Acquire) == 0
    }
}

impl IClassFactory_Impl for ClassFactory_Impl {
    #[allow(clippy::not_unsafe_ptr_arg_deref)]
    fn CreateInstance(
        &self,
        outer: Ref<'_, windows::core::IUnknown>,
        iid: *const GUID,
        ppv: *mut *mut core::ffi::c_void,
    ) -> windows::core::Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }

        if iid.is_null() {
            return Err(E_POINTER.into());
        }

        if ppv.is_null() {
            return Err(E_POINTER.into());
        }

        (self.constructor)(iid, ppv)
    }

    fn LockServer(&self, flock: BOOL) -> windows::core::Result<()> {
        if flock.as_bool() {
            LOCK_COUNT.fetch_add(1, Ordering::AcqRel);
        } else {
            LOCK_COUNT.fetch_sub(1, Ordering::AcqRel);
        }

        Ok(())
    }
}

#[macro_export]
macro_rules! dll_get_class_object_impl {
    (clsid = $clsid:ident, iid = $iid:ident, ppv = $ppv:ident, classes = [ $($class:ident),* ] ) => {{
        fn __dll_get_class_object_impl(
            clsid: *const GUID,
            iid: *const GUID,
            ppv: *mut *mut c_void,
        ) -> HRESULT {
            use windows::core::{ComObject, Interface, IUnknown};
            use $crate::class_factory::ClassFactory;
            use $crate::com::{CoClass, CreatableCoClass};

            if ppv.is_null() {
                return E_POINTER;
            } else {
                unsafe {
                    ppv.write(std::ptr::null_mut());
                }
            }

            if clsid.is_null() {
                return E_POINTER;
            }

            if iid.is_null() {
                return E_POINTER;
            }

            let class_factory = match unsafe { *clsid } {
                $($class::CLSID => ClassFactory::new(|iid, ppv| unsafe {
                    ComObject::new($class::new()?)
                    .as_interface::<IUnknown>()
                    .query(iid, ppv).ok()
                })),*,
                _ => return CLASS_E_CLASSNOTAVAILABLE,
            };

            #[allow(unreachable_code)]
            unsafe {
                ComObject::new(class_factory)
                    .as_interface::<IUnknown>()
                    .query(iid, ppv)
            }
        }

        __dll_get_class_object_impl($clsid, $iid, $ppv)
    }};
}
