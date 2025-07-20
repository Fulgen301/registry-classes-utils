use std::ops::Deref;

use transaction::Key;
use windows::core::{PCWSTR, w};

use crate::com::{CoClass, GuidExt};

pub mod transaction {
    use std::cell::Cell;

    use windows::{
        Win32::{
            Foundation::{
                E_ILLEGAL_STATE_CHANGE, ERROR_FILE_NOT_FOUND, ERROR_SUCCESS, HANDLE, WIN32_ERROR,
            },
            Storage::FileSystem::{CommitTransaction, CreateTransaction, RollbackTransaction},
            System::{
                Registry::{
                    HKEY, KEY_READ, KEY_WRITE, REG_BINARY, REG_DWORD, REG_EXPAND_SZ,
                    REG_OPEN_CREATE_OPTIONS, REG_OPTION_NON_VOLATILE, REG_OPTION_VOLATILE,
                    REG_QWORD, REG_SZ, REG_VALUE_TYPE, RegCreateKeyTransactedW, RegDeleteTreeW,
                    RegDeleteValueW, RegOpenKeyTransactedW,
                },
                Threading::INFINITE,
            },
        },
        core::{GUID, Owned, PCWSTR},
    };

    use crate::com::GuidExt;

    pub struct Transaction {
        handle: Owned<HANDLE>,
        key_options: REG_OPEN_CREATE_OPTIONS,
        committed: Cell<bool>,
    }

    impl Transaction {
        pub fn new(description: PCWSTR, volatile: bool) -> windows::core::Result<Self> {
            Ok(Self {
                handle: unsafe {
                    Owned::new(CreateTransaction(
                        std::ptr::null_mut(),
                        std::ptr::null_mut(),
                        0,
                        0,
                        0,
                        INFINITE,
                        description,
                    )?)
                },
                key_options: if volatile {
                    REG_OPTION_VOLATILE
                } else {
                    REG_OPTION_NON_VOLATILE
                },

                committed: Cell::new(false),
            })
        }

        pub fn commit(&self) -> windows::core::Result<()> {
            if self.committed.get() {
                return Err(E_ILLEGAL_STATE_CHANGE.into());
            }

            unsafe {
                CommitTransaction(*self.handle)?;
            }

            self.committed.replace(true);
            Ok(())
        }
    }

    impl Drop for Transaction {
        fn drop(&mut self) {
            if !self.committed.get() {
                unsafe {
                    let _ = RollbackTransaction(*self.handle);
                }
            }
        }
    }

    unsafe fn reg_create_key_transacted(
        key: HKEY,
        sub_key: PCWSTR,
        options: REG_OPEN_CREATE_OPTIONS,
        transaction: HANDLE,
    ) -> windows::core::Result<HKEY> {
        let mut result = HKEY::default();

        unsafe {
            RegCreateKeyTransactedW(
                key,
                sub_key,
                None,
                None,
                options,
                KEY_READ | KEY_WRITE,
                None,
                &raw mut result,
                None,
                transaction,
                None,
            )
            .ok()?;
        }

        Ok(result)
    }

    #[allow(unused)]
    unsafe fn open_key_transacted(
        key: HKEY,
        sub_key: PCWSTR,
        transaction: HANDLE,
    ) -> windows::core::Result<HKEY> {
        let mut result = HKEY::default();

        unsafe {
            RegOpenKeyTransactedW(
                key,
                sub_key,
                None,
                KEY_READ | KEY_WRITE,
                &raw mut result,
                transaction,
                None,
            )
            .ok()?;
        }

        Ok(result)
    }

    pub struct Key<'a> {
        transaction: &'a Transaction,
        key: Owned<HKEY>,
    }

    impl<'a> Key<'a> {
        pub fn predefined(
            transaction: &'a Transaction,
            key: HKEY,
            sub_key: PCWSTR,
        ) -> windows::core::Result<Self> {
            let mut result = HKEY::default();

            unsafe {
                RegCreateKeyTransactedW(
                    key,
                    sub_key,
                    None,
                    None,
                    transaction.key_options,
                    KEY_READ | KEY_WRITE,
                    None,
                    &raw mut result,
                    None,
                    *transaction.handle,
                    None,
                )
                .ok()?;
            }

            Ok(Self {
                transaction,
                key: unsafe {
                    Owned::new(reg_create_key_transacted(
                        key,
                        sub_key,
                        transaction.key_options,
                        *transaction.handle,
                    )?)
                },
            })
        }

        pub fn create_subkey(&self, sub_key: PCWSTR) -> windows::core::Result<Key<'a>> {
            Ok(Self {
                transaction: self.transaction,
                key: unsafe {
                    Owned::new(reg_create_key_transacted(
                        *self.key,
                        sub_key,
                        self.transaction.key_options,
                        *self.transaction.handle,
                    )?)
                },
            })
        }

        #[allow(unused)]
        pub fn open_subkey(&self, sub_key: PCWSTR) -> windows::core::Result<Key<'a>> {
            Ok(Self {
                transaction: self.transaction,
                key: unsafe {
                    Owned::new(open_key_transacted(
                        *self.key,
                        sub_key,
                        *self.transaction.handle,
                    )?)
                },
            })
        }

        pub fn delete_subkey(&self, subkey: PCWSTR) -> windows::core::Result<()> {
            self.delete_tree_internal(subkey)
        }

        pub fn delete_tree(&self) -> windows::core::Result<()> {
            self.delete_tree_internal(PCWSTR::null())
        }

        fn delete_tree_internal(&self, subkey: PCWSTR) -> windows::core::Result<()> {
            match unsafe { RegDeleteTreeW(*self.key, subkey) } {
                ERROR_SUCCESS | ERROR_FILE_NOT_FOUND => Ok(()),
                e => e.ok(),
            }
        }

        pub fn set_u32(&self, name: PCWSTR, value: u32) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_le_bytes()), REG_DWORD)
        }

        #[allow(unused)]
        pub fn set_u64(&self, name: PCWSTR, value: u64) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_le_bytes()), REG_QWORD)
        }

        pub fn set_binary(&self, name: PCWSTR, value: &[u8]) -> windows::core::Result<()> {
            self.set_value(name, Some(value), REG_BINARY)
        }

        #[allow(unused)]
        pub fn set_str(&self, name: PCWSTR, value: &str) -> windows::core::Result<()> {
            self.set_value(
                name,
                Some(&value.encode_utf16().collect::<Vec<_>>()),
                REG_SZ,
            )
        }

        #[allow(unused)]
        pub fn set_str_expand(&self, name: PCWSTR, value: &str) -> windows::core::Result<()> {
            self.set_value(
                name,
                Some(&value.encode_utf16().collect::<Vec<_>>()),
                REG_EXPAND_SZ,
            )
        }

        pub fn set_pcwstr(&self, name: PCWSTR, value: PCWSTR) -> windows::core::Result<()> {
            self.set_value(
                name,
                if value.is_null() {
                    None
                } else {
                    Some(unsafe { value.as_wide() })
                },
                REG_SZ,
            )
        }

        pub fn set_pcwstr_expand(&self, name: PCWSTR, value: PCWSTR) -> windows::core::Result<()> {
            self.set_value(
                name,
                if value.is_null() {
                    None
                } else {
                    Some(unsafe { value.as_wide() })
                },
                REG_EXPAND_SZ,
            )
        }

        pub fn set_guid(&self, name: PCWSTR, value: &GUID) -> windows::core::Result<()> {
            self.set_value(name, Some(&value.to_wide()), REG_SZ)
        }

        fn set_value<T>(
            &self,
            name: PCWSTR,
            value: Option<&[T]>,
            value_type: REG_VALUE_TYPE,
        ) -> windows::core::Result<()> {
            unsafe extern "system" {
                #[allow(unused)]
                fn RegSetValueExW(
                    hkey: HKEY,
                    lpvaluename: PCWSTR,
                    reserved: u32,
                    dwtype: REG_VALUE_TYPE,
                    lpdata: *const u8,
                    cbdata: u32,
                ) -> WIN32_ERROR;
            }

            unsafe {
                RegSetValueExW(
                    *self.key,
                    name,
                    0,
                    value_type,
                    value.map_or(std::ptr::null(), |v| v.as_ptr().cast()),
                    (value.map_or(0, |v| v.len()) * std::mem::size_of::<T>()) as u32,
                )
                .ok()
            }
        }

        pub fn delete_value(&self, name: PCWSTR) -> windows::core::Result<()> {
            match unsafe { RegDeleteValueW(*self.key, name) } {
                ERROR_SUCCESS | ERROR_FILE_NOT_FOUND => Ok(()),
                e => e.ok(),
            }
        }
    }
}

#[derive(Clone, Copy)]
pub struct NullTerminatedSlice<'a>(&'a [u16]);

impl<'a> NullTerminatedSlice<'a> {
    pub fn new(slice: &'a [u16]) -> Option<Self> {
        if slice.last() != Some(&0u16) {
            None
        } else {
            Some(Self(slice))
        }
    }
}

impl Deref for NullTerminatedSlice<'_> {
    type Target = [u16];

    fn deref(&self) -> &Self::Target {
        self.0
    }
}

pub fn register_com_extension<'a, T: CoClass>(
    classes: &'a Key,
    module_path: NullTerminatedSlice,
    description: PCWSTR,
    apartment_type: PCWSTR,
) -> windows::core::Result<Key<'a>> {
    let clsid_string = T::CLSID.to_wide();
    let com_object = classes
        .create_subkey(w!("CLSID"))?
        .create_subkey(PCWSTR::from_raw(clsid_string.as_ptr()))?;

    com_object.set_pcwstr(PCWSTR::null(), description)?;

    com_object
        .create_subkey(w!("ProgId"))?
        .set_pcwstr(PCWSTR::null(), T::PROG_ID)?;

    com_object
        .create_subkey(w!("VersionIndependentProgId"))?
        .set_pcwstr(PCWSTR::null(), T::VERSION_INDEPENDENT_PROG_ID)?;

    let inproc = com_object.create_subkey(w!("InprocServer32"))?;
    inproc.set_pcwstr(PCWSTR::null(), PCWSTR::from_raw(module_path.as_ptr()))?;
    inproc.set_pcwstr(w!("ThreadingModel"), apartment_type)?;

    classes
        .create_subkey(T::PROG_ID)?
        .create_subkey(w!("CLSID"))?
        .set_guid(PCWSTR::null(), &T::CLSID)?;

    classes
        .create_subkey(T::VERSION_INDEPENDENT_PROG_ID)?
        .create_subkey(w!("CLSID"))?
        .set_guid(PCWSTR::null(), &T::CLSID)?;

    Ok(com_object)
}

pub fn unregister_com_extension<T: CoClass>(classes: &Key) -> windows::core::Result<()> {
    let mut buffer = [0u16; 39 + 6];
    unsafe {
        buffer[..6]
            .as_mut_ptr()
            .copy_from_nonoverlapping(w!("CLSID\\").as_ptr(), 6);
    }

    let clsid_string = T::CLSID.to_wide();
    buffer[6..].copy_from_slice(&clsid_string);
    classes.delete_subkey(PCWSTR::from_raw(buffer.as_ptr()))?;

    classes.delete_subkey(T::PROG_ID)?;
    classes.delete_subkey(T::VERSION_INDEPENDENT_PROG_ID)?;
    Ok(())
}
