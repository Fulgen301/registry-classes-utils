use std::{
    fmt::Display,
    io::{Cursor, Write},
};

use windows::core::{GUID, PCWSTR};

pub trait CoClass {
    const CLSID: GUID;
    const PROG_ID: PCWSTR;
    const VERSION_INDEPENDENT_PROG_ID: PCWSTR;
}

pub trait CreatableCoClass: CoClass + Sized {
    fn new() -> windows::core::Result<Self>;
}

struct GuidWrapper<'a>(&'a GUID);

impl Display for GuidWrapper<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{{{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}}}",
            self.0.data1,
            self.0.data2,
            self.0.data3,
            self.0.data4[0],
            self.0.data4[1],
            self.0.data4[2],
            self.0.data4[3],
            self.0.data4[4],
            self.0.data4[5],
            self.0.data4[6],
            self.0.data4[7]
        )
    }
}

pub trait GuidExt {
    fn to_ascii_with_nul(&self) -> [u8; 39];
    fn to_wide(&self) -> [u16; 39] {
        self.to_ascii_with_nul().map(|value| value as u16)
    }
}

impl GuidExt for GUID {
    fn to_ascii_with_nul(&self) -> [u8; 39] {
        let mut cursor = Cursor::new([0u8; 39]);
        write!(cursor, "{}", GuidWrapper(self)).unwrap();
        assert!(cursor.position() == 38);
        cursor.into_inner()
    }
}
