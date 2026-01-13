//! System and user interaction statements.
//!
//! This module contains parsers for VB6 statements that interact with the system or user:
//! - Application control (`AppActivate`, `Stop`)
//! - User feedback (`Beep`)
//! - UI management (`Load`, `Unload`)
//! - Registry operations (`DeleteSetting`, `SaveSetting`)
//! - Graphics (`SavePicture`)
//! - Input simulation (`SendKeys`)

pub(crate) mod app_activate;
pub(crate) mod beep;
pub(crate) mod delete_setting;
pub(crate) mod load;
pub(crate) mod savepicture;
pub(crate) mod savesetting;
pub(crate) mod sendkeys;
pub(crate) mod stop;
pub(crate) mod unload;
