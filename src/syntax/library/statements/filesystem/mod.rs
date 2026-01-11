//! Filesystem operation statements.
//!
//! This module contains parsers for VB6 statements that manipulate the filesystem:
//! - Directory navigation (ChDir, ChDrive)
//! - Directory management (MkDir, RmDir)
//! - File attributes (SetAttr)

pub(crate) mod ch_dir;
pub(crate) mod ch_drive;
pub(crate) mod mkdir;
pub(crate) mod rmdir;
pub(crate) mod setattr;
