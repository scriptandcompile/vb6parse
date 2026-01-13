//! String manipulation statements.
//!
//! This module contains parsers for VB6 statements that manipulate strings:
//! - String alignment (`LSet`, `RSet`)
//! - String replacement (`Mid`, `MidB`)

pub(crate) mod lset;
pub(crate) mod mid;
pub(crate) mod midb;
pub(crate) mod rset;
