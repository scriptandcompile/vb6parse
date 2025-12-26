//! VB6 built-in library functions and statements.
//!
//! This module contains parsers for all VB6 built-in functions and statements
//! organized by category to make them easier to discover and maintain.
//!
//! ## Functions
//!
//! VB6 library functions are organized into the following categories:
//! - **`math`** - Mathematical functions (`Abs`, `Sin`, `Cos`, `Sqr`, etc.)
//! - **`string`** - String manipulation and formatting (`Left`, `Right`, `Mid`, `Format`, etc.)
//! - **`conversion`** - Type conversion functions (`Hex`, `Oct`, `CVErr`, `VarType`)
//! - **`datetime`** - Date and time functions (`Date`, `Time`, `DateAdd`, `DateDiff`, etc.)
//! - **`financial`** - Financial calculation functions (`PMT`, `PV`, `FV`, `NPV`, `IRR`, etc.)
//! - **`arrays`** - Array manipulation (`Array`, `Filter`, `Join`, `Split`, `LBound`, `UBound`)
//! - **`file`** - File system functions (`Dir`, `FileLen`, `FileAttr`, etc.)
//! - **`type_checking`** - Type checking functions (`IsArray`, `IsDate`, `IsNumeric`, etc.)
//! - **`interaction`** - User interaction (`MsgBox`, `InputBox`, `Shell`, etc.)
//! - **`objects`** - Object manipulation (`CreateObject`, `GetObject`, `CallByName`, etc.)
//! - **`resources`** - Resource loading (`LoadPicture`, `LoadResString`, etc.)
//! - **`logic`** - Logical/conditional functions (`IIf`, `Choose`, `Switch`)
//! - **`environment`** - Environment and system (`Environ`, `GetSetting`, `SaveSetting`, etc.)
//! - **`graphics`** - Graphics and color functions (`RGB`, `QBColor`)
//!
//! ## Statements
//!
//! VB6 library statements remain in their individual files under the `statements` module.

pub mod functions;
pub mod statements;
