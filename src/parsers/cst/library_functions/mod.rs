//! Built-in VB6 library functions.
//!
//! This module provides documentation and parsing support for VB6's built-in
//! library functions. While these functions are parsed as regular function calls
//! (CallExpression nodes), this module serves to document their behavior and
//! provide reference implementations.
//!
//! The library functions handled here include:
//! - Abs: Returns the absolute value of a number
//! - Array: Returns a Variant containing an array
//! - Asc: Returns the character code of the first letter in a string
//! - Atn: Returns the arctangent of a number in radians
//! - CallByName: Executes a method or accesses a property by name at runtime
//! - Choose: Returns a value from a list of choices based on an index
//! - Chr: Returns the character associated with the specified character code
//! - Command: Returns the command-line arguments passed to the program
//! - Cos: Returns the cosine of an angle in radians
//!
//! Note: Unlike library statements (which are keywords), library functions are
//! called like regular functions and are parsed as CallExpression nodes in the CST.
//! This module primarily serves as documentation and reference for VB6's
//! built-in function library.

mod abs;
mod array;
mod asc;
mod atn;
mod callbyname;
mod choose;
mod chr;
mod command;
mod cos;
