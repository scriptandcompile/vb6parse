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
//! - CreateObject: Creates and returns a reference to an ActiveX object
//! - CurDir: Returns the current path for the specified drive
//! - CVErr: Returns a Variant of subtype Error containing an error number
//! - Date: Returns the current system date
//! - DateAdd: Returns a date to which a specified time interval has been added
//! - DateDiff: Returns the number of time intervals between two dates
//! - DatePart: Returns a specified part of a given date
//! - DateSerial: Returns a date for a specified year, month, and day
//! - DateValue: Returns a date from a string expression
//! - Day: Returns the day of the month (1-31)
//! - DDB: Returns depreciation using the double-declining balance method
//! - Dir: Returns a file or directory name matching a pattern
//! - DoEvents: Yields execution to process other events
//! - Environ: Returns the value of an operating system environment variable
//! - EOF: Returns a Boolean indicating whether the end of a file has been reached
//! - Error: Returns the error message corresponding to a given error number
//! - Exp: Returns e (the base of natural logarithms) raised to a power
//! - FileAttr: Returns the file mode or file handle for an open file
//! - FileDateTime: Returns the date and time when a file was created or last modified
//! - FileLen: Returns the length of a file in bytes
//! - Filter: Returns a zero-based array containing a subset of a string array based on filter criteria
//! - Format: Returns a formatted string expression according to format instructions
//! - FormatCurrency: Returns an expression formatted as a currency value using the system currency symbol
//! - FormatDateTime: Returns an expression formatted as a date or time
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
mod createobject;
mod curdir;
mod cverr;
mod date;
mod dateadd;
mod datediff;
mod datepart;
mod dateserial;
mod datevalue;
mod day;
mod ddb;
mod dir;
mod doevents;
mod environ;
mod eof;
mod error;
mod exp;
mod fileattr;
mod filedatetime;
mod filelen;
mod filter;
mod format;
mod formatcurrency;
mod formatdatetime;
