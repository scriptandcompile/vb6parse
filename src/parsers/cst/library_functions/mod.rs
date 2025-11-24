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
//! - Fix: Returns the integer portion of a number
//! - Format: Returns a formatted string expression according to format instructions
//! - FormatCurrency: Returns an expression formatted as a currency value using the system currency symbol
//! - FormatDateTime: Returns an expression formatted as a date or time
//! - FormatNumber: Returns an expression formatted as a number
//! - FormatPercent: Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character
//! - FreeFile: Returns the next file number available for use by the Open statement
//! - Fv: Returns the future value of an annuity based on periodic, fixed payments and a fixed interest rate
//! - GetAllSettings: Returns a list of key settings and their values from the Windows registry
//! - GetAttr: Returns an Integer representing the attributes of a file, directory, or folder
//! - GetAutoServerSettings: Returns information about the security settings for a DCOM server
//! - GetObject: Returns a reference to an ActiveX object from a file or a running instance
//! - GetSetting: Returns a registry key setting value from the Windows registry
//! - Hex: Returns a String representing the hexadecimal value of a number
//! - Hour: Returns an Integer specifying a whole number between 0 and 23, inclusive, representing the hour of the day
//! - IIf: Returns one of two parts, depending on the evaluation of an expression
//! - IMEStatus: Returns an Integer indicating the current Input Method Editor (IME) mode of Microsoft Windows
//! - Input: Returns String containing characters from a file opened in Input or Binary mode
//! - InputBox: Displays a prompt in a dialog box, waits for user input, and returns a String
//! - InStr: Returns a Long specifying the position of the first occurrence of one string within another
//! - InStrRev: Returns the position of an occurrence of one string within another, from the end of string
//! - Int: Returns the integer portion of a number
//! - IPmt: Returns the interest payment for a given period of an annuity
//! - IRR: Returns the internal rate of return for a series of periodic cash flows
//! - IsArray: Returns a Boolean indicating whether a variable is an array
//! - IsDate: Returns a Boolean indicating whether an expression can be converted to a date
//! - IsEmpty: Returns a Boolean indicating whether a Variant variable has been initialized
//! - IsError: Returns a Boolean indicating whether an expression is an error value
//! - IsMissing: Returns a Boolean indicating whether an optional Variant parameter was passed to a procedure
//! - IsNull: Returns a Boolean indicating whether an expression contains no valid data (Null)
//! - IsNumeric: Returns a Boolean indicating whether an expression can be evaluated as a number
//! - IsObject: Returns a Boolean indicating whether an identifier represents an object variable
//! - Join: Returns a string created by joining a number of substrings contained in an array
//! - LBound: Returns a Long containing the smallest available subscript for the indicated dimension of an array
//! - LCase: Returns a String that has been converted to lowercase
//! - Left: Returns a String containing a specified number of characters from the left side of a string
//! - Len: Returns a Long containing the number of characters in a string or the number of bytes required to store a variable
//! - LoadPicture: Returns a picture object containing an image from a file or memory
//! - LoadResData: Returns the data stored in a resource (.res) file
//! - LoadResPicture: Returns a picture object containing an image from a resource (.res) file
//! - LoadResString: Returns a string from a resource (.res) file
//! - Loc: Returns the current read/write position within an open file
//! - LOF: Returns the size, in bytes, of a file opened using the Open statement
//! - Log: Returns the natural logarithm of a number
//! - LTrim: Returns a string with leading spaces removed
//! - Mid: Returns a specified number of characters from a string
//! - Minute: Returns the minute of the hour (0-59)
//! - MIRR: Returns the modified internal rate of return for a series of periodic cash flows
//! - Month: Returns the month of the year (1-12)
//! - MonthName: Returns the name of the specified month
//! - MsgBox: Displays a message in a dialog box and returns which button was clicked
//! - Now: Returns the current system date and time
//! - NPer: Returns the number of periods for an annuity based on periodic fixed payments and a fixed interest rate
//! - NPV: Returns the net present value of an investment based on a series of periodic cash flows and a discount rate
//! - Oct: Returns a string representing the octal (base-8) value of a number
//! - RTrim: Returns a string with trailing spaces removed
//! - Trim: Returns a string with both leading and trailing spaces removed
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
mod fix;
mod format;
mod formatcurrency;
mod formatdatetime;
mod formatnumber;
mod formatpercent;
mod freefile;
mod fv;
mod getallsettings;
mod getattr;
mod getautoserversettings;
mod getobject;
mod getsetting;
mod hex;
mod hour;
mod iif;
mod imestatus;
mod input;
mod inputbox;
mod instr;
mod instrrev;
mod int;
mod ipmt;
mod irr;
mod isarray;
mod isdate;
mod isempty;
mod iserror;
mod ismissing;
mod isnull;
mod isnumeric;
mod isobject;
mod join;
mod lbound;
mod lcase;
mod left;
mod len;
mod loadpicture;
mod loadresdata;
mod loadrespicture;
mod loadresstring;
mod loc;
mod lof;
mod log;
mod ltrim;
mod mid;
mod minute;
mod mirr;
mod month;
mod monthname;
mod msgbox;
mod now;
mod nper;
mod npv;
mod oct;
mod rtrim;
mod trim;
