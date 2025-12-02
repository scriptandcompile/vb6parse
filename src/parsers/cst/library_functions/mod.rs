//! Built-in VB6 library functions.
//!
//! This module provides documentation and parsing support for VB6's built-in
//! library functions. While these functions are parsed as regular function calls
//! (`CallExpression` nodes), this module serves to document their behavior and
//! provide unit tests to ensure parsing support.
//!
//! The library functions handled here include:
//! - `Abs`: Returns the absolute value of a number
//! - `Array`: Returns a Variant containing an array
//! - `Asc`: Returns the character code of the first letter in a string
//! - `AscB`: Returns the byte value of the first byte in a string
//! - `AscW`: Returns the Unicode character code of the first character in a string
//! - `Atn`: Returns the arctangent of a number in radians
//! - `CallByName`: Executes a method or accesses a property by name at runtime
//! - `Choose`: Returns a value from a list of choices based on an index
//! - `Chr`: Returns the character associated with the specified character code
//! - `Chr$`: Returns a String containing the character associated with the specified character code
//! - `ChrB`: Returns a String containing the character associated with the specified ANSI character code
//! - `ChrB$`: Returns a String containing the character associated with the specified ANSI character code
//! - `ChrW`: Returns a String containing the Unicode character associated with the specified character code
//! - `Command`: Returns the command-line arguments passed to the program
//! - `Cos`: Returns the cosine of an angle in radians
//! - `CreateObject`: Creates and returns a reference to an `ActiveX` object
//! - `CurDir`: Returns the current path for the specified drive
//! - `CVErr`: Returns a Variant of subtype Error containing an error number
//! - `Date`: Returns the current system date
//! - `DateAdd`: Returns a date to which a specified time interval has been added
//! - `DateDiff`: Returns the number of time intervals between two dates
//! - `DatePart`: Returns a specified part of a given date
//! - `DateSerial`: Returns a date for a specified year, month, and day
//! - `DateValue`: Returns a date from a string expression
//! - `Day`: Returns the day of the month (1-31)
//! - `DDB`: Returns depreciation using the double-declining balance method
//! - `Dir`: Returns a file or directory name matching a pattern
//! - `DoEvents`: Yields execution to process other events
//! - `Environ`: Returns the value of an operating system environment variable
//! - `EOF`: Returns a Boolean indicating whether the end of a file has been reached
//! - `Error`: Returns the error message corresponding to a given error number
//! - `Error$`: Returns the error message string corresponding to a given error number
//! - `Exp`: Returns e (the base of natural logarithms) raised to a power
//! - `FileAttr`: Returns the file mode or file handle for an open file
//! - `FileDateTime`: Returns the date and time when a file was created or last modified
//! - `FileLen`: Returns the length of a file in bytes
//! - `Filter`: Returns a zero-based array containing a subset of a string array based on filter criteria
//! - `Fix`: Returns the integer portion of a number
//! - `Format`: Returns a formatted string expression according to format instructions
//! - `FormatCurrency`: Returns an expression formatted as a currency value using the system currency symbol
//! - `FormatDateTime`: Returns an expression formatted as a date or time
//! - `FormatNumber`: Returns an expression formatted as a number
//! - `FormatPercent`: Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character
//! - `FreeFile`: Returns the next file number available for use by the Open statement
//! - `Fv`: Returns the future value of an annuity based on periodic, fixed payments and a fixed interest rate
//! - `GetAllSettings`: Returns a list of key settings and their values from the Windows registry
//! - `GetAttr`: Returns an `Integer` representing the attributes of a file, directory, or folder
//! - `GetAutoServerSettings`: Returns information about the security settings for a `DCOM` server
//! - `GetObject`: Returns a reference to an `ActiveX` object from a file or a running instance
//! - `GetSetting`: Returns a registry key setting value from the Windows registry
//! - `Hex`: Returns a `String` representing the hexadecimal value of a number
//! - `Hour`: Returns an `Integer` specifying a whole number between 0 and 23, inclusive, representing the hour of the day
//! - `IIf`: Returns one of two parts, depending on the evaluation of an expression
//! - `IMEStatus`: Returns an `Integer` indicating the current `Input Method Editor` (`IME`) mode of Microsoft Windows
//! - `Input`: Returns `String` containing characters from a file opened in Input or Binary mode
//! - `InputBox`: Displays a prompt in a dialog box, waits for user input, and returns a `String`
//! - `InStr`: Returns a `Long` specifying the position of the first occurrence of one string within another
//! - `InStrRev`: Returns the position of an occurrence of one string within another, from the end of string
//! - `Int`: Returns the integer portion of a number
//! - `IPmt`: Returns the interest payment for a given period of an annuity
//! - `IRR`: Returns the internal rate of return for a series of periodic cash flows
//! - `IsArray`: Returns a `Boolean` indicating whether a variable is an array
//! - `IsDate`: Returns a `Boolean` indicating whether an expression can be converted to a date
//! - `IsEmpty`: Returns a `Boolean` indicating whether a Variant variable has been initialized
//! - `IsError`: Returns a `Boolean` indicating whether an expression is an error value
//! - `IsMissing`: Returns a `Boolean` indicating whether an optional Variant parameter was passed to a procedure
//! - `IsNull`: Returns a `Boolean` indicating whether an expression contains no valid data (Null)
//! - `IsNumeric`: Returns a `Boolean` indicating whether an expression can be evaluated as a number
//! - `IsObject`: Returns a `Boolean` indicating whether an identifier represents an object variable
//! - `Join`: Returns a `String` created by joining a number of substrings contained in an array
//! - `LBound`: Returns a `Long` containing the smallest available subscript for the indicated dimension of an array
//! - `LCase`: Returns a `String` that has been converted to lowercase
//! - `LCase$`: Returns a `String` that has been converted to lowercase (explicit String type)
//! - `Left`: Returns a `String` containing a specified number of characters from the left side of a string
//! - `Len`: Returns a `Long` containing the number of characters in a string or the number of bytes required to store a variable
//! - `LenB`: Returns a `Long` containing the number of bytes used to represent a string in memory
//! - `LoadPicture`: Returns a picture object containing an image from a file or memory
//! - `LoadResData`: Returns the data stored in a resource (.res) file
//! - `LoadResPicture`: Returns a picture object containing an image from a resource (.res) file
//! - `LoadResString`: Returns a string from a resource (.res) file
//! - `Loc`: Returns the current read/write position within an open file
//! - `LOF`: Returns the size, in bytes, of a file opened using the Open statement
//! - `Log`: Returns the natural logarithm of a number
//! - `LTrim`: Returns a string with leading spaces removed
//! - `Mid`: Returns a specified number of characters from a string
//! - `Mid$`: Returns a specified number of characters from a string (always returns `String`)
//! - `MidB`: Returns a specified number of bytes from a string
//! - `Minute`: Returns the minute of the hour (0-59)
//! - `MIRR`: Returns the modified internal rate of return for a series of periodic cash flows
//! - `Month`: Returns the month of the year (1-12)
//! - `MonthName`: Returns the name of the specified month
//! - `MsgBox`: Displays a message in a dialog box and returns which button was clicked
//! - `Now`: Returns the current system date and time
//! - `NPer`: Returns the number of periods for an annuity based on periodic fixed payments and a fixed interest rate
//! - `NPV`: Returns the net present value of an investment based on a series of periodic cash flows and a discount rate
//! - `Oct`: Returns a `String` representing the octal (base-8) value of a number
//! - `Partition`: Returns a `String` indicating where a number occurs within a calculated series of ranges
//! - `Pmt`: Returns the payment for an annuity based on periodic, fixed payments and a fixed interest rate
//! - `PPmt`: Returns the principal payment for a given period of an annuity
//! - `PV`: Returns the present value of an annuity based on periodic, fixed payments and a fixed interest rate
//! - `QBColor`: Returns the RGB color code corresponding to the specified `QBasic` color number
//! - `Rate`: Returns the interest rate per period for an annuity
//! - `Replace`: Returns a `String` in which a specified substring has been replaced with another substring
//! - `RGB`: Returns a `Long` representing an RGB color value from red, green, and blue color components
//! - `Right`: Returns a `String` containing a specified number of characters from the right side of a string
//! - `Rnd`: Returns a `Single` containing a pseudo-random number
//! - `Round`: Returns a number rounded to a specified number of decimal places
//! - `RTrim`: Returns a `String` with trailing spaces removed
//! - `Second`: Returns an `Integer` specifying the second of the minute (0-59)
//! - `Seek`: Returns a `Long` specifying the current read/write position within a file
//! - `Sgn`: Returns an `Integer` indicating the sign of a number (-1, 0, or 1)
//! - `Shell`: Runs an executable program and returns a task ID
//! - `Sin`: Returns the sine of an angle in radians
//! - `SLN`: Returns straight-line depreciation of an asset for a single period
//! - `Space`: Returns a `String` consisting of the specified number of spaces
//! - `Spc`: Positions output by inserting spaces in Print statements
//! - `Split`: Returns a zero-based array containing a specified number of substrings
//! - `Sqr`: Returns the square root of a number
//! - `Str`: Converts a number to a string representation
//! - `StrComp`: Compares two `String`s and returns a value indicating their relationship
//! - `StrConv`: Converts a `String` to a specified format
//! - `String`: Returns a `String` consisting of a repeating character
//! - `StrReverse`: Returns a `String` in which the character order is reversed
//! - `Switch`: Evaluates a list of expressions and returns a value associated with the first expression that is True
//! - `SYD`: Returns the sum-of-years digits depreciation of an asset for a specified period
//! - `Tab`: Positions output at a specific column in Print statements
//! - `Tan`: Returns the tangent of an angle in radians
//! - `Time`: Returns the current system time
//! - `Timer`: Returns the number of seconds elapsed since midnight
//! - `TimeSerial`: Returns a time for a specific hour, minute, and second
//! - `TimeValue`: Returns a time value from a string expression
//! - `Trim`: Returns a `String` with both leading and trailing spaces removed
//! - `TypeName`: Returns a `String` describing the data type of a variable or expression
//! - `UBound`: Returns the largest available subscript for the indicated dimension of an array
//! - `UCase`: Returns a `String` that has been converted to uppercase
//! - `UCase$`: Returns a `String` that has been converted to uppercase (explicit String type)
//! - `VarType`: Returns an `Integer` constant indicating the Variant subtype of a variable or expression
//! - `Weekday`: Returns an `Integer` representing the day of the week
//! - `WeekdayName`: Returns a `String` indicating the specified day of the week
//! - `Year`: Returns an `Integer` representing the year
//! - `TypeName`: Returns a `String` describing the data type of a variable or expression
//! - `UBound`: Returns the largest available subscript for the indicated dimension of an array
//! - `UCase`: Returns a `String` that has been converted to uppercase
//! - `Weekday`: Returns an `Integer` representing the day of the week
//! - `WeekdayName`: Returns a `String` indicating the specified day of the week
//! - `Year`: Returns an `Integer` representing the year
//!
//!
//! Note: Unlike library statements (which are keywords), library functions are
//! called like regular functions and are parsed as `CallExpression` nodes in the CST.
//! This module primarily serves as documentation and reference for VB6's
//! built-in function library.

mod abs;
mod array;
mod asc;
mod ascb;
mod ascw;
mod atn;
mod callbyname;
mod choose;
mod chr;
mod chr_dollar;
mod chrb;
mod chrb_dollar;
mod chrw;
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
mod error_dollar;
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
mod lcase_dollar;
mod left;
mod len;
mod lenb;
mod loadpicture;
mod loadresdata;
mod loadrespicture;
mod loadresstring;
mod loc;
mod lof;
mod log;
mod ltrim;
mod mid;
mod mid_dollar;
mod midb;
mod minute;
mod mirr;
mod month;
mod monthname;
mod msgbox;
mod now;
mod nper;
mod npv;
mod oct;
mod partition;
mod pmt;
mod ppmt;
mod pv;
mod qbcolor;
mod rate;
mod replace;
mod rgb;
mod right;
mod rnd;
mod round;
mod rtrim;
mod second;
mod seek;
mod sgn;
mod shell;
mod sin;
mod sln;
mod space;
mod spc;
mod split;
mod sqr;
mod str;
mod strcomp;
mod strconv;
mod string;
mod strreverse;
mod switch;
mod syd;
mod tab;
mod tan;
mod time;
mod timer;
mod timeserial;
mod timevalue;
mod trim;
mod typename;
mod ubound;
mod ucase;
mod ucase_dollar;
mod vartype;
mod weekday;
mod weekdayname;
mod year;
