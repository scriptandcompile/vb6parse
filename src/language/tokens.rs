use bstr::BStr;

/// Represents a VB6 token.
///
/// This is a simple enum that represents the different types of tokens that can be parsed from VB6 code.
///
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub enum VB6Token<'a> {
    /// Represents whitespace.
    /// This is a collection of spaces, tabs, and other whitespace characters.
    Whitespace(&'a BStr),
    /// Represents a newline.
    /// This can be a carriage return, a newline, or a carriage return followed by a newline.
    Newline(&'a BStr),
    /// Represents a comment.
    /// Includes the single quote character.
    Comment(&'a BStr),
    /// Represents the ReDim keyword.
    ///
    /// Used at a procedure level to reallocate storage space for a dynamic
    /// array.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    ReDimKeyword(&'a BStr),
    /// Represents the Preserve keyword.
    ///
    /// Used with the ReDim keyword to preserve the contents of an array when
    /// reallocating storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    PreserveKeyword(&'a BStr),
    /// Represents the Dim keyword.
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    DimKeyword(&'a BStr),
    /// Represents the Declare keyword.
    ///
    /// Used at the module level to declare references to external procedures
    /// in a dynamic-link library (DLL).
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    DeclareKeyword(&'a BStr),
    LibKeyword(&'a BStr),
    WithKeyword(&'a BStr),
    /// Represents the 'WithEvents' keyword.
    ///
    /// Used with the 'Dim' keyword to declare a variable that can respond to
    /// events raised by an object.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    WithEventsKeyword(&'a BStr),

    OptionKeyword(&'a BStr),
    ExplicitKeyword(&'a BStr),

    PrivateKeyword(&'a BStr),
    PublicKeyword(&'a BStr),

    ConstKeyword(&'a BStr),
    AsKeyword(&'a BStr),
    ByValKeyword(&'a BStr),
    ByRefKeyword(&'a BStr),
    OptionalKeyword(&'a BStr),

    FunctionKeyword(&'a BStr),
    SubKeyword(&'a BStr),
    EndKeyword(&'a BStr),

    /// Represents the boolean literal `True`.
    TrueKeyword(&'a BStr),
    /// Represents the boolean literal `False`.
    FalseKeyword(&'a BStr),

    EnumKeyword(&'a BStr),
    TypeKeyword(&'a BStr),

    BooleanKeyword(&'a BStr),
    ByteKeyword(&'a BStr),
    LongKeyword(&'a BStr),
    SingleKeyword(&'a BStr),
    StringKeyword(&'a BStr),
    IntegerKeyword(&'a BStr),

    /// Represents a string literal.
    /// The string literal is enclosed in double quotes.
    StringLiteral(&'a BStr),

    IfKeyword(&'a BStr),
    ElseKeyword(&'a BStr),
    AndKeyword(&'a BStr),
    OrKeyword(&'a BStr),
    NotKeyword(&'a BStr),
    ThenKeyword(&'a BStr),

    GotoKeyword(&'a BStr),
    ExitKeyword(&'a BStr),

    ForKeyword(&'a BStr),
    ToKeyword(&'a BStr),
    StepKeyword(&'a BStr),
    NextKeyword(&'a BStr),

    /// Represents a dollar sign '$'.
    DollarSign(&'a BStr),
    /// Represents an underscore '_'.
    Underscore(&'a BStr),
    /// Represents an ampersand '&'.
    Ampersand(&'a BStr),
    /// Represents a percent sign '%'.
    Percent(&'a BStr),
    /// Represents an octothorpe '#'.
    Octothorpe(&'a BStr),

    /// Represents a left paranthesis '('.
    LeftParanthesis(&'a BStr),
    /// Represents a right paranthesis ')'.
    RightParanthesis(&'a BStr),

    /// Represents a left square bracket '['.
    LeftSquareBracket(&'a BStr),
    /// Represents a right square bracket ']'.
    RightSquareBracket(&'a BStr),

    /// Represents a comma ','.
    Comma(&'a BStr),
    /// Represents a semicolon ';'.
    Semicolon(&'a BStr),

    /// Represents the 'at' symbol '@'.
    AtSign(&'a BStr),

    /// Represents an exclamation mark '!'.
    ExclamationMark(&'a BStr),

    /// Represents an equality operator '=' can also be the assignment operator.
    EqualityOperator(&'a BStr),
    /// Represents a less than operator '<'.
    LessThanOperator(&'a BStr),
    /// Represents a greater than operator '>'.
    GreaterThanOperator(&'a BStr),
    /// Represents a multiplication operator '*'.
    MultiplicationOperator(&'a BStr),
    /// Represents a subtraction operator '-'.
    SubtractionOperator(&'a BStr),
    /// Represents an addition operator '+'.
    AdditionOperator(&'a BStr),
    /// Represents a division operator '/'.
    DivisionOperator(&'a BStr),
    /// Represents a backward slash operator '\\'.
    BackwardSlashOperator(&'a BStr),
    /// Represents a period operator '.'.
    PeriodOperator(&'a BStr),
    /// Represents a colon operator ':'.
    ColonOperator(&'a BStr),
    /// Represents an exponentiation operator '^'.
    ExponentiationOperator(&'a BStr),

    /// Represents a variable name.
    /// This is a name that starts with a letter and can contain letters, numbers, and underscores.
    VariableName(&'a BStr),
    /// Represents a number.
    /// This is just a collection of digits and hasn't been parsed into a
    /// specific kind of number yet.
    Number(&'a BStr),
}
