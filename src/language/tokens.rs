/// Represents a VB6 token.
///
/// This is a simple enum that represents the different types of tokens that can be parsed from VB6 code.
///
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub enum VB6Token<'a> {
    /// Represents whitespace.
    /// This is a collection of spaces, tabs, and other whitespace characters.
    Whitespace(&'a str),
    /// Represents a newline.
    /// This can be a carriage return, a newline, or a carriage return followed by a newline.
    Newline(&'a str),
    /// Represents a comment.
    /// Includes the single quote character.
    Comment(&'a str),
    /// Represents the 'ReDim' keyword.
    ///
    /// Used at a procedure level to reallocate storage space for a dynamic
    /// array.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    ReDimKeyword(&'a str),
    /// Represents the 'Preserve' keyword.
    ///
    /// Used with the 'ReDim' keyword to preserve the contents of an array when
    /// reallocating storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    PreserveKeyword(&'a str),
    /// Represents the 'Dim' keyword.
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    DimKeyword(&'a str),
    /// Represents the 'Declare' keyword.
    ///
    /// Used at the module level to declare references to external procedures
    /// in a dynamic-link library (DLL).
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    DeclareKeyword(&'a str),
    /// Represents the 'Alias' keyword.
    ///
    /// Used optionally in a Declare statement. Indicates that the procedure
    /// being called has another name in the DLL. This is useful when the
    /// external procedure name is the same as a keyword. You can also use Alias
    /// when a DLL procedure has the same name as a public variable, constant,
    /// or any other procedure in the same scope. Alias is also useful if any
    /// characters in the DLL procedure name aren't allowed by the DLL naming
    /// convention.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    AliasKeyword(&'a str),
    /// Represents the 'Lib' keyword.
    ///
    /// Indicates that a DLL or code resource contains the procedure being declared.
    /// The Lib clause is required for all declarations.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    LibKeyword(&'a str),
    /// Represents the 'With' keyword.
    ///
    /// Executes a series of statements on a single object or a user-defined type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266330(v=vs.60))
    WithKeyword(&'a str),
    /// Represents the 'WithEvents' keyword.
    ///
    /// Used with the 'Dim' keyword to declare a variable that can respond to
    /// events raised by an object.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    WithEventsKeyword(&'a str),
    /// Represents the 'Base' keyword.
    ///
    /// Used at module level to declare the default lower bound for array
    /// subscripts.
    ///
    ///[Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266179(v=vs.60))
    BaseKeyword(&'a str),
    /// Represents the 'Compare' keyword.
    ///
    /// Used at module level to declare the default comparison method to use
    /// when string data is compared.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266181(v=vs.60))
    CompareKeyword(&'a str),
    /// Represents the 'Option' keyword.
    ///
    /// Used at the module level in the Option Base, Option Compare, Option
    /// Explicit, or Option Private statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266185(v=vs.60))
    OptionKeyword(&'a str),
    /// Represents the 'Explicit' keyword.
    ///
    /// Used at the module level in the Option Explicit statement to force
    /// explicit declaration of all variables in that module.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266183(v=vs.60))
    ExplicitKeyword(&'a str),
    /// Represents the 'Private' keyword.
    ///
    /// Used at the module level to declare private variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266189(v=vs.60))
    PrivateKeyword(&'a str),
    /// Represents the 'Public' keyword.
    ///
    /// Used at the module level to declare public variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266207(v=vs.60))
    PublicKeyword(&'a str),
    /// Represents the 'Const' keyword.
    ///
    /// Declares constants for use in place of literal values.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243294(v=vs.60))
    ConstKeyword(&'a str),
    /// Represents the 'As' keyword.
    ///
    /// The 'As' keyword is used in these contexts:
    /// Const statement, Declare statement, Dim statement, Function statement,
    /// Name statement, Open statement, Open statement, private statement,
    /// Property Get statement, Property Let statement, Property Set statement,
    /// Public statement, ReDim statement, Static statement, Sub statement, and
    /// Type statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445149(v=vs.60))
    AsKeyword(&'a str),
    /// Represents the 'ByVal' keyword.
    ///
    /// Used in the following contexts:
    /// Call statement, Declare statement, Function statement, Property Get
    /// statement, Property Let statement, Property Set statement, and Sub
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445152(v=vs.60))
    ByValKeyword(&'a str),
    /// Represents the 'ByRef' keyword.
    ///
    /// Used in the following contexts:
    /// Call statement, Declare statement, Function statement, Property Get
    /// statement, Property Let statement, Property Set statement, and Sub
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445151(v=vs.60))
    ByRefKeyword(&'a str),
    /// Represents the 'Optional' keyword.
    ///
    /// Used in the following contexts:
    /// Declare statement, Function statement, Property Get statement,
    /// Property Let statement, Property Set statement, and Sub statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445195(v=vs.60))
    OptionalKeyword(&'a str),
    /// Represents the 'Function' keyword.
    ///
    /// Used to declare the name, argument, and code that forms the body of a
    /// function procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243374(v=vs.60))
    FunctionKeyword(&'a str),
    /// Represents the 'Static' keyword.
    ///
    /// Used at the procedure level to declare variable and allocate storage space.
    /// Variables declared the with Static statement retain their values as long
    /// as the module is loaded in memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266296(v=vs.60))
    StaticKeyword(&'a str),
    /// Represents the 'Sub' keyword.
    ///
    /// Used to declare the name, argument, and code that form the body of a sub
    /// procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266305(v=vs.60))
    SubKeyword(&'a str),
    /// Represents the 'End' keyword.
    ///
    /// Used to end a procedure or block.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243356(v=vs.60))
    EndKeyword(&'a str),
    /// Represents the 'True' keyword.
    ///
    /// The True keyword is used to represent the boolean value true and has a
    /// value equal to -1.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445231(v=vs.60))
    TrueKeyword(&'a str),
    /// Represents the 'False' keyword.
    ///
    /// The False keyword is used to represent the boolean value false and has a
    /// value equal to 0.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445157(v=vs.60))
    FalseKeyword(&'a str),
    /// Represents the 'Enum' keyword.
    ///
    /// Used to declare a type for an enumeration.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243358(v=vs.60))
    EnumKeyword(&'a str),
    /// Represents the 'Type' keyword.
    ///
    /// Used at the module level to declare a user-defined data type containing
    /// one or more elements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266315(v=vs.60))
    TypeKeyword(&'a str),
    /// Represents the 'Boolean' keyword.
    ///
    /// Used to declare a variable that can contain one of two values: True or
    /// False.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    BooleanKeyword(&'a str),
    /// Represents the 'Double' keyword.
    ///
    /// Used to declare a variable that can contain a double-precision floating-point
    /// number.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DoubleKeyword(&'a str),
    /// Represents the 'Currency' keyword.
    ///
    /// Used to declare a variable that can contain a currency value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    CurrencyKeyword(&'a str),
    /// Represents the 'Decimal' keyword.
    ///
    /// Used to declare a variable that can contain a decimal value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DecimalKeyword(&'a str),
    /// Represents the 'Date' keyword.
    ///
    /// Used to declare a variable that can contain a date value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DateKeyword(&'a str),
    /// Represents the 'Object' keyword.
    ///
    /// Used to declare a variable that can contain an object reference.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ObjectKeyword(&'a str),
    /// Represents the 'Variant' keyword.
    ///
    /// Used to declare a variable that can contain multiple kinds of types of
    /// data.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    VariantKeyword(&'a str),
    /// Represents the 'Byte' keyword.
    ///
    /// Used to declare a variable that can contain a byte value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ByteKeyword(&'a str),
    /// Represents the 'Long' keyword.
    ///
    /// Used to declare a variable that can contain a long integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    LongKeyword(&'a str),
    /// Represents the 'Single' keyword.
    ///
    /// Used to declare a variable that can contain a single-precision
    /// floating-point value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    SingleKeyword(&'a str),
    /// Represents the 'String' keyword.
    ///
    /// Used to declare a variable that can contain a string value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    StringKeyword(&'a str),
    /// Represents the 'Integer' keyword.
    ///
    /// Used to declare a variable that can contain an integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    IntegerKeyword(&'a str),
    /// Represents a string literal.
    ///
    /// The string literal includes the enclosing double quotes.
    StringLiteral(&'a str),
    /// Represents the 'If' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    IfKeyword(&'a str),
    /// Represents the 'Else' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseKeyword(&'a str),
    /// Represents the 'ElseIf' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseIfKeyword(&'a str),
    /// Represents the 'And' keyword.
    ///
    /// Used to perform a logical conjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242751(v=vs.60))
    AndKeyword(&'a str),
    /// Represents the 'Or' keyword.
    ///
    /// Used to perform a logical disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242850(v=vs.60))
    OrKeyword(&'a str),
    /// Represents the 'Xor' keyword.
    ///
    /// Used to perform a logical exclusive disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242859(v=vs.60))
    XorKeyword(&'a str),
    /// Represents the 'Mod' keyword.
    ///
    /// Used to perform a modulus operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242823(v=vs.60))
    ModKeyword(&'a str),
    /// Represents the 'Eqv' keyword.
    ///
    /// Used to perform a logical equivalence operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242780(v=vs.60))
    EqvKeyword(&'a str),
    /// Represents the 'AddressOf' keyword.
    ///
    /// A unary operator that obtains the address of the procedure it precedes
    /// and is used with API procedures that expect a function pointer at that
    /// position in the argument list.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242738(v=vs.60))
    AddressOfKeyword(&'a str),
    /// Represents the 'Imp' keyword.
    ///
    /// Used to perform a logical implication operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242794(v=vs.60))
    ImpKeyword(&'a str),
    /// Represents the 'Is' keyword.
    ///
    /// Used to perform a reference comparison between two object variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242809(v=vs.60))
    IsKeyword(&'a str),
    /// Represents the 'Like' keyword.
    ///
    /// Used to compare two strings.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242817(v=vs.60))
    LikeKeyword(&'a str),
    /// Represents the 'Not' keyword.
    ///
    /// Used to perform a logical negation on an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242842(v=vs.60))
    NotKeyword(&'a str),
    /// Represents the 'Then' keyword.
    ///
    /// Used to indicate the start of a block of code that is executed if the
    /// condition in an If statement is true.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445224(v=vs.60))
    ThenKeyword(&'a str),
    /// Represents the 'Goto' keyword.
    ///
    /// Branches unconditionally to a specific line within a procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243380(v=vs.60))
    GotoKeyword(&'a str),
    /// Represents the 'Exit' keyword.
    ///
    /// Exits a block of Do..Loop, For..Next, Function, Sub, or Property code.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243366(v=vs.60))
    ExitKeyword(&'a str),
    /// Represents the 'For' keyword.
    ///
    /// Used to declare a For..Next loop, or a For Each...Next loop.
    /// Repeats a group of statements a specified number of times.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243370(v=vs.60))
    ForKeyword(&'a str),
    /// Represents the 'To' keyword.
    ///
    /// The To keyword is used in these contexts:
    ///
    /// Dim statement, For...Next statement, Lock statement, Unlock statement,
    /// Private statement, Public statement, ReDim statement, Select Case statement,
    /// Static statement, and Type statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445229(v=vs.60))
    ToKeyword(&'a str),
    /// Represents the 'Lock' keyword.
    ///
    /// Controls access by other processes to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    LockKeyword(&'a str),
    /// Represents the 'Unlock' keyword.
    ///
    /// Controls access by other processes to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    UnlockKeyword(&'a str),
    /// Represents the 'Step' keyword.
    ///
    /// Used in the For...Next statement to specify the increment of the loop
    /// variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445219(v=vs.60))
    StepKeyword(&'a str),
    /// Represents the 'Stop' keyword.
    ///
    /// Used to suspend execution of a program.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266300(v=vs.60))
    StopKeyword(&'a str),
    /// Represents the 'While' keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WhileKeyword(&'a str),
    /// Represents the 'Wend' keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WendKeyword(&'a str),
    /// Represents the 'Width' keyword.
    ///
    /// Assigns an output line width to a file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266324(v=vs.60))
    WidthKeyword(&'a str),
    /// Represents the 'Write' keyword.
    ///
    /// Used to write data to a sequential file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266338(v=vs.60))
    WriteKeyword(&'a str),
    /// Represents the 'Time' keyword.
    ///
    /// Used to set the System time.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266310(v=vs.60))
    TimeKeyword(&'a str),
    /// Represents the 'SetAttr' keyword.
    ///
    /// Used to set attribute information for a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266286(v=vs.60))
    SetAttrKeyword(&'a str),
    /// Represents the 'Set' keyword.
    ///
    /// Used to assign an object reference to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266283(v=vs.60))
    SetKeyword(&'a str),
    /// Represents the 'SendKeys' keyword.
    ///
    /// Used to send one or more keystrokes to the active window as if typed at
    /// the keyboard.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266279(v=vs.60))
    SendKeysKeyword(&'a str),
    /// Represents the 'Select' keyword.
    ///
    /// Used to execute one of a several groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    SelectKeyword(&'a str),
    /// Represents the 'Case' keyword.
    ///
    /// Used to execute one of a several groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    CaseKeyword(&'a str),
    /// Represents the 'Seek' keyword.
    ///
    /// Used to set the position for the next read/write operation on a file
    /// opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266268(v=vs.60))
    SeekKeyword(&'a str),
    /// Represents the 'SaveSetting' keyword.
    ///
    /// Saves or creates an application entry in the application's entry in the Windows registry.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266261(v=vs.60))
    SaveSettingKeyword(&'a str),
    /// Represents the 'SavePicture' keyword.
    ///
    /// Saves a graphic from the `Picture` or `Image` property of an object or
    /// control (if one is associated with it) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445827(v=vs.60))
    SavePictureKeyword(&'a str),
    /// Represents the 'RSet' keyword.
    ///
    /// Right aligns a string within a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266256(v=vs.60))
    RSetKeyword(&'a str),
    /// Represents the 'RmDir' keyword.
    ///
    /// Removes an existing directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266252(v=vs.60))
    RmDirKeyword(&'a str),
    /// Represents the 'Resume' keyword.
    ///
    /// Resumes execution after an error-handling routine is finished.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266247(v=vs.60))
    ResumeKeyword(&'a str),
    /// Represents the 'Reset' keyword.
    ///
    /// Closes all disk files opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266242(v=vs.60))
    ResetKeyword(&'a str),
    /// Represents a 'REM' line comment.
    ///
    /// Includes the 'REM' characters and the comment text but not the newline.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266237(v=vs.60))
    RemComment(&'a str),
    /// Represents the 'Randomize' keyword.
    ///
    /// Initializes the random-number generator with a seed value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266225(v=vs.60))
    RandomizeKeyword(&'a str),
    /// Represents the 'RaiseEvent' keyword.
    ///
    /// Fires an event declared at module level within a class, form, or
    /// document.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266219(v=vs.60))
    RaiseEventKeyword(&'a str),
    /// Represents the 'Put' keyword.
    ///
    /// Writes data from a variable to a disk file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266212(v=vs.60))
    PutKeyword(&'a str),
    /// Represents the 'Property' keyword.
    ///
    /// Declares the name, argument, and code that forms the body of a property
    /// procedure, which sets a reference to a property of an object.
    ///
    /// Used in Property Get, Property Let, and Property Set statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266202(v=vs.60))
    PropertyKeyword(&'a str),
    /// Represents the 'Print' keyword.
    ///
    /// Writes display-formatted data to a sequential file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266187(v=vs.60))
    PrintKeyword(&'a str),
    /// Represents the 'Open' keyword.
    ///
    /// Enables input/output (I/O) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266177(v=vs.60))
    OpenKeyword(&'a str),
    /// Represents the 'On' keyword.
    ///
    /// Branch to one of several specified lines, dependin on the value of an expression.
    /// Used in the following contexts:
    ///
    /// On...GoSub statement, On...Goto statement, and On...Error statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266175(v=vs.60))
    OnKeyword(&'a str),
    /// Represents the 'Name' keyword.
    ///
    /// Renames a disk file, directory, or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266171(v=vs.60))
    NameKeyword(&'a str),
    /// Represents the 'MkDir' keyword.
    ///
    /// Creates a new directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266169(v=vs.60))
    MkDirKeyword(&'a str),
    /// Represents the 'Mid' keyword.
    ///
    /// Replaces a specified number of characters in a Variant (String) variable
    /// with characters from another string.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266166(v=vs.60))
    MidKeyword(&'a str),
    /// Represents the 'LSet' keyword.
    ///
    /// Left alligns a string within a string variable, or copies a variable of
    /// one user-defined type to another variable of a different user-defined type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266163(v=vs.60))
    LSetKeyword(&'a str),
    /// Represents the 'Load' keyword.
    ///
    /// Loads a form or control into memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445825(v=vs.60))
    LoadKeyword(&'a str),
    /// Represents the 'Line' keyword.
    ///
    /// Reads a single line from an open sequential file and assigns it to a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243392(v=vs.60))
    LineKeyword(&'a str),
    /// Represents the 'Input' keyword.
    ///
    /// Reads data from an open sequential file and assigns the data to variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243386(v=vs.60))
    InputKeyword(&'a str),
    //// Represents the 'Let' keyword.
    ///
    /// Assigns the value of an expression to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243390(v=vs.60))
    LetKeyword(&'a str),
    /// Represents the 'Kill' keyword.
    ///
    /// Deletes files from a disk.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243388(v=vs.60))
    KillKeyword(&'a str),
    /// Represents the 'Implements' keyword.
    ///
    /// Specifies an interface or class that will be implemented in the class module in which it appears.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243384(v=vs.60))
    ImplementsKeyword(&'a str),
    /// Represents the 'Get' keyword.
    ///
    /// Reads data from an open disk file into a variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243376(v=vs.60))
    GetKeyword(&'a str),
    /// Represents the 'FileCopy' keyword.
    ///
    /// Copies a file from one location to another.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243368(v=vs.60))
    FileCopyKeyword(&'a str),
    /// Represents the 'Event' keyword.
    ///
    /// Declares a user-defined event.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243364(v=vs.60))
    EventKeyword(&'a str),
    /// Represents the 'Error' keyword.
    ///
    /// Simulates the occurance of an error.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243362(v=vs.60))
    ErrorKeyword(&'a str),
    /// Represents the 'Erase' keyword.
    ///
    /// Reinitializes the elements of a fixed-size array and releases dynamic-array
    /// storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243360(v=vs.60))
    EraseKeyword(&'a str),
    /// Represents the 'Do' keyword.
    ///
    /// Repeats a block of statements while a condition is True or until a
    /// condition becomes True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243354(v=vs.60))
    DoKeyword(&'a str),
    /// Represents the 'Until' keyword.
    ///
    /// Used in the Do...Loop statement to specify the condition under which
    /// the loop terminates.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243354(v=vs.60))
    UntilKeyword(&'a str),
    /// Represents the 'DeleteSetting' keyword.
    ///
    /// Deletes a section or key setting from an application's entry in the
    /// Windows registry.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243347(v=vs.60))
    DeleteSettingKeyword(&'a str),
    /// Represents the 'DefBool' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Boolean data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefBoolKeyword(&'a str),
    /// Represents the 'DefByte' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Byte data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefByteKeyword(&'a str),
    /// Represents the 'DefInt' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Int data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefIntKeyword(&'a str),
    /// Represents the 'DefLng' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Long data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefLngKeyword(&'a str),
    /// Represents the 'DefCur' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Currency data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefCurKeyword(&'a str),
    /// Represents the 'DefSng' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Single data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefSngKeyword(&'a str),
    /// Represents the 'DefDbl' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Double data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDblKeyword(&'a str),
    /// Represents the 'DefDec' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Decimal data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDecKeyword(&'a str),
    /// Represents the 'DefDate' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Date data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDateKeyword(&'a str),
    /// Represents the 'DefStr' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the String data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefStrKeyword(&'a str),
    /// Represents the 'DefObj' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Object data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefObjKeyword(&'a str),
    /// Represents the 'DefVar' keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for Function and
    /// PropertyGet procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Variant data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefVarKeyword(&'a str),
    /// Represents the 'Close' keyword.
    ///
    /// Concludes input/output (I/O) to a file opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243283(v=vs.60))
    CloseKeyword(&'a str),
    /// Represents the 'ChDir' keyword.
    ///
    /// Changes the current drive.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243270(v=vs.60))
    ChDriveKeyword(&'a str),
    /// Represents the 'ChDir' keyword.
    ///
    /// Changes the current directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243256(v=vs.60))
    ChDirKeyword(&'a str),
    /// Represents the 'Call' keyword.
    ///
    /// Transfers control to a sub procedure, Function procedure, or dynamic-link
    /// library (DLL) procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243242(v=vs.60))
    CallKeyword(&'a str),
    /// Represents the 'Beep' keyword.
    ///
    /// Sounds a tone through the computer's speaker.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243233(v=vs.60))
    BeepKeyword(&'a str),
    /// Represents the 'AppActivate' keyword.
    ///
    /// Activates an application window.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243211(v=vs.60))
    AppActivateKeyword(&'a str),
    /// Represents the 'Friend' keyword.
    ///
    /// Modifies the definition of a procedure in a form module or class moduel
    /// to make the procedure callable from modules that are outside the class,
    /// but part of the project within which the class is defined. Friend
    /// procedures cannot be used in standard modules.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445159(v=vs.60))
    FriendKeyword(&'a str),
    /// Represents the 'Binary' keyword.
    ///
    /// The Binary keyword is used in these contexts:
    /// Open statement, Option Compare statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445150(v=vs.60))
    BinaryKeyword(&'a str),
    /// Represents the 'Empty' keyword.
    ///
    /// The Empty keyword is used as a Variant subtype. It indicates an
    /// uninitialized variable value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445155(v=vs.60))
    EmptyKeyword(&'a str),
    /// Represents the 'Next' keyword.
    ///
    /// The next keyword is used in these contexts:
    ///
    /// For...Next statement, For Each...Next statement, On Error statement, and
    /// a Resume statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445184(v=vs.60))
    NextKeyword(&'a str),
    /// Represents the 'New' keyword.
    ///
    /// Keyword that enables implicit creation of an object. If you use New when
    /// declaring the object variable, a new instance of the object is created
    /// on first reference to it, so you don't have to use the Set statement to
    /// assign the object reference. The New keyword can't be used to declare
    /// variables of any intrinsic data types, can't be used to declare
    /// instances of dependent objects, and can't be used with `WithEvents`.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    NewKeyword(&'a str),
    /// Represents the 'Len' keyword.
    ///
    /// The Len keyword is used in these contexts:
    ///
    /// Len Function, and the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445169(v=vs.60))
    LenKeyword(&'a str),
    /// Represents the 'Me' keyword.
    ///
    /// The 'Me' keyword behaves like an implicitly declared variable. It is
    /// automatically available to every procedure in a class module. When a
    /// class can have more than one instance, 'Me' provides a way to refer to
    /// the specific instance of the class where the code is executing. Using
    /// 'Me' is particularly useful for passing information about the currently
    /// executing instance of a class to a procedure in another module.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445177(v=vs.60))
    MeKeyword(&'a str),
    /// Represents the 'Null' keyword.
    ///
    /// The Null keyword is used as a Variant subtype. It indicates that a
    /// variable contains no valid data.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445190(v=vs.60))
    NullKeyword(&'a str),
    /// Represents the 'ParamArray' keyword.
    ///
    /// The 'ParamArray' keyword is used in these contexts:
    ///
    /// Declare statement, Function statement, Property Get statement,
    /// Property Let statement, and Sub statement.
    ///
    /// Used only as the last argument in arglist to indicate that the final
    /// argument is an Optional array of Variant elements. The 'ParamArray'
    /// keyword allows you to provide an arbitrary number of arguments. It may
    /// not be used with 'ByVal', 'ByRef', or Optional.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445198(v=vs.60))
    ParamArrayKeyword(&'a str),
    /// Represents a dollar sign '$'.
    ///
    /// Often used to indicate a variable is a string or that a function
    /// works with strings.
    DollarSign(&'a str),
    /// Represents an underscore '_'.
    ///
    /// Used to indicate that a statement continues on the next line.
    /// It must be preceded by at least one white space and must be the last
    /// character on the line.
    Underscore(&'a str),
    /// Represents an ampersand '&'.
    ///
    /// Used to force string concatenation of two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242763(v=vs.60))
    Ampersand(&'a str),
    /// Represents a percent sign '%'.
    Percent(&'a str),
    /// Represents an octothorpe '#'.
    Octothorpe(&'a str),
    /// Represents a left paranthesis '('.
    LeftParentheses(&'a str),
    /// Represents a right paranthesis ')'.
    RightParentheses(&'a str),
    /// Represents a left square bracket '['.
    LeftSquareBracket(&'a str),
    /// Represents a right square bracket ']'.
    RightSquareBracket(&'a str),
    /// Represents a comma ','.
    Comma(&'a str),
    /// Represents a semicolon ';'.
    Semicolon(&'a str),
    /// Represents the 'at' symbol '@'.
    AtSign(&'a str),
    /// Represents an exclamation mark '!'.
    ExclamationMark(&'a str),
    /// Represents an equality operator '=' can also be the assignment operator.
    ///
    /// Used to assign a value to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242760(v=vs.60))
    EqualityOperator(&'a str),
    /// Represents a less than operator '<'.
    LessThanOperator(&'a str),
    /// Represents a greater than operator '>'.
    GreaterThanOperator(&'a str),
    /// Represents a multiplication operator '*'.
    MultiplicationOperator(&'a str),
    /// Represents a subtraction operator '-'.
    SubtractionOperator(&'a str),
    /// Represents an addition operator '+'.
    AdditionOperator(&'a str),
    /// Represents a division operator '/'.
    DivisionOperator(&'a str),
    /// Represents a backward slash operator '\\'.
    BackwardSlashOperator(&'a str),
    /// Represents a period operator '.'.
    PeriodOperator(&'a str),
    /// Represents a colon operator ':'.
    ColonOperator(&'a str),
    /// Represents an exponentiation operator '^'.
    ExponentiationOperator(&'a str),
    /// Represents an Identifier, variable, or function name.
    /// This is a name that starts with a letter and can contain letters, numbers, and underscores.
    Identifier(&'a str),
    /// Represents a number.
    /// This is just a collection of digits and hasn't been parsed into a
    /// specific kind of number yet.
    Number(&'a str),
}
