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
    /// Represents the 'ReDim' keyword.
    ///
    /// Used at a procedure level to reallocate storage space for a dynamic
    /// array.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    ReDimKeyword(&'a BStr),
    /// Represents the 'Preserve' keyword.
    ///
    /// Used with the ReDim keyword to preserve the contents of an array when
    /// reallocating storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    PreserveKeyword(&'a BStr),
    /// Represents the 'Dim' keyword.
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    DimKeyword(&'a BStr),
    /// Represents the 'Declare' keyword.
    ///
    /// Used at the module level to declare references to external procedures
    /// in a dynamic-link library (DLL).
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    DeclareKeyword(&'a BStr),
    /// Represents the 'Alias' keyword.
    ///
    /// Used optionally in a Declate statement. Indicates that the procedure
    /// being called has another name in the DLL. This is useful when the
    /// external procedure name is the same as a keyword. You can also use Alias
    /// when a DLL procedure has the same name as a public variable, constant,
    /// or any other procedure in the same scope. Alias is also useful if any
    /// characters in the DLL procedure name aren't allowed by the DLL naming
    /// convention.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    AliasKeyword(&'a BStr),
    /// Represents the 'Lib' keyword.
    ///
    /// Indicates that a DLL or code resource contains the procedure being declared.
    /// The Lib clause is required for all declarations.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    LibKeyword(&'a BStr),
    /// Represents the 'With' keyword.
    ///
    /// Executes a series of statements on a single object or a user-defined type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266330(v=vs.60))
    WithKeyword(&'a BStr),
    /// Represents the 'WithEvents' keyword.
    ///
    /// Used with the 'Dim' keyword to declare a variable that can respond to
    /// events raised by an object.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    WithEventsKeyword(&'a BStr),
    /// Represents the 'Base' keyword.
    ///
    /// Used at module level to declare the default lower bound for array
    /// subscripts.
    ///
    ///[Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266179(v=vs.60))
    BaseKeyword(&'a BStr),
    /// Represents the 'Compare' keyword.
    ///
    /// Used at module level to declare the default comparison method to use
    /// when string data is compared.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266181(v=vs.60))
    CompareKeyword(&'a BStr),
    /// Represents the 'Option' keyword.
    ///
    /// Used at the module level in the Option Base, Option Compare, Option
    /// Explicit, or Option Private statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266185(v=vs.60))
    OptionKeyword(&'a BStr),
    /// Represents the 'Explicit' keyword.
    ///
    /// Used at the module level in the Option Explicit statement to force
    /// explicit declaration of all variables in that module.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266183(v=vs.60))
    ExplicitKeyword(&'a BStr),
    /// Represents the 'Private' keyword.
    ///
    /// Used at the module level to declare private vairables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266189(v=vs.60))
    PrivateKeyword(&'a BStr),
    /// Represents the 'Public' keyword.
    ///
    /// Used at the module level to declare public variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266207(v=vs.60))
    PublicKeyword(&'a BStr),
    /// Represents the 'Const' keyword.
    ///
    /// Declares constants for use in place of literal values.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243294(v=vs.60))
    ConstKeyword(&'a BStr),
    /// Represents the 'As' keyword.
    ///
    /// The 'As' keyword is used in these contexts:
    /// Const statement, Declare statemenet, Dim statement, Function statenement,
    /// Name statement, Open statement, Open statement, private statement,
    /// Property Get statement, Property Let statement, Property Set statement,
    /// Public statement, ReDim statement, Static statement, Sub statement, and
    /// Type statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445149(v=vs.60))
    AsKeyword(&'a BStr),
    /// Represents the 'ByVal' keyword.
    ///
    /// Used in the following contexts:
    /// Call statement, Declare statement, Function statement, Property Get
    /// statement, Property Let statement, Property Set statement, and Sub
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445152(v=vs.60))
    ByValKeyword(&'a BStr),
    /// Represents the 'ByRef' keyword.
    ///
    /// Used in the following contexts:
    /// Call statement, Declare statement, Function statement, Property Get
    /// statement, Property Let statement, Property Set statement, and Sub
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445151(v=vs.60))
    ByRefKeyword(&'a BStr),
    /// Represents the 'Optional' keyword.
    ///
    /// Used in the following contexts:
    /// Declare statement, Function statement, Property Get statement,
    /// Property Let statement, Property Set statement, and Sub statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445195(v=vs.60))
    OptionalKeyword(&'a BStr),
    /// Represents the 'Function' keyword.
    ///
    /// Used to declare the name, argument, and code that forms the body of a
    /// function procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243374(v=vs.60))
    FunctionKeyword(&'a BStr),
    /// Represents the 'Static' keyword.
    ///
    /// Used at the procedure level to declare variable and allocate storage space.
    /// Variables declared the with Static statement retain their values as long
    /// as the module is loaded in memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266296(v=vs.60))
    StaticKeyword(&'a BStr),
    /// Represents the 'Sub' keyword.
    ///
    /// Used to declare the name, argument, and code that form the body of a sub
    /// procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266305(v=vs.60))
    SubKeyword(&'a BStr),
    /// Represents the 'End' keyword.
    ///
    /// Used to end a procedure or block.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243356(v=vs.60))
    EndKeyword(&'a BStr),
    /// Represents the 'True' keyword.
    ///
    /// The True keyword is used to represent the boolean value true and has a
    /// value equal to -1.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445231(v=vs.60))
    TrueKeyword(&'a BStr),
    /// Represents the 'False' keyword.
    ///
    /// The False keyword is used to represent the boolean value false and has a
    /// value equal to 0.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445157(v=vs.60))
    FalseKeyword(&'a BStr),
    /// Represents the 'Enum' keyword.
    ///
    /// Used to declare a type for an enumeration.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243358(v=vs.60))
    EnumKeyword(&'a BStr),
    /// Represents the 'Type' keyword.
    ///
    /// Used at the module level to declare a user-defined data type containing
    /// one or more elements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266315(v=vs.60))
    TypeKeyword(&'a BStr),
    /// Represents the 'Boolean' keyword.
    ///
    /// Used to declare a variable that can contain one of two values: True or
    /// False.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    BooleanKeyword(&'a BStr),
    /// Represents the 'Double' keyword.
    ///
    /// Used to declare a variable that can contain a double-precision floating-point
    /// number.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DoubleKeyword(&'a BStr),
    /// Represents the 'Currency' keyword.
    ///
    /// Used to declare a variable that can contain a currency value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    CurrencyKeyword(&'a BStr),
    /// Represents the 'Decimal' keyword.
    ///
    /// Used to declare a variable that can contain a decimal value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DecimalKeyword(&'a BStr),
    /// Represents the 'Date' keyword.
    ///
    /// Used to declare a variable that can contain a date value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DateKeyword(&'a BStr),
    /// Represents the 'Object' keyword.
    ///
    /// Used to declare a variable that can contain an object reference.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ObjectKeyword(&'a BStr),
    /// Represents the 'Variant' keyword.
    ///
    /// Used to declare a variable that can contain multiple kinds of types of
    /// data.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    VariantKeyword(&'a BStr),
    /// Represents the 'Byte' keyword.
    ///
    /// Used to declare a variable that can contain a byte value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ByteKeyword(&'a BStr),
    /// Represents the 'Long' keyword.
    ///
    /// Used to declare a variable that can contain a long integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    LongKeyword(&'a BStr),
    /// Represents the 'Single' keyword.
    ///
    /// Used to declare a variable that can contain a single-precision
    /// floating-point value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    SingleKeyword(&'a BStr),
    /// Represents the 'String' keyword.
    ///
    /// Used to declare a variable that can contain a string value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    StringKeyword(&'a BStr),
    /// Represents the 'Integer' keyword.
    ///
    /// Used to declare a variable that can contain an integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    IntegerKeyword(&'a BStr),
    /// Represents a string literal.
    ///
    /// The string literal is enclosed in double quotes.
    StringLiteral(&'a BStr),
    /// Represents the 'If' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    IfKeyword(&'a BStr),
    /// Represents the 'Else' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseKeyword(&'a BStr),
    /// Represents the 'ElseIf' keyword.
    ///
    /// Used to to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseIfKeyword(&'a BStr),
    /// Represents the 'And' keyword.
    ///
    /// Used to perform a logical conjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242751(v=vs.60))
    AndKeyword(&'a BStr),
    /// Represents the 'Or' keyword.
    ///
    /// Used to perform a logical disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242850(v=vs.60))
    OrKeyword(&'a BStr),
    /// Represents the 'Xor' keyword.
    ///
    /// Used to perform a logical exclusive disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242859(v=vs.60))
    XorKeyword(&'a BStr),
    /// Represents the 'Mod' keyword.
    ///
    /// Used to perform a modulus operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242823(v=vs.60))
    ModKeyword(&'a BStr),
    /// Represents the 'Eqv' keyword.
    ///
    /// Used to perform a logical equivalence operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242780(v=vs.60))
    EqvKeyword(&'a BStr),
    /// Represents the 'AddressOf' keyword.
    ///
    /// A unary operator that obtains the address of the procedure it precedes
    /// and is used with API procedures that expect a function pointer at that
    /// position in the argument list.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242738(v=vs.60))
    AddressOfKeyword(&'a BStr),
    /// Represents the 'Imp' keyword.
    ///
    /// Used to perform a logical implication operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242794(v=vs.60))
    ImpKeyword(&'a BStr),
    /// Represents the 'Is' keyword.
    ///
    /// Used to perform a reference comparison between two object variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242809(v=vs.60))
    IsKeyword(&'a BStr),
    /// Represents the 'Like' keyword.
    ///
    /// Used to compare two strings.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242817(v=vs.60))
    LikeKeyword(&'a BStr),
    /// Represents the 'Not' keyword.
    ///
    /// Used to perform a logical negation on an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242842(v=vs.60))
    NotKeyword(&'a BStr),
    /// Represents the 'Then' keyword.
    ///
    /// Used to indicate the start of a block of code that is executed if the
    /// condition in an If statement is true.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445224(v=vs.60))
    ThenKeyword(&'a BStr),
    /// Represents the 'Goto' keyword.
    ///
    /// Branches unconditionally to a specific line within a procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243380(v=vs.60))
    GotoKeyword(&'a BStr),
    /// Represents the 'Exit' keyword.
    ///
    /// Exits a block of Do..Loop, For..Next, Function, Sub, or Property code.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243366(v=vs.60))
    ExitKeyword(&'a BStr),
    /// Represents the 'For' keyword.
    ///
    /// Used to declare a For..Next loop, or a For Each...Next loop.
    /// Repeates a group of statements a specified number of times.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243370(v=vs.60))
    ForKeyword(&'a BStr),
    /// Represents the 'To' keyword.
    ///
    /// The To keyword is used in these contexts:
    ///
    /// Dim statement, For...Next statement, Lock statement, Unlock statement,
    /// Private statement, Public statement, ReDim statement, Select Case statement,
    /// Static statement, and Type statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445229(v=vs.60))
    ToKeyword(&'a BStr),
    /// Represents the 'Lock' keyword.
    ///
    /// Controls access by other processess to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    LockKeyword(&'a BStr),
    /// Represents the 'Unlock' keyword.
    ///
    /// Controls access by other processess to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    UnlockKeyword(&'a BStr),
    /// Represents the 'Step' keyword.
    ///
    /// Used in the For...Next statement to specify the increment of the loop
    /// variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445219(v=vs.60))
    StepKeyword(&'a BStr),
    /// Represents the 'Stop' keyword.
    ///
    /// Used to suspend execution of a program.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266300(v=vs.60))
    StopKeyword(&'a BStr),
    /// Represents the 'While' keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WhileKeyword(&'a BStr),
    /// Represents the 'Wend' keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WendKeyword(&'a BStr),
    /// Represents the 'Width' keyword.
    ///
    /// Assigns an output line width to a file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266324(v=vs.60))
    WidthKeyword(&'a BStr),
    /// Represents the 'Write' keyword.
    ///
    /// Used to write data to a sequential file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266338(v=vs.60))
    WriteKeyword(&'a BStr),
    /// Represents the 'Time' keyword.
    ///
    /// Used to set the System time.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266310(v=vs.60))
    TimeKeyword(&'a BStr),
    /// Represents the 'SetAttr' keyword.
    ///
    /// Used to set attribute information for a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266286(v=vs.60))
    SetAttrKeyword(&'a BStr),
    /// Represents the 'Set' keyword.
    ///
    /// Used to assign an object reference to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266283(v=vs.60))
    SetKeyword(&'a BStr),
    /// Represents the 'SendKeys' keyword.
    ///
    /// Used to send one or more keystrokes to the active window as if typed at
    /// the keyboard.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266279(v=vs.60))
    SendKeysKeyword(&'a BStr),
    /// Represents the 'Select' keyword.
    ///
    /// Used to execute one of a seveeral groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    SelectKeyword(&'a BStr),
    /// Represents the 'Case' keyword.
    ///
    /// Used to execute one of a seveeral groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    CaseKeyword(&'a BStr),
    /// Represents the 'Seek' keyword.
    ///
    /// Used to set the position for the next read/write operation on a file
    /// opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266268(v=vs.60))
    SeekKeyword(&'a BStr),
    /// Represents the 'SaveSetting' keyword.
    ///
    /// Saves or creates an application entry in the application's entry in the Windows registry.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266261(v=vs.60))
    SaveSettingKeyword(&'a BStr),
    /// Represents the 'SavePicture' keyword.
    ///
    /// Saves a graphic from the `Picture` or `Image` property of an object or
    /// control (if one is associated with it) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445827(v=vs.60))
    SavePictureKeyword(&'a BStr),
    /// Represents the 'RSet' keyword.
    ///
    /// Right aligns a string within a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266256(v=vs.60))
    RSetKeyword(&'a BStr),
    /// Represents the 'RmDir' keyword.
    ///
    /// Removes an existing directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266252(v=vs.60))
    RmDirKeyword(&'a BStr),
    /// Represents the 'Resume' keyword.
    ///
    /// Resumes execution after an error-handling routine is finished.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266247(v=vs.60))
    ResumeKeyword(&'a BStr),
    /// Represents the 'Reset' keyword.
    ///
    /// Closes all disk files opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266242(v=vs.60))
    ResetKeyword(&'a BStr),
    /// Represents the 'Rem' keyword.
    ///
    /// Used to include explanator remarks in a program.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266237(v=vs.60))
    RemKeyword(&'a BStr),
    /// Represents the 'Randomize' keyword.
    ///
    /// Initializes the random-number generator with a seed value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266225(v=vs.60))
    RandomizeKeyword(&'a BStr),
    /// Represents the 'RaiseEvent' keyword.
    ///
    /// Fires an event declared at module level within a class, form, or
    /// document.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266219(v=vs.60))
    RaiseEventKeyword(&'a BStr),
    /// Represents the 'Put' keyword.
    ///
    /// Writes data from a variable to a disk file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266212(v=vs.60))
    PutKeyword(&'a BStr),
    /// Represents the 'Property' keyword.
    ///
    /// Declares the name, argument, and code that forms the body of a property
    /// procedure, which sets a reference to a property of an object.
    ///
    /// Used in Property Get, Property Let, and Property Set statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266202(v=vs.60))
    PropertyKeyword(&'a BStr),
    /// Represents the 'Print' keyword.
    ///
    /// Writes display-formatted data to a sequential file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266187(v=vs.60))
    PrintKeyword(&'a BStr),
    /// Represents the 'Open' keyword.
    ///
    /// Enables input/output (I/O) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266177(v=vs.60))
    OpenKeyword(&'a BStr),
    /// Represents the 'On' keyword.
    ///
    /// Branch to one of several specified lines, dependin on the value of an expression.
    /// Used in the following contexts:
    ///
    /// On...GoSub statement, On...Goto statement, and On...Error statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266175(v=vs.60))
    OnKeyword(&'a BStr),
    /// Represents the 'Name' keyword.
    ///
    /// Renames a disk file, directory, or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266171(v=vs.60))
    NameKeyword(&'a BStr),
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
