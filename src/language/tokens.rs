/// Represents a VB6 token.
///
/// This is a simple enum that represents the different types of tokens that can be parsed from VB6 code.
/// The text content is now provided separately in a tuple when returned from tokenization.
///
#[derive(Debug, PartialEq, Clone, Copy, Eq, serde::Serialize)]
pub enum VB6Token {
    /// Represents whitespace.
    /// This is a collection of spaces, tabs, and other whitespace characters.
    Whitespace,
    /// Represents a newline.
    /// This can be a carriage return, a newline, or a carriage return followed by a newline.
    Newline,
    /// Represents a comment that runs to the end of the line.
    ///
    /// Includes the single quote character but not the newline character.
    EndOfLineComment,
    /// Represents the 'Class' keyword.
    ///
    /// Used in the header of a class module to indicate that the module is a class module.
    ClassKeyword,
    /// Represents the `ReDim` keyword.
    ///
    /// Used at a procedure level to reallocate storage space for a dynamic
    /// array.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    ReDimKeyword,
    /// Represents the `Preserve` keyword.
    ///
    /// Used with the `ReDim` keyword to preserve the contents of an array when
    /// reallocating storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266231(v=vs.60))
    PreserveKeyword,
    /// Represents the `Dim` keyword.
    ///
    /// Used to declare variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    DimKeyword,
    /// Represents the `Declare` keyword.
    ///
    /// Used at the module level to declare references to external procedures
    /// in a dynamic-link library (DLL).
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    DeclareKeyword,
    /// Represents the `Alias` keyword.
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
    AliasKeyword,
    /// Represents the `Attribute` keyword.
    ///
    /// Used to define metadata for a class, method, or property.
    AttributeKeyword,
    /// Represents the `Begin` keyword.
    ///
    /// Used to indicate the beginning of a block of code for a header section
    /// in a module, class, or form.
    BeginKeyword,
    /// Represents the `Lib` keyword.
    ///
    /// Indicates that a DLL or code resource contains the procedure being declared.
    /// The Lib clause is required for all declarations.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    LibKeyword,
    /// Represents the `With` keyword.
    ///
    /// Executes a series of statements on a single object or a user-defined type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266330(v=vs.60))
    WithKeyword,
    /// Represents the `WithEvents` keyword.
    ///
    /// Used with the `Dim` keyword to declare a variable that can respond to
    /// events raised by an object.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243352(v=vs.60))
    WithEventsKeyword,
    /// Represents the `Base` keyword.
    ///
    /// Used at module level to declare the default lower bound for array
    /// subscripts.
    ///
    ///[Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266179(v=vs.60))
    BaseKeyword,
    /// Represents the `Compare` keyword.
    ///
    /// Used at module level to declare the default comparison method to use
    /// when string data is compared.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266181(v=vs.60))
    CompareKeyword,
    /// Represents the `Option` keyword.
    ///
    /// Used at the module level in the Option Base, Option Compare, Option
    /// Explicit, or Option Private statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266185(v=vs.60))
    OptionKeyword,
    /// Represents the `Explicit` keyword.
    ///
    /// Used at the module level in the Option Explicit statement to force
    /// explicit declaration of all variables in that module.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266183(v=vs.60))
    ExplicitKeyword,
    /// Represents the `Private` keyword.
    ///
    /// Used at the module level to declare private variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266189(v=vs.60))
    PrivateKeyword,
    /// Represents the `Public` keyword.
    ///
    /// Used at the module level to declare public variables and allocate storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266207(v=vs.60))
    PublicKeyword,
    /// Represents the `Const` keyword.
    ///
    /// Declares constants for use in place of literal values.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243294(v=vs.60))
    ConstKeyword,
    /// Represents the `As` keyword.
    ///
    /// The `As` keyword is used in these contexts:
    /// `Const` statement, `Declare` statement, `Dim` statement, `Function` statement,
    /// `Name` statement, `Open` statement, `Private` statement,
    /// `Property Get` statement, `Property Let` statement, `Property Set` statement,
    /// `Public` statement, `ReDim` statement, `Static` statement, `Sub` statement, and
    /// `Type` statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445149(v=vs.60))
    AsKeyword,
    /// Represents the `ByVal` keyword.
    ///
    /// Used in the following contexts:
    /// `Call` statement, `Declare` statement, `Function` statement, `Property Get`
    /// statement, `Property Let` statement, `Property Set` statement, and `Sub`
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445152(v=vs.60))
    ByValKeyword,
    /// Represents the `ByRef` keyword.
    ///
    /// Used in the following contexts:
    /// `Call` statement, `Declare` statement, `Function` statement, `Property Get`
    /// statement, `Property Let` statement, `Property Set` statement, and `Sub`
    /// statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445151(v=vs.60))
    ByRefKeyword,
    /// Represents the `Optional` keyword.
    ///
    /// Used in the following contexts:
    /// `Declare` statement, `Function` statement, `Property Get` statement,
    /// `Property Let` statement, `Property Set` statement, and `Sub` statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445195(v=vs.60))
    OptionalKeyword,
    /// Represents the `Function` keyword.
    ///
    /// Used to declare the name, argument, and code that forms the body of a
    /// function procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243374(v=vs.60))
    FunctionKeyword,
    /// Represents the `Static` keyword.
    ///
    /// Used at the procedure level to declare variable and allocate storage space.
    /// Variables declared the with `Static` statement retain their values as long
    /// as the module is loaded in memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266296(v=vs.60))
    StaticKeyword,
    /// Represents the `Sub` keyword.
    ///
    /// Used to declare the name, argument, and code that form the body of a sub
    /// procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266305(v=vs.60))
    SubKeyword,
    /// Represents the `End` keyword.
    ///
    /// Used to end a procedure or block.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243356(v=vs.60))
    EndKeyword,
    /// Represents the `True` keyword.
    ///
    /// The `True` keyword is used to represent the boolean value true and has a
    /// value equal to -1.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445231(v=vs.60))
    TrueKeyword,
    /// Represents the `False` keyword.
    ///
    /// The `False` keyword is used to represent the boolean value false and has a
    /// value equal to 0.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445157(v=vs.60))
    FalseKeyword,
    /// Represents the `Enum` keyword.
    ///
    /// Used to declare a type for an enumeration.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243358(v=vs.60))
    EnumKeyword,
    /// Represents the `Type` keyword.
    ///
    /// Used at the module level to declare a user-defined data type containing
    /// one or more elements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266315(v=vs.60))
    TypeKeyword,
    /// Represents the `Boolean` keyword.
    ///
    /// Used to declare a variable that can contain one of two values: True or
    /// False.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    BooleanKeyword,
    /// Represents the `Double` keyword.
    ///
    /// Used to declare a variable that can contain a double-precision floating-point
    /// number.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DoubleKeyword,
    /// Represents the `Currency` keyword.
    ///
    /// Used to declare a variable that can contain a currency value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    CurrencyKeyword,
    /// Represents the `Decimal` keyword.
    ///
    /// Used to declare a variable that can contain a decimal value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DecimalKeyword,
    /// Represents the `Date` keyword.
    ///
    /// Used to declare a variable that can contain a date value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    DateKeyword,
    /// Represents the `Object` keyword.
    ///
    /// Used to declare a variable that can contain an object reference.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ObjectKeyword,
    /// Represents the `Variant` keyword.
    ///
    /// Used to declare a variable that can contain multiple kinds of types of
    /// data.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    VariantKeyword,
    /// Represents the `Byte` keyword.
    ///
    /// Used to declare a variable that can contain a byte value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    ByteKeyword,
    /// Represents the `Long` keyword.
    ///
    /// Used to declare a variable that can contain a long integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    LongKeyword,
    /// Represents the `Single` keyword.
    ///
    /// Used to declare a variable that can contain a single-precision
    /// floating-point value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    SingleKeyword,
    /// Represents the `String` keyword.
    ///
    /// Used to declare a variable that can contain a string value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    StringKeyword,
    /// Represents the `Integer` keyword.
    ///
    /// Used to declare a variable that can contain an integer value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263420(v=vs.60))
    IntegerKeyword,
    /// Represents a string literal.
    ///
    /// The string literal includes the enclosing double quotes.
    StringLiteral,
    /// Represents the `If` keyword.
    ///
    /// Used to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    IfKeyword,
    /// Represents the `Else` keyword.
    ///
    /// Used to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseKeyword,
    /// Represents the `ElseIf` keyword.
    ///
    /// Used to conditionally execute a block of code depending on the value
    /// of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243382(v=vs.60))
    ElseIfKeyword,
    /// Represents the `And` keyword.
    ///
    /// Used to perform a logical conjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242751(v=vs.60))
    AndKeyword,
    /// Represents the `Or` keyword.
    ///
    /// Used to perform a logical disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242850(v=vs.60))
    OrKeyword,
    /// Represents the `Xor` keyword.
    ///
    /// Used to perform a logical exclusive disjunction on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242859(v=vs.60))
    XorKeyword,
    /// Represents the `Mod` keyword.
    ///
    /// Used to perform a modulus operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242823(v=vs.60))
    ModKeyword,
    /// Represents the `Eqv` keyword.
    ///
    /// Used to perform a logical equivalence operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242780(v=vs.60))
    EqvKeyword,
    /// Represents the `AddressOf` keyword.
    ///
    /// A unary operator that obtains the address of the procedure it precedes
    /// and is used with API procedures that expect a function pointer at that
    /// position in the argument list.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242738(v=vs.60))
    AddressOfKeyword,
    /// Represents the `Imp` keyword.
    ///
    /// Used to perform a logical implication operation on two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242794(v=vs.60))
    ImpKeyword,
    /// Represents the `Is` keyword.
    ///
    /// Used to perform a reference comparison between two object variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242809(v=vs.60))
    IsKeyword,
    /// Represents the `Like` keyword.
    ///
    /// Used to compare two strings.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242817(v=vs.60))
    LikeKeyword,
    /// Represents the `Not` keyword.
    ///
    /// Used to perform a logical negation on an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242842(v=vs.60))
    NotKeyword,
    /// Represents the `Then` keyword.
    ///
    /// Used to indicate the start of a block of code that is executed if the
    /// condition in an `If` statement is true.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445224(v=vs.60))
    ThenKeyword,
    /// Represents the `Goto` keyword.
    ///
    /// Branches unconditionally to a specific line within a procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243380(v=vs.60))
    GotoKeyword,
    /// Represents the `GoSub` keyword.
    ///
    /// Branches to and returns from a subroutine within a procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/gosubreturn-statement)
    GoSubKeyword,
    /// Represents the `Return` keyword.
    ///
    /// Returns from a subroutine within a procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/gosubreturn-statement)
    ReturnKeyword,
    /// Represents the `Exit` keyword.
    ///
    /// Exits a block of `Do..Loop`, `For..Next`, `Function`, `Sub`, or `Property` code.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243366(v=vs.60))
    ExitKeyword,
    /// Represents the `For` keyword.
    ///
    /// Used to declare a `For..Next` loop, or a `For Each...Next` loop.
    /// Repeats a group of statements a specified number of times.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243370(v=vs.60))
    ForKeyword,
    /// Represents the `To` keyword.
    ///
    /// The `To` keyword is used in these contexts:
    ///
    /// `Dim` statement, `For...Next` statement, `Lock` statement, `Unlock` statement,
    /// `Private` statement, `Public` statement, `ReDim` statement, `Select Case` statement,
    /// `Static` statement, and `Type` statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445229(v=vs.60))
    ToKeyword,
    /// Represents the `Lock` keyword.
    ///
    /// Controls access by other processes to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    LockKeyword,
    /// Represents the `Unlock` keyword.
    ///
    /// Controls access by other processes to all or part of a file opened using
    /// the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266161(v=vs.60))
    UnlockKeyword,
    /// Represents the `Step` keyword.
    ///
    /// Used in the `For...Next` statement to specify the increment of the loop
    /// variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445219(v=vs.60))
    StepKeyword,
    /// Represents the `Stop` keyword.
    ///
    /// Used to suspend execution of a program.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266300(v=vs.60))
    StopKeyword,
    /// Represents the `While` keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WhileKeyword,
    /// Represents the `Wend` keyword.
    ///
    /// Used to execute a series of statements as long as a given condition is
    /// True.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266320(v=vs.60))
    WendKeyword,
    /// Represents the `Width` keyword.
    ///
    /// Assigns an output line width to a file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266324(v=vs.60))
    WidthKeyword,
    /// Represents the `Write` keyword.
    ///
    /// Used to write data to a sequential file opened with the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266338(v=vs.60))
    WriteKeyword,
    /// Represents the `Time` keyword.
    ///
    /// Used to set the System time.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266310(v=vs.60))
    TimeKeyword,
    /// Represents the `SetAttr` keyword.
    ///
    /// Used to set attribute information for a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266286(v=vs.60))
    SetAttrKeyword,
    /// Represents the `Set` keyword.
    ///
    /// Used to assign an object reference to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266283(v=vs.60))
    SetKeyword,
    /// Represents the `SendKeys` keyword.
    ///
    /// Used to send one or more keystrokes to the active window as if typed at
    /// the keyboard.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266279(v=vs.60))
    SendKeysKeyword,
    /// Represents the `Select` keyword.
    ///
    /// Used to execute one of a several groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    SelectKeyword,
    /// Represents the `Case` keyword.
    ///
    /// Used to execute one of a several groups of statements, depending on the
    /// value of an expression.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266274(v=vs.60))
    CaseKeyword,
    /// Represents the `Seek` keyword.
    ///
    /// Used to set the position for the next read/write operation on a file
    /// opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266268(v=vs.60))
    SeekKeyword,
    /// Represents the `SaveSetting` keyword.
    ///
    /// Saves or creates an application entry in the application's entry in the Windows registry.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266261(v=vs.60))
    SaveSettingKeyword,
    /// Represents the `SavePicture` keyword.
    ///
    /// Saves a graphic from the `Picture` or `Image` property of an object or
    /// control (if one is associated with it) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445827(v=vs.60))
    SavePictureKeyword,
    /// Represents the `RSet` keyword.
    ///
    /// Right aligns a string within a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266256(v=vs.60))
    RSetKeyword,
    /// Represents the `RmDir` keyword.
    ///
    /// Removes an existing directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266252(v=vs.60))
    RmDirKeyword,
    /// Represents the `Resume` keyword.
    ///
    /// Resumes execution after an error-handling routine is finished.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266247(v=vs.60))
    ResumeKeyword,
    /// Represents the `Reset` keyword.
    ///
    /// Closes all disk files opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266242(v=vs.60))
    ResetKeyword,
    /// Represents a `REM` line comment.
    ///
    /// Includes the `REM` characters and the comment text but not the newline.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266237(v=vs.60))
    RemComment,
    /// Represents the `Randomize` keyword.
    ///
    /// Initializes the random-number generator with a seed value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266225(v=vs.60))
    RandomizeKeyword,
    /// Represents the `RaiseEvent` keyword.
    ///
    /// Fires an event declared at module level within a class, form, or
    /// document.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266219(v=vs.60))
    RaiseEventKeyword,
    /// Represents the `Put` keyword.
    ///
    /// Writes data from a variable to a disk file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266212(v=vs.60))
    PutKeyword,
    /// Represents the `Property` keyword.
    ///
    /// Declares the name, argument, and code that forms the body of a property
    /// procedure, which sets a reference to a property of an object.
    ///
    /// Used in `Property Get`, `Property Let`, and `Property Set` statements.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266202(v=vs.60))
    PropertyKeyword,
    /// Represents the `Print` keyword.
    ///
    /// Writes display-formatted data to a sequential file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266187(v=vs.60))
    PrintKeyword,
    /// Represents the `Open` keyword.
    ///
    /// Enables input/output (I/O) to a file.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266177(v=vs.60))
    OpenKeyword,
    /// Represents the `On` keyword.
    ///
    /// Branch to one of several specified lines, depending on the value of an expression.
    /// Used in the following contexts:
    ///
    /// `On...GoSub` statement, `On...Goto` statement, and `On...Error` statements.
    ///
    /// Also used when specifying `Option Explicit On` or `Option Explicit Off`.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266175(v=vs.60))
    OnKeyword,
    /// Represents the `Off` keyword.
    ///
    /// Used when specifying `Option Explicit On` or `Option Explicit Off`.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266175(v=vs.60))
    OffKeyword,
    /// Represents the `Name` keyword.
    ///
    /// Renames a disk file, directory, or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266171(v=vs.60))
    NameKeyword,
    /// Represents the `MkDir` keyword.
    ///
    /// Creates a new directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266169(v=vs.60))
    MkDirKeyword,
    /// Represents the `Mid` keyword.
    ///
    /// Replaces a specified number of characters in a `Variant` (String) variable
    /// with characters from another string.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266166(v=vs.60))
    MidKeyword,
    /// Represents the `MidB` keyword.
    ///
    /// Used in `MidB` statements to replace bytes in a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/midb-statement)
    MidBKeyword,
    /// Represents the `LSet` keyword.
    ///
    /// Left alligns a string within a string variable, or copies a variable of
    /// one user-defined type to another variable of a different user-defined type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa266163(v=vs.60))
    LSetKeyword,
    /// Represents the `Load` keyword.
    ///
    /// Loads a form or control into memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445825(v=vs.60))
    LoadKeyword,
    /// Represents the `Unload` keyword.
    ///
    /// Removes a form or control from memory.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unload-statement)
    UnloadKeyword,
    /// Represents the `Line` keyword.
    ///
    /// Reads a single line from an open sequential file and assigns it to a string variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243392(v=vs.60))
    LineKeyword,
    /// Represents the `Input` keyword.
    ///
    /// Reads data from an open sequential file and assigns the data to variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243386(v=vs.60))
    InputKeyword,
    //// Represents the `Let` keyword.
    ///
    /// Assigns the value of an expression to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243390(v=vs.60))
    LetKeyword,
    /// Represents the `Kill` keyword.
    ///
    /// Deletes files from a disk.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243388(v=vs.60))
    KillKeyword,
    /// Represents the `Implements` keyword.
    ///
    /// Specifies an interface or class that will be implemented in the class module in which it appears.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243384(v=vs.60))
    ImplementsKeyword,
    /// Represents the `Get` keyword.
    ///
    /// Reads data from an open disk file into a variable.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243376(v=vs.60))
    GetKeyword,
    /// Represents the `FileCopy` keyword.
    ///
    /// Copies a file from one location to another.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243368(v=vs.60))
    FileCopyKeyword,
    /// Represents the `Event` keyword.
    ///
    /// Declares a user-defined event.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243364(v=vs.60))
    EventKeyword,
    /// Represents the `Error` keyword.
    ///
    /// Simulates the occurance of an error.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243362(v=vs.60))
    ErrorKeyword,
    /// Represents the `Erase` keyword.
    ///
    /// Reinitializes the elements of a fixed-size array and releases dynamic-array
    /// storage space.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243360(v=vs.60))
    EraseKeyword,
    /// Represents the `Do` keyword.
    ///
    /// Repeats a block of statements while a condition is `True` or until a
    /// condition becomes `True`.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243354(v=vs.60))
    DoKeyword,
    /// Represents the `Until` keyword.
    ///
    /// Used in the `Do...Loop` statement to specify the condition under which
    /// the loop terminates.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243354(v=vs.60))
    UntilKeyword,
    /// Represents the `Loop` keyword.
    ///
    /// Used with the `Do` keyword to terminate a `Do...Loop` statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243354(v=vs.60))
    LoopKeyword,
    /// Represents the `DeleteSetting` keyword.
    ///
    /// Deletes a section or key setting from an application's entry in the
    /// Windows registry.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243347(v=vs.60))
    DeleteSettingKeyword,
    /// Represents the `DefBool` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Boolean data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefBoolKeyword,
    /// Represents the `DefByte` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Byte data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefByteKeyword,
    /// Represents the `DefInt` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Int data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefIntKeyword,
    /// Represents the `DefLng` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Long data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefLngKeyword,
    /// Represents the `DefCur` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Currency data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefCurKeyword,
    /// Represents the `DefSng` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Single data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefSngKeyword,
    /// Represents the `DefDbl` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Double data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDblKeyword,
    /// Represents the `DefDec` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Decimal data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDecKeyword,
    /// Represents the `DefDate` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Date data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefDateKeyword,
    /// Represents the `DefStr` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the String data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefStrKeyword,
    /// Represents the `DefObj` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Object data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefObjKeyword,
    /// Represents the `DefVar` keyword.
    ///
    /// Used at module level to set the default data type for variables,
    /// arguments passed to procedures, and the return type for `Function` and
    /// `PropertyGet` procedures whose names start with the specified characters.
    ///
    /// Sets the default for a variable to the Variant data type.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa263421(v=vs.60))
    DefVarKeyword,
    /// Represents the `Close` keyword.
    ///
    /// Concludes input/output (I/O) to a file opened using the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243283(v=vs.60))
    CloseKeyword,
    /// Represents the `ChDir` keyword.
    ///
    /// Changes the current drive.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243270(v=vs.60))
    ChDriveKeyword,
    /// Represents the `ChDir` keyword.
    ///
    /// Changes the current directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243256(v=vs.60))
    ChDirKeyword,
    /// Represents the `Call` keyword.
    ///
    /// Transfers control to a `Sub` procedure, `Function` procedure, or dynamic-link
    /// library (DLL) procedure.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243242(v=vs.60))
    CallKeyword,
    /// Represents the `Beep` keyword.
    ///
    /// Sounds a tone through the computer's speaker.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243233(v=vs.60))
    BeepKeyword,
    /// Represents the `AppActivate` keyword.
    ///
    /// Activates an application window.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243211(v=vs.60))
    AppActivateKeyword,
    /// Represents the `Friend` keyword.
    ///
    /// Modifies the definition of a procedure in a form module or class module
    /// to make the procedure callable from modules that are outside the class,
    /// but part of the project within which the class is defined. `Friend`
    /// procedures cannot be used in standard modules.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445159(v=vs.60))
    FriendKeyword,
    /// Represents the `Binary` keyword.
    ///
    /// The `Binary` keyword is used in these contexts:
    /// `Open` statement, `Option Compare` statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445150(v=vs.60))
    BinaryKeyword,
    /// Represents the `Random` keyword.
    ///
    /// The `Random` keyword is used in the `Open` statement to specify random access mode.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    RandomKeyword,
    /// Represents the `Read` keyword.
    ///
    /// The `Read` keyword is used in the `Open` statement to specify read access mode.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    ReadKeyword,
    /// Represents the `Output` keyword.
    ///
    /// The Output keyword is used in the Open statement to specify output mode.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    OutputKeyword,
    /// Represents the 'Append' keyword.
    ///
    /// The Append keyword is used in the Open statement to specify append mode.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    AppendKeyword,
    /// Represents the 'Access' keyword.
    ///
    /// The Access keyword is used in the Open statement to specify access restrictions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/open-statement)
    AccessKeyword,
    /// Represents the 'Text' keyword.
    ///
    /// The Text keyword is used in the Option Compare statement to specify text-based string comparison.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-compare-statement)
    TextKeyword,
    /// Represents the 'Database' keyword.
    ///
    /// The Database keyword is used in the Option Compare statement to specify database-based string comparison.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-compare-statement)
    DatabaseKeyword,
    /// Represents the 'Empty' keyword.
    ///
    /// The Empty keyword is used as a Variant subtype. It indicates an
    /// uninitialized variable value.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445155(v=vs.60))
    EmptyKeyword,
    /// Represents the 'Module' keyword.
    ///
    /// The Module keyword is used in the Option Private statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-private-statement)
    ModuleKeyword,
    /// Represents the 'Next' keyword.
    ///
    /// The next keyword is used in these contexts:
    ///
    /// For...Next statement, For Each...Next statement, On Error statement, and
    /// a Resume statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445184(v=vs.60))
    NextKeyword,
    /// Represents the 'Each' keyword.
    ///
    /// The Each keyword is used in these contexts:
    ///
    /// For Each element In group...Next statement
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243370(v=vs.60))
    EachKeyword,
    /// Represents the 'In' keyword.
    ///
    /// The In keyword is used in these contexts:
    ///
    /// For Each element In group...Next statement
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243370(v=vs.60))
    InKeyword,
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
    NewKeyword,
    /// Represents the 'Len' keyword.
    ///
    /// The Len keyword is used in these contexts:
    ///
    /// Len Function, and the Open statement.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445169(v=vs.60))
    LenKeyword,
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
    MeKeyword,
    /// Represents the 'Null' keyword.
    ///
    /// The Null keyword is used as a Variant subtype. It indicates that a
    /// variable contains no valid data.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445190(v=vs.60))
    NullKeyword,
    /// Represents the `ParamArray` keyword.
    ///
    /// The `ParamArray` keyword is used in these contexts:
    ///
    /// `Declare` statement, `Function` statement, `Property Get` statement,
    /// `Property Let` statement, and `Sub` statement.
    ///
    /// Used only as the last argument in arglist to indicate that the final
    /// argument is an `Optional` array of `Variant` elements. The `ParamArray`
    /// keyword allows you to provide an arbitrary number of arguments. It may
    /// not be used with `ByVal`, `ByRef`, or `Optional`.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445198(v=vs.60))
    ParamArrayKeyword,
    /// Represents a dollar sign '$'.
    ///
    /// Often used to indicate a variable is a string or that a function
    /// works with strings.
    DollarSign,
    /// Represents an underscore `_`.
    ///
    /// Used to indicate that a statement continues on the next line.
    /// It must be preceded by at least one whitespace and must be the last
    /// character on the line.
    Underscore,
    /// Represents an ampersand `&`.
    ///
    /// Used to force string concatenation of two expressions.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242763(v=vs.60))
    Ampersand,
    /// Represents a percent sign `%`.
    Percent,
    /// Represents an octothorpe `#`.
    Octothorpe,
    /// Represents a left parenthesis `(`.
    LeftParenthesis,
    /// Represents a right parenthesis `)`.
    RightParenthesis,
    /// Represents a left square bracket `[`.
    LeftSquareBracket,
    /// Represents a right square bracket `]`.
    RightSquareBracket,
    /// Represents a comma `,`.
    Comma,
    /// Represents a semicolon `;`.
    Semicolon,
    /// Represents the 'at' symbol `@`.
    AtSign,
    /// Represents an exclamation mark `!`.
    ExclamationMark,
    /// Represents the `Version` keyword.
    ///
    /// The `Version` keyword is used to specify the version of the header
    /// information for a module / class / form.
    VersionKeyword,
    /// Represents an equality operator `=` can also be the assignment operator.
    ///
    /// Used to assign a value to a variable or property.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa242760(v=vs.60))
    EqualityOperator,
    /// Represents a less than operator `<`.
    LessThanOperator,
    /// Represents a greater than operator `>`.
    GreaterThanOperator,
    /// Represents a multiplication operator `*`.
    MultiplicationOperator,
    /// Represents a subtraction operator `-`.
    SubtractionOperator,
    /// Represents an addition operator `+`.
    AdditionOperator,
    /// Represents a division operator `/`.
    DivisionOperator,
    /// Represents a backward slash operator `\\`.
    BackwardSlashOperator,
    /// Represents a period operator `.`.
    PeriodOperator,
    /// Represents a colon operator `:`.
    ColonOperator,
    /// Represents an exponentiation operator `^`.
    ExponentiationOperator,
    /// Represents an Identifier, variable, or function name.
    /// This is a name that starts with a letter and can contain letters, numbers, and underscores.
    Identifier,
    /// Represents an `Integer` literal with `%` suffix or plain integer (e.g., 42, 42%).
    IntegerLiteral,
    /// Represents a `Long` literal with `&` suffix (e.g., 42&).
    LongLiteral,
    /// Represents a `Single` (float) literal with `!` suffix or decimal/exponent without suffix (e.g., 3.14, 3.14!, 1.5E+10).
    SingleLiteral,
    /// Represents a `Double` literal with `#` suffix (e.g., 3.14#).
    DoubleLiteral,
    /// Represents a `Decimal` literal with `@` suffix (e.g., 12.50@).
    DecimalLiteral,
    /// Represents a `Date` literal with `#` delimiters (e.g., #1/1/2000#).
    DateLiteral,
}

impl VB6Token {

    /// Returns true if the token is a VB6 operator.
    #[must_use]
    pub fn is_operator(&self) -> bool {
        matches!(self,
            VB6Token::EqualityOperator
            | VB6Token::LessThanOperator
            | VB6Token::GreaterThanOperator
            | VB6Token::MultiplicationOperator
            | VB6Token::SubtractionOperator
            | VB6Token::AdditionOperator
            | VB6Token::DivisionOperator
            | VB6Token::BackwardSlashOperator
            | VB6Token::PeriodOperator
            | VB6Token::ColonOperator
            | VB6Token::ExponentiationOperator
            | VB6Token::Ampersand)
    }

    /// Returns true if the token is a VB6 keyword.
    #[must_use]
    pub fn is_keyword(&self) -> bool {
        matches!(self,
            VB6Token::AddressOfKeyword
            | VB6Token::ImpKeyword
            | VB6Token::IsKeyword
            | VB6Token::LikeKeyword
            | VB6Token::NotKeyword
            | VB6Token::ThenKeyword
            | VB6Token::GotoKeyword
            | VB6Token::GoSubKeyword
            | VB6Token::ReturnKeyword
            | VB6Token::ExitKeyword
            | VB6Token::ForKeyword
            | VB6Token::ToKeyword
            | VB6Token::LockKeyword
            | VB6Token::UnlockKeyword
            | VB6Token::StepKeyword
            | VB6Token::StopKeyword
            | VB6Token::WhileKeyword
            | VB6Token::WendKeyword
            | VB6Token::WidthKeyword
            | VB6Token::WriteKeyword
            | VB6Token::TimeKeyword
            | VB6Token::SetAttrKeyword
            | VB6Token::SetKeyword
            | VB6Token::SendKeysKeyword
            | VB6Token::SelectKeyword
            | VB6Token::CaseKeyword
            | VB6Token::SeekKeyword
            | VB6Token::SaveSettingKeyword
            | VB6Token::SavePictureKeyword
            | VB6Token::RSetKeyword
            | VB6Token::RmDirKeyword
            | VB6Token::ResumeKeyword
            | VB6Token::ResetKeyword
            | VB6Token::RemComment
            | VB6Token::RandomizeKeyword
            | VB6Token::RaiseEventKeyword
            | VB6Token::PutKeyword
            | VB6Token::PropertyKeyword
            | VB6Token::PrintKeyword
            | VB6Token::OpenKeyword
            | VB6Token::OnKeyword
            | VB6Token::OffKeyword
            | VB6Token::NameKeyword
            | VB6Token::MkDirKeyword
            | VB6Token::MidBKeyword
            | VB6Token::MidKeyword
            | VB6Token::LSetKeyword
            | VB6Token::LoadKeyword
            | VB6Token::UnloadKeyword
            | VB6Token::LineKeyword
            | VB6Token::InputKeyword
            | VB6Token::LetKeyword
            | VB6Token::KillKeyword
            | VB6Token::ImplementsKeyword
            | VB6Token::GetKeyword
            | VB6Token::FileCopyKeyword
            | VB6Token::EventKeyword
            | VB6Token::ErrorKeyword
            | VB6Token::EraseKeyword
            | VB6Token::DoKeyword
            | VB6Token::UntilKeyword
            | VB6Token::LoopKeyword
            | VB6Token::DeleteSettingKeyword
            | VB6Token::DefBoolKeyword
            | VB6Token::DefByteKeyword
            | VB6Token::DefIntKeyword
            | VB6Token::DefLngKeyword
            | VB6Token::DefCurKeyword
            | VB6Token::DefSngKeyword
            | VB6Token::DefDblKeyword
            | VB6Token::DefDecKeyword
            | VB6Token::DefDateKeyword
            | VB6Token::DefStrKeyword
            | VB6Token::DefObjKeyword
            | VB6Token::DefVarKeyword
            | VB6Token::CloseKeyword
            | VB6Token::ChDriveKeyword
            | VB6Token::ChDirKeyword
            | VB6Token::CallKeyword
            | VB6Token::BeepKeyword
            | VB6Token::AppActivateKeyword
            | VB6Token::FriendKeyword
            | VB6Token::BinaryKeyword
            | VB6Token::RandomKeyword
            | VB6Token::ReadKeyword
            | VB6Token::OutputKeyword
            | VB6Token::AppendKeyword
            | VB6Token::AccessKeyword
            | VB6Token::TextKeyword
            | VB6Token::DatabaseKeyword
            | VB6Token::EmptyKeyword
            | VB6Token::ModuleKeyword
            | VB6Token::NextKeyword
            | VB6Token::EachKeyword
            | VB6Token::InKeyword
            | VB6Token::NewKeyword
            | VB6Token::LenKeyword
            | VB6Token::MeKeyword
            | VB6Token::NullKeyword
            | VB6Token::ParamArrayKeyword
            | VB6Token::VersionKeyword)
    }
}
