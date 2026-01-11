//! Syntax kinds for the VB6 CST.
//!
//! This module defines the `SyntaxKind` enum, which represents all possible
//! node and token types in the Concrete Syntax Tree (CST) for VB6 code.
//!
//! Each variant of the enum corresponds to a specific syntactic construct
//! in the VB6 language, including statements, expressions, keywords, literals,
//! operators, and punctuation.
//!
//! The `SyntaxKind` enum is used throughout the parser to identify and
//! categorize different parts of the syntax tree.

use std::fmt::Display;

use crate::language::Token;

/// Syntax kinds for the VB6 CST.
///
/// This enum represents all possible node and token types in the CST.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, serde::Serialize)]
#[repr(u16)]
pub enum SyntaxKind {
    /// Root node of the syntax tree
    Root = 0,

    // Statement nodes
    /// Module definition statement
    ModuleStatement,
    /// Class definition statement
    ClassStatement,
    /// Sub procedure statement
    SubStatement,
    /// Function procedure statement
    FunctionStatement,
    /// Property procedure statement
    PropertyStatement,
    /// Declare statement
    DeclareStatement,
    /// Event declaration statement
    EventStatement,
    /// Implements statement
    ImplementsStatement,
    /// `DefType` statement
    DefTypeStatement,
    /// Dim statement
    DimStatement,
    /// `ReDim` statement
    ReDimStatement,
    /// Erase statement
    EraseStatement,
    /// Const statement
    ConstStatement,
    /// Type statement
    TypeStatement,
    /// Enum statement
    EnumStatement,
    /// If statement
    IfStatement,
    /// `ElseIf` clause of an If statement
    ElseIfClause,
    /// Else clause of an If statement
    ElseClause,
    /// For statement
    ForStatement,
    /// For Each statement
    ForEachStatement,
    /// While statement
    WhileStatement,
    /// Do statement
    DoStatement,
    /// Select Case statement
    SelectCaseStatement,
    /// Case clause of a Select Case statement
    CaseClause,
    /// Case Else clause of a Select Case statement
    CaseElseClause,
    /// With statement
    WithStatement,
    /// Call statement
    CallStatement,
    /// `RaiseEvent` statement
    RaiseEventStatement,
    /// Set statement
    SetStatement,
    /// Let statement
    LetStatement,
    /// Assignment statement
    AssignmentStatement,
    /// Goto statement
    GotoStatement,
    /// `GoSub` statement
    GoSubStatement,
    /// Return statement
    ReturnStatement,
    /// Resume statement
    ResumeStatement,
    /// Exit statement
    ExitStatement,
    /// On Error statement
    OnErrorStatement,
    /// On `GoTo` statement
    OnGoToStatement,
    /// On `GoSub` statement
    OnGoSubStatement,
    /// `AppActivate` statement
    AppActivateStatement,
    /// Beep statement
    BeepStatement,
    /// `ChDir` statement
    ChDirStatement,
    /// `ChDrive` statement
    ChDriveStatement,
    /// Close statement
    CloseStatement,
    /// Date statement
    DateStatement,
    /// `DeleteSetting` statement
    DeleteSettingStatement,
    /// Reset statement
    ResetStatement,
    /// `SavePicture` statement
    SavePictureStatement,
    /// `SaveSetting` statement
    SaveSettingStatement,
    /// Seek statement
    SeekStatement,
    /// `SendKeys` statement
    SendKeysStatement,
    /// `SetAttr` statement
    SetAttrStatement,
    /// Stop statement
    StopStatement,
    /// Time statement
    TimeStatement,
    /// Randomize statement
    RandomizeStatement,
    /// Error statement
    ErrorStatement,
    /// `FileCopy` statement
    FileCopyStatement,
    /// Get statement
    GetStatement,
    /// Put statement
    PutStatement,
    /// Input statement
    InputStatement,
    /// `LineInput` statement
    LineInputStatement,
    /// Kill statement
    KillStatement,
    /// Load statement
    LoadStatement,
    /// Unload statement
    UnloadStatement,
    /// Lock statement
    LockStatement,
    /// Unlock statement
    UnlockStatement,
    /// `LSet` statement
    LSetStatement,
    /// `RSet` statement
    RSetStatement,
    /// Mid statement
    MidStatement,
    /// `MidB` statement
    MidBStatement,
    /// `MkDir` statement
    MkDirStatement,
    /// `RmDir` statement
    RmDirStatement,
    /// Name statement
    NameStatement,
    /// Open statement
    OpenStatement,
    /// Print statement
    PrintStatement,
    /// Width statement
    WidthStatement,
    /// Write statement
    WriteStatement,
    /// Label statement
    LabelStatement,
    /// Attribute statement
    AttributeStatement,
    /// Option statement
    OptionStatement,
    /// Object statement
    ObjectStatement,

    // Class/Form Header nodes
    /// Version statement
    VersionStatement,
    /// Properties block
    PropertiesBlock,
    /// Properties type declaration
    PropertiesType,
    /// Properties name declaration
    PropertiesName,
    /// Property Statement
    Property,
    /// Property key
    PropertyKey,
    /// Property value
    PropertyValue,
    /// Property group
    PropertyGroup,
    /// Property group name
    PropertyGroupName,

    // Expression nodes
    /// Binary expression
    BinaryExpression,
    /// Unary expression
    UnaryExpression,
    /// Literal expression
    LiteralExpression,
    /// Identifier expression
    IdentifierExpression,
    /// Member access expression
    MemberAccessExpression,
    /// Call expression
    CallExpression,
    /// Parenthesized expression
    ParenthesizedExpression,
    /// Numeric literal expression
    NumericLiteralExpression,
    /// String literal expression
    StringLiteralExpression,
    /// Boolean literal expression
    BooleanLiteralExpression,
    /// `AddressOf` expression
    AddressOfExpression,
    /// `TypeOf` expression
    TypeOfExpression,
    /// New expression
    NewExpression,

    // Other structural nodes
    /// Argument list
    ArgumentList,
    /// Parameter list
    ParameterList,
    /// Parameter
    Parameter,
    /// Argument
    Argument,
    /// Block of code/statements
    StatementList,

    // Token kinds - map from Token
    // We start these at a higher offset to avoid conflicts
    /// Whitespace token
    Whitespace = 1000,
    /// Newline token
    Newline,
    /// End-of-line comment token
    EndOfLineComment,
    /// Rem comment token
    RemComment,

    // Keywords
    /// Class keyword
    ClassKeyword,
    /// `ReDim` keyword
    ReDimKeyword,
    /// Preserve keyword
    PreserveKeyword,
    /// Dim keyword
    DimKeyword,
    /// Declare keyword
    DeclareKeyword,
    /// Alias keyword
    AliasKeyword,
    /// Attribute keyword
    AttributeKeyword,
    /// Begin keyword
    BeginKeyword,
    /// Lib keyword
    LibKeyword,
    /// With keyword
    WithKeyword,
    /// `WithEvents` keyword
    WithEventsKeyword,
    /// Base keyword
    BaseKeyword,
    /// Compare keyword
    CompareKeyword,
    /// Option keyword
    OptionKeyword,
    /// Explicit keyword
    ExplicitKeyword,
    /// Private keyword
    PrivateKeyword,
    /// Public keyword
    PublicKeyword,
    /// Const keyword
    ConstKeyword,
    /// As keyword
    AsKeyword,
    /// `ByVal` keyword
    ByValKeyword,
    /// `ByRef` keyword
    ByRefKeyword,
    /// Optional keyword
    OptionalKeyword,
    /// Function keyword
    FunctionKeyword,
    /// Static keyword
    StaticKeyword,
    /// Sub keyword
    SubKeyword,
    /// End keyword
    EndKeyword,
    /// True keyword
    TrueKeyword,
    /// False keyword
    FalseKeyword,
    /// Enum keyword
    EnumKeyword,
    /// Type keyword
    TypeKeyword,
    /// Boolean keyword
    BooleanKeyword,
    /// Double keyword
    DoubleKeyword,
    /// Currency keyword
    CurrencyKeyword,
    /// Decimal keyword
    DecimalKeyword,
    /// Date keyword
    DateKeyword,
    /// Object keyword
    ObjectKeyword,
    /// Variant keyword
    VariantKeyword,
    /// Byte keyword
    ByteKeyword,
    /// Long keyword
    LongKeyword,
    /// Single keyword
    SingleKeyword,
    /// String keyword
    StringKeyword,
    /// Integer keyword
    IntegerKeyword,
    /// If keyword
    IfKeyword,
    /// Else keyword
    ElseKeyword,
    /// `ElseIf` keyword
    ElseIfKeyword,
    /// And keyword
    AndKeyword,
    /// Or keyword
    OrKeyword,
    /// Xor keyword
    XorKeyword,
    /// Mod keyword
    ModKeyword,
    /// Eqv keyword
    EqvKeyword,
    /// `AddressOf` keyword
    AddressOfKeyword,
    /// Imp keyword
    ImpKeyword,
    /// Is keyword
    IsKeyword,
    /// Like keyword
    LikeKeyword,
    /// Not keyword
    NotKeyword,
    /// Then keyword
    ThenKeyword,
    /// Goto keyword
    GotoKeyword,
    /// `GoSub` keyword
    GoSubKeyword,
    /// Return keyword
    ReturnKeyword,
    /// Exit keyword
    ExitKeyword,
    /// For keyword
    ForKeyword,
    /// Each keyword
    EachKeyword,
    /// In keyword
    InKeyword,
    /// To keyword
    ToKeyword,
    /// Lock keyword
    LockKeyword,
    /// Unlock keyword
    UnlockKeyword,
    /// Step keyword
    StepKeyword,
    /// Stop keyword
    StopKeyword,
    /// While keyword
    WhileKeyword,
    /// Wend keyword
    WendKeyword,
    /// Width keyword
    WidthKeyword,
    /// Write keyword
    WriteKeyword,
    /// Time keyword
    TimeKeyword,
    /// `SetAttr` keyword
    SetAttrKeyword,
    /// Set keyword
    SetKeyword,
    /// `SendKeys` keyword
    SendKeysKeyword,
    /// Select keyword
    SelectKeyword,
    /// Case keyword
    CaseKeyword,
    /// Seek keyword
    SeekKeyword,
    /// `SaveSetting` keyword
    SaveSettingKeyword,
    /// `SavePicture` keyword
    SavePictureKeyword,
    /// `RSet` keyword
    RSetKeyword,
    /// `RmDir` keyword
    RmDirKeyword,
    /// Resume keyword
    ResumeKeyword,
    /// Reset keyword
    ResetKeyword,
    /// Randomize keyword
    RandomizeKeyword,
    /// `RaiseEvent` keyword
    RaiseEventKeyword,
    /// Put keyword
    PutKeyword,
    /// Property keyword
    PropertyKeyword,
    /// Print keyword
    PrintKeyword,
    /// Open keyword
    OpenKeyword,
    /// On keyword
    OnKeyword,
    /// Off keyword
    OffKeyword,
    /// Name keyword
    NameKeyword,
    /// `MkDir` keyword
    MkDirKeyword,
    /// Mid keyword
    MidKeyword,
    /// `MidB` keyword
    MidBKeyword,
    /// `LSet` keyword
    LSetKeyword,
    /// Load keyword
    LoadKeyword,
    /// Unload keyword
    UnloadKeyword,
    /// Line keyword
    LineKeyword,
    /// Input keyword
    InputKeyword,
    /// Let keyword
    LetKeyword,
    /// Kill keyword
    KillKeyword,
    /// Implies keyword
    ImplementsKeyword,
    /// Get keyword
    GetKeyword,
    /// `FileCopy` keyword
    FileCopyKeyword,
    /// Event keyword
    EventKeyword,
    /// Error keyword
    ErrorKeyword,
    /// Erase keyword
    EraseKeyword,
    /// Do keyword
    DoKeyword,
    /// Until keyword
    UntilKeyword,
    /// `DeleteSetting` keyword
    DeleteSettingKeyword,
    /// `DefBool` keyword
    DefBoolKeyword,
    /// `DefByte` keyword
    DefByteKeyword,
    /// `DefInt` keyword
    DefIntKeyword,
    /// `DefLng` keyword
    DefLngKeyword,
    /// `DefCur` keyword
    DefCurKeyword,
    /// `DefSng` keyword
    DefSngKeyword,
    /// `DefDbl` keyword
    DefDblKeyword,
    /// `DefDec` keyword
    DefDecKeyword,
    /// `DefDate` keyword
    DefDateKeyword,
    /// `DefStr` keyword
    DefStrKeyword,
    /// `DefObj` keyword
    DefObjKeyword,
    /// `DefVar` keyword
    DefVarKeyword,
    /// Close keyword
    CloseKeyword,
    /// `ChDrive` keyword
    ChDriveKeyword,
    /// `ChDir` keyword
    ChDirKeyword,
    /// Call keyword
    CallKeyword,
    /// Beep keyword
    BeepKeyword,
    /// `AppActivate` keyword
    AppActivateKeyword,
    /// Friend keyword
    FriendKeyword,
    /// Binary keyword
    BinaryKeyword,
    /// Random keyword
    RandomKeyword,
    /// Read keyword
    ReadKeyword,
    /// Output keyword
    OutputKeyword,
    /// Append keyword
    AppendKeyword,
    /// Access keyword
    AccessKeyword,
    /// Text keyword
    TextKeyword,
    /// Database keyword
    DatabaseKeyword,
    /// Empty keyword
    EmptyKeyword,
    /// Module keyword
    ModuleKeyword,
    /// Next keyword
    NextKeyword,
    /// New keyword
    NewKeyword,
    /// Len keyword
    LenKeyword,
    /// Me keyword
    MeKeyword,
    /// Null keyword
    NullKeyword,
    /// `ParamArray` keyword
    ParamArrayKeyword,
    /// Version keyword
    VersionKeyword,
    /// Loop keyword
    LoopKeyword,
    /// Nothing keyword
    NothingKeyword,
    /// Any keyword
    AnyKeyword,

    // Literals and identifiers
    /// Identifier token
    Identifier,
    /// String literal token
    StringLiteral,
    /// Integer literal token
    IntegerLiteral,
    /// Long literal token
    LongLiteral,
    /// Single literal token
    SingleLiteral,
    /// Double literal token
    DoubleLiteral,
    /// Decimal literal token
    DecimalLiteral,
    /// Currency literal token
    CurrencyLiteral,
    /// Date literal token
    DateLiteral,

    // Operators and punctuation
    /// Dollar sign token
    DollarSign,
    /// Underscore token
    Underscore,
    /// Ampersand token
    Ampersand,
    /// Percent token
    Percent,
    /// Octothorpe '#' token
    Octothorpe,
    /// Left parenthesis '(' token
    LeftParenthesis,
    /// Right parenthesis ')' token
    RightParenthesis,
    /// Left curly brace '{' token
    LeftCurlyBrace,
    /// Right curly brace '}' token
    RightCurlyBrace,
    /// Left square bracket '[' token
    LeftSquareBracket,
    /// Right square bracket ']' token
    RightSquareBracket,
    /// Comma ',' token
    Comma,
    /// Semicolon ';' token
    Semicolon,
    /// At sign '@' token
    AtSign,
    /// Exclamation mark '!' token
    ExclamationMark,
    /// Equality operator '=' token
    EqualityOperator,
    /// Inequality operator '<>' token
    InequalityOperator,
    /// Less than or equal operator '<=' token
    LessThanOrEqualOperator,
    /// Greater than or equal operator '>=' token
    GreaterThanOrEqualOperator,
    /// Less than operator '<' token
    LessThanOperator,
    /// Greater than operator '>' token
    GreaterThanOperator,
    /// Multiplication operator '*' token
    MultiplicationOperator,
    /// Subtraction operator '-' token
    SubtractionOperator,
    /// Addition operator '+' token
    AdditionOperator,
    /// Division operator '/' token
    DivisionOperator,
    /// Backward slash operator '\' token
    BackwardSlashOperator,
    /// Period operator '.' token
    PeriodOperator,
    /// Colon operator ':' token
    ColonOperator,
    /// Exponentiation operator '^' token
    ExponentiationOperator,

    // Error recovery
    /// Error node for error recovery
    Error,

    /// Unknown syntax kind
    Unknown,
}

impl Display for SyntaxKind {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{self:?}")
    }
}

impl From<Token> for SyntaxKind {
    fn from(token: Token) -> Self {
        match token {
            Token::Whitespace => SyntaxKind::Whitespace,
            Token::Newline => SyntaxKind::Newline,
            Token::EndOfLineComment => SyntaxKind::EndOfLineComment,
            Token::RemComment => SyntaxKind::RemComment,
            Token::ClassKeyword => SyntaxKind::ClassKeyword,
            Token::ReDimKeyword => SyntaxKind::ReDimKeyword,
            Token::PreserveKeyword => SyntaxKind::PreserveKeyword,
            Token::DimKeyword => SyntaxKind::DimKeyword,
            Token::DeclareKeyword => SyntaxKind::DeclareKeyword,
            Token::AliasKeyword => SyntaxKind::AliasKeyword,
            Token::AttributeKeyword => SyntaxKind::AttributeKeyword,
            Token::BeginKeyword => SyntaxKind::BeginKeyword,
            Token::LibKeyword => SyntaxKind::LibKeyword,
            Token::WithKeyword => SyntaxKind::WithKeyword,
            Token::WithEventsKeyword => SyntaxKind::WithEventsKeyword,
            Token::BaseKeyword => SyntaxKind::BaseKeyword,
            Token::CompareKeyword => SyntaxKind::CompareKeyword,
            Token::OptionKeyword => SyntaxKind::OptionKeyword,
            Token::ExplicitKeyword => SyntaxKind::ExplicitKeyword,
            Token::PrivateKeyword => SyntaxKind::PrivateKeyword,
            Token::PublicKeyword => SyntaxKind::PublicKeyword,
            Token::ConstKeyword => SyntaxKind::ConstKeyword,
            Token::AsKeyword => SyntaxKind::AsKeyword,
            Token::ByValKeyword => SyntaxKind::ByValKeyword,
            Token::ByRefKeyword => SyntaxKind::ByRefKeyword,
            Token::OptionalKeyword => SyntaxKind::OptionalKeyword,
            Token::FunctionKeyword => SyntaxKind::FunctionKeyword,
            Token::StaticKeyword => SyntaxKind::StaticKeyword,
            Token::SubKeyword => SyntaxKind::SubKeyword,
            Token::EndKeyword => SyntaxKind::EndKeyword,
            Token::TrueKeyword => SyntaxKind::TrueKeyword,
            Token::FalseKeyword => SyntaxKind::FalseKeyword,
            Token::EnumKeyword => SyntaxKind::EnumKeyword,
            Token::TypeKeyword => SyntaxKind::TypeKeyword,
            Token::BooleanKeyword => SyntaxKind::BooleanKeyword,
            Token::DoubleKeyword => SyntaxKind::DoubleKeyword,
            Token::CurrencyKeyword => SyntaxKind::CurrencyKeyword,
            Token::DecimalKeyword => SyntaxKind::DecimalKeyword,
            Token::DateKeyword => SyntaxKind::DateKeyword,
            Token::ObjectKeyword => SyntaxKind::ObjectKeyword,
            Token::VariantKeyword => SyntaxKind::VariantKeyword,
            Token::ByteKeyword => SyntaxKind::ByteKeyword,
            Token::LongKeyword => SyntaxKind::LongKeyword,
            Token::SingleKeyword => SyntaxKind::SingleKeyword,
            Token::StringKeyword => SyntaxKind::StringKeyword,
            Token::IntegerKeyword => SyntaxKind::IntegerKeyword,
            Token::StringLiteral => SyntaxKind::StringLiteral,
            Token::IntegerLiteral => SyntaxKind::IntegerLiteral,
            Token::LongLiteral => SyntaxKind::LongLiteral,
            Token::SingleLiteral => SyntaxKind::SingleLiteral,
            Token::DoubleLiteral => SyntaxKind::DoubleLiteral,
            Token::DecimalLiteral => SyntaxKind::DecimalLiteral,
            Token::DateLiteral => SyntaxKind::DateLiteral,
            Token::IfKeyword => SyntaxKind::IfKeyword,
            Token::ElseKeyword => SyntaxKind::ElseKeyword,
            Token::ElseIfKeyword => SyntaxKind::ElseIfKeyword,
            Token::AndKeyword => SyntaxKind::AndKeyword,
            Token::OrKeyword => SyntaxKind::OrKeyword,
            Token::XorKeyword => SyntaxKind::XorKeyword,
            Token::ModKeyword => SyntaxKind::ModKeyword,
            Token::EqvKeyword => SyntaxKind::EqvKeyword,
            Token::AddressOfKeyword => SyntaxKind::AddressOfKeyword,
            Token::ImpKeyword => SyntaxKind::ImpKeyword,
            Token::IsKeyword => SyntaxKind::IsKeyword,
            Token::LikeKeyword => SyntaxKind::LikeKeyword,
            Token::NotKeyword => SyntaxKind::NotKeyword,
            Token::ThenKeyword => SyntaxKind::ThenKeyword,
            Token::GotoKeyword => SyntaxKind::GotoKeyword,
            Token::GoSubKeyword => SyntaxKind::GoSubKeyword,
            Token::ReturnKeyword => SyntaxKind::ReturnKeyword,
            Token::ExitKeyword => SyntaxKind::ExitKeyword,
            Token::ForKeyword => SyntaxKind::ForKeyword,
            Token::EachKeyword => SyntaxKind::EachKeyword,
            Token::InKeyword => SyntaxKind::InKeyword,
            Token::ToKeyword => SyntaxKind::ToKeyword,
            Token::LockKeyword => SyntaxKind::LockKeyword,
            Token::UnlockKeyword => SyntaxKind::UnlockKeyword,
            Token::StepKeyword => SyntaxKind::StepKeyword,
            Token::StopKeyword => SyntaxKind::StopKeyword,
            Token::WhileKeyword => SyntaxKind::WhileKeyword,
            Token::WendKeyword => SyntaxKind::WendKeyword,
            Token::WidthKeyword => SyntaxKind::WidthKeyword,
            Token::WriteKeyword => SyntaxKind::WriteKeyword,
            Token::TimeKeyword => SyntaxKind::TimeKeyword,
            Token::SetAttrKeyword => SyntaxKind::SetAttrKeyword,
            Token::SetKeyword => SyntaxKind::SetKeyword,
            Token::SendKeysKeyword => SyntaxKind::SendKeysKeyword,
            Token::SelectKeyword => SyntaxKind::SelectKeyword,
            Token::CaseKeyword => SyntaxKind::CaseKeyword,
            Token::SeekKeyword => SyntaxKind::SeekKeyword,
            Token::SaveSettingKeyword => SyntaxKind::SaveSettingKeyword,
            Token::SavePictureKeyword => SyntaxKind::SavePictureKeyword,
            Token::RSetKeyword => SyntaxKind::RSetKeyword,
            Token::RmDirKeyword => SyntaxKind::RmDirKeyword,
            Token::ResumeKeyword => SyntaxKind::ResumeKeyword,
            Token::ResetKeyword => SyntaxKind::ResetKeyword,
            Token::RandomizeKeyword => SyntaxKind::RandomizeKeyword,
            Token::RaiseEventKeyword => SyntaxKind::RaiseEventKeyword,
            Token::PutKeyword => SyntaxKind::PutKeyword,
            Token::PropertyKeyword => SyntaxKind::PropertyKeyword,
            Token::PrintKeyword => SyntaxKind::PrintKeyword,
            Token::OpenKeyword => SyntaxKind::OpenKeyword,
            Token::OnKeyword => SyntaxKind::OnKeyword,
            Token::OffKeyword => SyntaxKind::OffKeyword,
            Token::NameKeyword => SyntaxKind::NameKeyword,
            Token::MkDirKeyword => SyntaxKind::MkDirKeyword,
            Token::MidBKeyword => SyntaxKind::MidBKeyword,
            Token::MidKeyword => SyntaxKind::MidKeyword,
            Token::LSetKeyword => SyntaxKind::LSetKeyword,
            Token::LoadKeyword => SyntaxKind::LoadKeyword,
            Token::UnloadKeyword => SyntaxKind::UnloadKeyword,
            Token::LineKeyword => SyntaxKind::LineKeyword,
            Token::InputKeyword => SyntaxKind::InputKeyword,
            Token::LetKeyword => SyntaxKind::LetKeyword,
            Token::KillKeyword => SyntaxKind::KillKeyword,
            Token::ImplementsKeyword => SyntaxKind::ImplementsKeyword,
            Token::GetKeyword => SyntaxKind::GetKeyword,
            Token::FileCopyKeyword => SyntaxKind::FileCopyKeyword,
            Token::EventKeyword => SyntaxKind::EventKeyword,
            Token::ErrorKeyword => SyntaxKind::ErrorKeyword,
            Token::EraseKeyword => SyntaxKind::EraseKeyword,
            Token::DoKeyword => SyntaxKind::DoKeyword,
            Token::UntilKeyword => SyntaxKind::UntilKeyword,
            Token::LoopKeyword => SyntaxKind::LoopKeyword,
            Token::DeleteSettingKeyword => SyntaxKind::DeleteSettingKeyword,
            Token::DefBoolKeyword => SyntaxKind::DefBoolKeyword,
            Token::DefByteKeyword => SyntaxKind::DefByteKeyword,
            Token::DefIntKeyword => SyntaxKind::DefIntKeyword,
            Token::DefLngKeyword => SyntaxKind::DefLngKeyword,
            Token::DefCurKeyword => SyntaxKind::DefCurKeyword,
            Token::DefSngKeyword => SyntaxKind::DefSngKeyword,
            Token::DefDblKeyword => SyntaxKind::DefDblKeyword,
            Token::DefDecKeyword => SyntaxKind::DefDecKeyword,
            Token::DefDateKeyword => SyntaxKind::DefDateKeyword,
            Token::DefStrKeyword => SyntaxKind::DefStrKeyword,
            Token::DefObjKeyword => SyntaxKind::DefObjKeyword,
            Token::DefVarKeyword => SyntaxKind::DefVarKeyword,
            Token::CloseKeyword => SyntaxKind::CloseKeyword,
            Token::ChDriveKeyword => SyntaxKind::ChDriveKeyword,
            Token::ChDirKeyword => SyntaxKind::ChDirKeyword,
            Token::CallKeyword => SyntaxKind::CallKeyword,
            Token::BeepKeyword => SyntaxKind::BeepKeyword,
            Token::AppActivateKeyword => SyntaxKind::AppActivateKeyword,
            Token::FriendKeyword => SyntaxKind::FriendKeyword,
            Token::BinaryKeyword => SyntaxKind::BinaryKeyword,
            Token::RandomKeyword => SyntaxKind::RandomKeyword,
            Token::ReadKeyword => SyntaxKind::ReadKeyword,
            Token::OutputKeyword => SyntaxKind::OutputKeyword,
            Token::AppendKeyword => SyntaxKind::AppendKeyword,
            Token::AccessKeyword => SyntaxKind::AccessKeyword,
            Token::TextKeyword => SyntaxKind::TextKeyword,
            Token::DatabaseKeyword => SyntaxKind::DatabaseKeyword,
            Token::EmptyKeyword => SyntaxKind::EmptyKeyword,
            Token::ModuleKeyword => SyntaxKind::ModuleKeyword,
            Token::NextKeyword => SyntaxKind::NextKeyword,
            Token::NewKeyword => SyntaxKind::NewKeyword,
            Token::LenKeyword => SyntaxKind::LenKeyword,
            Token::MeKeyword => SyntaxKind::MeKeyword,
            Token::NullKeyword => SyntaxKind::NullKeyword,
            Token::ParamArrayKeyword => SyntaxKind::ParamArrayKeyword,
            Token::DollarSign => SyntaxKind::DollarSign,
            Token::Underscore => SyntaxKind::Underscore,
            Token::Ampersand => SyntaxKind::Ampersand,
            Token::Percent => SyntaxKind::Percent,
            Token::Octothorpe => SyntaxKind::Octothorpe,
            Token::LeftParenthesis => SyntaxKind::LeftParenthesis,
            Token::RightParenthesis => SyntaxKind::RightParenthesis,
            Token::LeftCurlyBrace => SyntaxKind::LeftCurlyBrace,
            Token::RightCurlyBrace => SyntaxKind::RightCurlyBrace,
            Token::LeftSquareBracket => SyntaxKind::LeftSquareBracket,
            Token::RightSquareBracket => SyntaxKind::RightSquareBracket,
            Token::Comma => SyntaxKind::Comma,
            Token::Semicolon => SyntaxKind::Semicolon,
            Token::AtSign => SyntaxKind::AtSign,
            Token::ExclamationMark => SyntaxKind::ExclamationMark,
            Token::VersionKeyword => SyntaxKind::VersionKeyword,
            Token::EqualityOperator => SyntaxKind::EqualityOperator,
            Token::InequalityOperator => SyntaxKind::InequalityOperator,
            Token::LessThanOrEqualOperator => SyntaxKind::LessThanOrEqualOperator,
            Token::GreaterThanOrEqualOperator => SyntaxKind::GreaterThanOrEqualOperator,
            Token::LessThanOperator => SyntaxKind::LessThanOperator,
            Token::GreaterThanOperator => SyntaxKind::GreaterThanOperator,
            Token::MultiplicationOperator => SyntaxKind::MultiplicationOperator,
            Token::SubtractionOperator => SyntaxKind::SubtractionOperator,
            Token::AdditionOperator => SyntaxKind::AdditionOperator,
            Token::DivisionOperator => SyntaxKind::DivisionOperator,
            Token::BackwardSlashOperator => SyntaxKind::BackwardSlashOperator,
            Token::PeriodOperator => SyntaxKind::PeriodOperator,
            Token::ColonOperator => SyntaxKind::ColonOperator,
            Token::ExponentiationOperator => SyntaxKind::ExponentiationOperator,
            Token::Identifier => SyntaxKind::Identifier,
        }
    }
}

impl SyntaxKind {
    pub(crate) fn from_raw(raw: rowan::SyntaxKind) -> Self {
        assert!(raw.0 <= SyntaxKind::Unknown as u16);
        unsafe { std::mem::transmute::<u16, SyntaxKind>(raw.0) }
    }

    /// Convert `SyntaxKind` to rowan's raw `SyntaxKind` (for internal use in builders)
    pub(crate) fn to_raw(self) -> rowan::SyntaxKind {
        rowan::SyntaxKind(self as u16)
    }
}
