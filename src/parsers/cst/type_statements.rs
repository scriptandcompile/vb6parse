//! Type statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Type (user-defined type) statements.
//!
//! Type statement syntax:
//!
//! \[Public | Private\] Type typename
//! elementname \[(subscripts)\] As type
//! \[elementname \[(subscripts)\] As type\]
//! ...
//! End Type
//!
//! User-defined types (UDTs) provide a way to create custom data structures
//! that group related variables of different data types under one name.
//!
//! [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-statement)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Visual Basic 6 Type statement with syntax:
    ///
    /// \[Public | Private\] Type typename
    /// elementname \[(subscripts)\] As type
    /// \[elementname \[(subscripts)\] As type\]
    /// ...
    /// End Type
    ///
    /// The Type statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the Type is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the Type is not available outside the project. |
    /// | Private     | Optional | Indicates that the Type is accessible only to other procedures in the module where it is declared. |
    /// | typename    | Required | Name of the user-defined type; follows standard variable naming conventions. |
    /// | elementname | Required | Name of an element (field) of the user-defined type. Element names follow standard variable naming conventions, except that keywords can be used. |
    /// | subscripts  | Optional | Dimensions of an array element. Use only parentheses when declaring an array whose size can change. The subscript syntax has these parts: \[lower To\] upper \[, \[lower To\] upper\] ... |
    /// | type        | Required | Data type of the element; may be Byte, Boolean, Integer, Long, Currency, Single, Double, Decimal (not currently supported), Date, String (variable-length or fixed-length), Object, Variant, another user-defined type, or an object type. |
    ///
    /// Remarks:
    /// - User-defined types are typically used to create records similar to those in databases.
    /// - User-defined types can contain elements of different data types.
    /// - Array elements within user-defined types can be dynamic arrays (using empty parentheses).
    /// - Fixed-length strings can be used in user-defined types.
    /// - Type statements can only be used at the module level. Once you declare a user-defined type using the Type statement, you can declare a variable of that type anywhere within the scope of the declaration.
    /// - User-defined types are useful for passing multiple related values as a single unit to procedures.
    /// - Cannot be used in class modules unless they are Private.
    ///
    /// ## Examples
    ///
    /// ### Basic User-Defined Type
    ///
    /// ```vb
    /// Type Employee
    ///     EmployeeID As Long
    ///     FirstName As String
    ///     LastName As String
    ///     HireDate As Date
    ///     Salary As Currency
    /// End Type
    /// ```
    ///
    /// ### Type with Fixed-Length String
    ///
    /// ```vb
    /// Type CustomerRecord
    ///     CustomerID As Long
    ///     CustomerName As String * 50
    ///     Address As String * 100
    ///     City As String * 30
    ///     ZipCode As String * 10
    /// End Type
    /// ```
    ///
    /// ### Type with Array Element
    ///
    /// ```vb
    /// Type SalesData
    ///     SalesPersonID As Long
    ///     MonthlySales(1 To 12) As Currency
    ///     QuarterlySales(1 To 4) As Currency
    /// End Type
    /// ```
    ///
    /// ### Nested User-Defined Types
    ///
    /// ```vb
    /// Type Address
    ///     Street As String
    ///     City As String
    ///     State As String
    ///     ZipCode As String
    /// End Type
    ///
    /// Type Person
    ///     Name As String
    ///     HomeAddress As Address
    ///     WorkAddress As Address
    /// End Type
    /// ```
    ///
    /// ### Public Type Declaration
    ///
    /// ```vb
    /// Public Type Point
    ///     x As Single
    ///     y As Single
    /// End Type
    /// ```
    ///
    /// ### Private Type Declaration
    ///
    /// ```vb
    /// Private Type InternalData
    ///     Buffer(0 To 255) As Byte
    ///     Length As Integer
    /// End Type
    /// ```
    ///
    /// ### Type with Variant Element
    ///
    /// ```vb
    /// Type FlexibleRecord
    ///     RecordType As Integer
    ///     Data As Variant
    /// End Type
    /// ```
    ///
    /// ### Type for API Structures
    ///
    /// ```vb
    /// Type RECT
    ///     Left As Long
    ///     Top As Long
    ///     Right As Long
    ///     Bottom As Long
    /// End Type
    /// ```
    ///
    /// ## Common Patterns
    ///
    /// ### Using Type in Declarations
    ///
    /// ```vb
    /// Dim emp As Employee
    /// emp.EmployeeID = 1001
    /// emp.FirstName = "John"
    /// emp.LastName = "Doe"
    /// ```
    ///
    /// ### Passing Type to Procedures
    ///
    /// ```vb
    /// Sub UpdateEmployee(empData As Employee)
    ///     ' Update database with employee data
    ///     Debug.Print empData.FirstName & " " & empData.LastName
    /// End Sub
    /// ```
    ///
    /// ### Arrays of User-Defined Types
    ///
    /// ```vb
    /// Dim employees(1 To 100) As Employee
    /// employees(1).EmployeeID = 1001
    /// employees(1).FirstName = "John"
    /// ```
    ///
    /// ## Best Practices
    ///
    /// 1. Use meaningful names for Type and element names
    /// 2. Use fixed-length strings when the length is known and constant
    /// 3. Group related data into a single Type
    /// 4. Use Public for Types that need to be shared across modules
    /// 5. Use Private for Types that are module-specific
    /// 6. Document complex Types with comments
    /// 7. Consider performance implications of large Types
    ///
    /// ## Important Notes
    ///
    /// - Type statements cannot be nested within procedures
    /// - Type statements must appear at module level
    /// - In class modules, Type must be Private
    /// - Elements of a Type can be other user-defined types
    /// - User-defined types are passed by value unless explicitly passed ByRef
    /// - Type members are accessed using the dot (.) operator
    ///
    /// ## See Also
    ///
    /// - `Dim` statement (declaring variables of user-defined types)
    /// - `Public` statement (module-level public declarations)
    /// - `Private` statement (module-level private declarations)
    /// - Fixed-length strings (`String * length`)
    ///
    /// ## References
    ///
    /// - [Microsoft Docs: Type Statement](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/type-statement)
    /// - [User-Defined Types](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/creating-your-own-data-types)
    pub(super) fn parse_type_statement(&mut self) {
        // if we are now parsing a type statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::TypeStatement.to_raw());

        // Consume optional Public/Private keyword
        if self.at_token(VB6Token::PublicKeyword) || self.at_token(VB6Token::PrivateKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Type" keyword
        self.consume_token();

        // Consume any whitespace after "Type"
        self.consume_whitespace();

        // Consume type name (keywords can be used as type names in VB6)
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(VB6Token::Newline);

        // Parse type members until "End Type"
        while !self.is_at_end() {
            // Check if we've reached "End Type"
            if self.at_token(VB6Token::EndKeyword)
                && self.peek_next_keyword() == Some(VB6Token::TypeKeyword)
            {
                break;
            }

            // Consume type member lines (elementname [(subscripts)] As type)
            // This includes whitespace, comments, identifiers, operators, and newlines
            match self.current_token() {
                Some(VB6Token::Whitespace
                | VB6Token::Newline
                | VB6Token::EndOfLineComment
                | VB6Token::RemComment
                | VB6Token::Identifier
                | VB6Token::AsKeyword
                | VB6Token::LeftParenthesis
                | VB6Token::RightParenthesis
                | VB6Token::ToKeyword
                | VB6Token::IntegerLiteral
                | VB6Token::LongLiteral
                | VB6Token::Comma
                | VB6Token::MultiplicationOperator // For String * length
                | VB6Token::SubtractionOperator   // For negative array bounds
                // Data type keywords that can appear in Type members
                | VB6Token::ByteKeyword
                | VB6Token::BooleanKeyword
                | VB6Token::IntegerKeyword
                | VB6Token::LongKeyword
                | VB6Token::CurrencyKeyword
                | VB6Token::SingleKeyword
                | VB6Token::DoubleKeyword
                | VB6Token::DateKeyword
                | VB6Token::StringKeyword
                | VB6Token::ObjectKeyword
                | VB6Token::VariantKeyword) => {
                    self.consume_token();
                }
                _ => {
                    // Check if this is a keyword being used as an identifier (VB6 allows this)
                    if self.at_keyword() {
                        self.consume_token();
                    } else {
                        // Unknown token in type body, consume it
                        self.consume_token_as_unknown();
                    }
                }
            }
        }

        // Consume "End Type" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Type"
            self.consume_whitespace();

            // Consume "Type"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(VB6Token::Newline);
        }

        self.builder.finish_node(); // TypeStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn type_simple() {
        let source = r#"
Type Point
    x As Single
    y As Single
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("TypeKeyword"));
        assert!(debug.contains("EndKeyword"));
    }

    #[test]
    fn type_with_multiple_fields() {
        let source = r#"
Type Employee
    EmployeeID As Long
    FirstName As String
    LastName As String
    HireDate As Date
    Salary As Currency
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("EmployeeID"));
        assert!(debug.contains("FirstName"));
        assert!(debug.contains("Salary"));
    }

    #[test]
    fn type_public() {
        let source = r#"
Public Type Rectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("PublicKeyword"));
        assert!(debug.contains("Rectangle"));
    }

    #[test]
    fn type_private() {
        let source = r#"
Private Type InternalData
    Buffer As String
    Length As Integer
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("PrivateKeyword"));
        assert!(debug.contains("InternalData"));
    }

    #[test]
    fn type_with_fixed_length_string() {
        let source = r#"
Type CustomerRecord
    CustomerID As Long
    CustomerName As String * 50
    Address As String * 100
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("CustomerName"));
        assert!(debug.contains("String"));
        assert!(debug.contains("50"));
    }

    #[test]
    fn type_with_array_element() {
        let source = r#"
Type SalesData
    SalesPersonID As Long
    MonthlySales(1 To 12) As Currency
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("MonthlySales"));
        assert!(debug.contains("1"));
        assert!(debug.contains("12"));
    }

    #[test]
    fn type_with_multiple_array_dimensions() {
        let source = r#"
Type Matrix
    Data(1 To 10, 1 To 10) As Double
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Data"));
        assert!(debug.contains("10"));
    }

    #[test]
    fn type_with_dynamic_array() {
        let source = r#"
Type DynamicBuffer
    Items() As Variant
    Count As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Items"));
        assert!(debug.contains("Count"));
    }

    #[test]
    fn type_with_all_data_types() {
        let source = r#"
Type AllTypes
    ByteField As Byte
    BoolField As Boolean
    IntField As Integer
    LongField As Long
    CurrField As Currency
    SingleField As Single
    DoubleField As Double
    DateField As Date
    StringField As String
    ObjectField As Object
    VariantField As Variant
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("ByteField"));
        assert!(debug.contains("VariantField"));
    }

    #[test]
    fn type_with_nested_type_reference() {
        let source = r#"
Type Address
    Street As String
    City As String
End Type

Type Person
    Name As String
    HomeAddress As Address
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Address"));
        assert!(debug.contains("Person"));
        assert!(debug.contains("HomeAddress"));
    }

    #[test]
    fn type_with_comments() {
        let source = r#"
Type Employee
    ' Employee identification
    EmployeeID As Long
    FirstName As String  ' First name
    LastName As String   ' Last name
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("EmployeeID"));
    }

    #[test]
    fn type_empty() {
        let source = r#"
Type EmptyType
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("EmptyType"));
    }

    #[test]
    fn type_single_field() {
        let source = r#"
Type SimpleType
    Value As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Value"));
    }

    #[test]
    fn type_with_array_bounds() {
        let source = r#"
Type BoundedArray
    Items(0 To 99) As Integer
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("0"));
        assert!(debug.contains("99"));
    }

    #[test]
    fn type_with_negative_array_bounds() {
        let source = r#"
Type NegativeBounds
    Values(-10 To 10) As Single
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("10"));
    }

    #[test]
    fn type_api_structure() {
        let source = r#"
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("RECT"));
        assert!(debug.contains("Left"));
        assert!(debug.contains("Bottom"));
    }

    #[test]
    fn type_multiple_declarations() {
        let source = r#"
Type Type1
    Field1 As Integer
End Type

Type Type2
    Field2 As String
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // CST includes whitespace, so count is more than 2
        assert!(cst.child_count() >= 2);
        assert!(debug.contains("Type1"));
        assert!(debug.contains("Type2"));
    }

    #[test]
    fn type_with_object_reference() {
        let source = r#"
Type DataContainer
    RecordSet As Object
    Connection As Object
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("RecordSet"));
        assert!(debug.contains("Connection"));
    }

    #[test]
    fn type_with_variant_field() {
        let source = r#"
Type FlexibleData
    DataType As Integer
    DataValue As Variant
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("DataValue"));
        assert!(debug.contains("Variant"));
    }

    #[test]
    fn type_complex_array_subscripts() {
        let source = r#"
Type ComplexArrays
    Matrix2D(1 To 5, 1 To 5) As Double
    Matrix3D(0 To 2, 0 To 2, 0 To 2) As Single
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Matrix2D"));
        assert!(debug.contains("Matrix3D"));
    }

    #[test]
    fn type_keyword_as_field_name() {
        let source = r#"
Type KeywordFields
    Name As String
    Type As Integer
    End As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Name"));
    }

    #[test]
    fn type_long_fixed_string() {
        let source = r#"
Type FileRecord
    FileName As String * 255
    FileSize As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("255"));
    }

    #[test]
    fn type_multiple_fixed_strings() {
        let source = r#"
Type ContactInfo
    FirstName As String * 30
    LastName As String * 30
    Phone As String * 15
    Email As String * 50
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("FirstName"));
        assert!(debug.contains("Email"));
    }

    #[test]
    fn type_with_inline_comments() {
        let source = r#"
Type Config
    MaxConnections As Integer  ' Maximum allowed connections
    TimeoutSeconds As Long     ' Timeout in seconds
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("MaxConnections"));
        assert!(debug.contains("TimeoutSeconds"));
    }

    #[test]
    fn type_with_rem_comments() {
        let source = r#"
Type Data
    Rem This is the ID field
    ID As Long
    Rem This is the name field
    Name As String
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("ID"));
        assert!(debug.contains("Name"));
    }

    #[test]
    fn type_uppercase_name() {
        let source = r#"
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("PROCESSENTRY32"));
    }

    #[test]
    fn type_mixed_case_fields() {
        let source = r#"
Type MixedCase
    firstName As String
    LastName As String
    EMPLOYEE_ID As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("firstName"));
        assert!(debug.contains("LastName"));
    }

    #[test]
    fn type_with_byte_array() {
        let source = r#"
Type BinaryData
    Buffer(0 To 255) As Byte
    Length As Integer
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Buffer"));
        assert!(debug.contains("Byte"));
    }

    #[test]
    fn type_comprehensive_example() {
        let source = r#"
Public Type CustomerRecord
    ' Identification
    CustomerID As Long
    AccountNumber As String * 20
    
    ' Personal Info
    FirstName As String * 50
    LastName As String * 50
    Email As String * 100
    
    ' Financial Data
    Balance As Currency
    CreditLimit As Currency
    TransactionHistory() As Variant
    
    ' Metadata
    CreatedDate As Date
    LastModified As Date
    IsActive As Boolean
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("PublicKeyword"));
        assert!(debug.contains("CustomerID"));
        assert!(debug.contains("Balance"));
        assert!(debug.contains("IsActive"));
    }

    #[test]
    fn type_with_zero_based_array() {
        let source = r#"
Type ZeroBasedData
    Items(0 To 9) As Integer
    Count As Long
End Type
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TypeStatement"));
        assert!(debug.contains("Items"));
        assert!(debug.contains("0"));
        assert!(debug.contains("9"));
        assert!(debug.contains("Count"));
    }
}
