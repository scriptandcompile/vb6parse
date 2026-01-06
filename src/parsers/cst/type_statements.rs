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

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
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
    /// - User-defined types are passed by value unless explicitly passed `ByRef`
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

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume optional Public/Private keyword
        if self.at_token(Token::PublicKeyword) || self.at_token(Token::PrivateKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Type" keyword
        self.consume_token();

        // Consume any whitespace after "Type"
        self.consume_whitespace();

        // Consume type name (keywords can be used as type names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        // Parse type members until "End Type"
        while !self.is_at_end() {
            // Check if we've reached "End Type"
            if self.at_token(Token::EndKeyword)
                && self.peek_next_keyword() == Some(Token::TypeKeyword)
            {
                break;
            }

            // Consume type member lines (elementname [(subscripts)] As type)
            // This includes whitespace, comments, identifiers, operators, and newlines
            match self.current_token() {
                Some(Token::Whitespace
                | Token::Newline
                | Token::EndOfLineComment
                | Token::RemComment
                | Token::Identifier
                | Token::AsKeyword
                | Token::LeftParenthesis
                | Token::RightParenthesis
                | Token::ToKeyword
                | Token::IntegerLiteral
                | Token::LongLiteral
                | Token::Comma
                | Token::MultiplicationOperator // For String * length
                | Token::SubtractionOperator   // For negative array bounds
                // Data type keywords that can appear in Type members
                | Token::ByteKeyword
                | Token::BooleanKeyword
                | Token::IntegerKeyword
                | Token::LongKeyword
                | Token::CurrencyKeyword
                | Token::SingleKeyword
                | Token::DoubleKeyword
                | Token::DateKeyword
                | Token::StringKeyword
                | Token::ObjectKeyword
                | Token::VariantKeyword) => {
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
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Type"
            self.consume_whitespace();

            // Consume "Type"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // TypeStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn type_simple() {
        let source = r"
Type Point
    x As Single
    y As Single
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Point"),
                Newline,
                Whitespace,
                Identifier ("x"),
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
                Whitespace,
                Identifier ("y"),
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_multiple_fields() {
        let source = r"
Type Employee
    EmployeeID As Long
    FirstName As String
    LastName As String
    HireDate As Date
    Salary As Currency
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Employee"),
                Newline,
                Whitespace,
                Identifier ("EmployeeID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("FirstName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("LastName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("HireDate"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DateKeyword,
                Newline,
                Whitespace,
                Identifier ("Salary"),
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_public() {
        let source = r"
Public Type Rectangle
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                PublicKeyword,
                Whitespace,
                TypeKeyword,
                Whitespace,
                Identifier ("Rectangle"),
                Newline,
                Whitespace,
                Identifier ("Left"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Top"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Right"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Bottom"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_private() {
        let source = r"
Private Type InternalData
    Buffer As String
    Length As Integer
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                PrivateKeyword,
                Whitespace,
                TypeKeyword,
                Whitespace,
                Identifier ("InternalData"),
                Newline,
                Whitespace,
                Identifier ("Buffer"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("Length"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_fixed_length_string() {
        let source = r"
Type CustomerRecord
    CustomerID As Long
    CustomerName As String * 50
    Address As String * 100
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("CustomerRecord"),
                Newline,
                Whitespace,
                Identifier ("CustomerID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("CustomerName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("50"),
                Newline,
                Whitespace,
                Identifier ("Address"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("100"),
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_array_element() {
        let source = r"
Type SalesData
    SalesPersonID As Long
    MonthlySales(1 To 12) As Currency
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("SalesData"),
                Newline,
                Whitespace,
                Identifier ("SalesPersonID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("MonthlySales"),
                LeftParenthesis,
                IntegerLiteral ("1"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("12"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_multiple_array_dimensions() {
        let source = r"
Type Matrix
    Data(1 To 10, 1 To 10) As Double
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Matrix"),
                Newline,
                Whitespace,
                Identifier ("Data"),
                LeftParenthesis,
                IntegerLiteral ("1"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("10"),
                Comma,
                Whitespace,
                IntegerLiteral ("1"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("10"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_dynamic_array() {
        let source = r"
Type DynamicBuffer
    Items() As Variant
    Count As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("DynamicBuffer"),
                Newline,
                Whitespace,
                Identifier ("Items"),
                LeftParenthesis,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                Whitespace,
                Identifier ("Count"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_all_data_types() {
        let source = r"
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
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("AllTypes"),
                Newline,
                Whitespace,
                Identifier ("ByteField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                ByteKeyword,
                Newline,
                Whitespace,
                Identifier ("BoolField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                Whitespace,
                Identifier ("IntField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                Whitespace,
                Identifier ("LongField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("CurrField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
                Whitespace,
                Identifier ("SingleField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
                Whitespace,
                Identifier ("DoubleField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                Whitespace,
                Identifier ("DateField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DateKeyword,
                Newline,
                Whitespace,
                Identifier ("StringField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("ObjectField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                ObjectKeyword,
                Newline,
                Whitespace,
                Identifier ("VariantField"),
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_nested_type_reference() {
        let source = r"
Type Address
    Street As String
    City As String
End Type

Type Person
    Name As String
    HomeAddress As Address
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Address"),
                Newline,
                Whitespace,
                Identifier ("Street"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("City"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Person"),
                Newline,
                Whitespace,
                NameKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("HomeAddress"),
                Whitespace,
                AsKeyword,
                Whitespace,
                Identifier ("Address"),
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_comments() {
        let source = r"
Type Employee
    ' Employee identification
    EmployeeID As Long
    FirstName As String  ' First name
    LastName As String   ' Last name
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Employee"),
                Newline,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("EmployeeID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("FirstName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("LastName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                EndOfLineComment,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_empty() {
        let source = r"
Type EmptyType
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("EmptyType"),
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_single_field() {
        let source = r"
Type SimpleType
    Value As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("SimpleType"),
                Newline,
                Whitespace,
                Identifier ("Value"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_array_bounds() {
        let source = r"
Type BoundedArray
    Items(0 To 99) As Integer
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("BoundedArray"),
                Newline,
                Whitespace,
                Identifier ("Items"),
                LeftParenthesis,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("99"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_negative_array_bounds() {
        let source = r"
Type NegativeBounds
    Values(-10 To 10) As Single
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("NegativeBounds"),
                Newline,
                Whitespace,
                Identifier ("Values"),
                LeftParenthesis,
                SubtractionOperator,
                IntegerLiteral ("10"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("10"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_api_structure() {
        let source = r"
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("RECT"),
                Newline,
                Whitespace,
                Identifier ("Left"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Top"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Right"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("Bottom"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_multiple_declarations() {
        let source = r"
Type Type1
    Field1 As Integer
End Type

Type Type2
    Field2 As String
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Type1"),
                Newline,
                Whitespace,
                Identifier ("Field1"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Type2"),
                Newline,
                Whitespace,
                Identifier ("Field2"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_object_reference() {
        let source = r"
Type DataContainer
    RecordSet As Object
    Connection As Object
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("DataContainer"),
                Newline,
                Whitespace,
                Identifier ("RecordSet"),
                Whitespace,
                AsKeyword,
                Whitespace,
                ObjectKeyword,
                Newline,
                Whitespace,
                Identifier ("Connection"),
                Whitespace,
                AsKeyword,
                Whitespace,
                ObjectKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_variant_field() {
        let source = r"
Type FlexibleData
    DataType As Integer
    DataValue As Variant
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("FlexibleData"),
                Newline,
                Whitespace,
                Identifier ("DataType"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                Whitespace,
                Identifier ("DataValue"),
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_complex_array_subscripts() {
        let source = r"
Type ComplexArrays
    Matrix2D(1 To 5, 1 To 5) As Double
    Matrix3D(0 To 2, 0 To 2, 0 To 2) As Single
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("ComplexArrays"),
                Newline,
                Whitespace,
                Identifier ("Matrix2D"),
                LeftParenthesis,
                IntegerLiteral ("1"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("5"),
                Comma,
                Whitespace,
                IntegerLiteral ("1"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("5"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                Whitespace,
                Identifier ("Matrix3D"),
                LeftParenthesis,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("2"),
                Comma,
                Whitespace,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("2"),
                Comma,
                Whitespace,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("2"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                SingleKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_keyword_as_field_name() {
        let source = r"
Type KeywordFields
    Name As String
    Type As Integer
    End As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("KeywordFields"),
                Newline,
                Whitespace,
                NameKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Unknown,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                Whitespace,
                Unknown,
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_long_fixed_string() {
        let source = r"
Type FileRecord
    FileName As String * 255
    FileSize As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("FileRecord"),
                Newline,
                Whitespace,
                Identifier ("FileName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("255"),
                Newline,
                Whitespace,
                Identifier ("FileSize"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_multiple_fixed_strings() {
        let source = r"
Type ContactInfo
    FirstName As String * 30
    LastName As String * 30
    Phone As String * 15
    Email As String * 50
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("ContactInfo"),
                Newline,
                Whitespace,
                Identifier ("FirstName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("30"),
                Newline,
                Whitespace,
                Identifier ("LastName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("30"),
                Newline,
                Whitespace,
                Identifier ("Phone"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("15"),
                Newline,
                Whitespace,
                Identifier ("Email"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("50"),
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_inline_comments() {
        let source = r"
Type Config
    MaxConnections As Integer  ' Maximum allowed connections
    TimeoutSeconds As Long     ' Timeout in seconds
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Config"),
                Newline,
                Whitespace,
                Identifier ("MaxConnections"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("TimeoutSeconds"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Whitespace,
                EndOfLineComment,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_rem_comments() {
        let source = r"
Type Data
    Rem This is the ID field
    ID As Long
    Rem This is the name field
    Name As String
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("Data"),
                Newline,
                Whitespace,
                RemComment,
                Newline,
                Whitespace,
                Identifier ("ID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                RemComment,
                Newline,
                Whitespace,
                NameKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_uppercase_name() {
        let source = r"
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("PROCESSENTRY32"),
                Newline,
                Whitespace,
                Identifier ("dwSize"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("cntUsage"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("th32ProcessID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_mixed_case_fields() {
        let source = r"
Type MixedCase
    firstName As String
    LastName As String
    EMPLOYEE_ID As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("MixedCase"),
                Newline,
                Whitespace,
                Identifier ("firstName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("LastName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("EMPLOYEE_ID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_byte_array() {
        let source = r"
Type BinaryData
    Buffer(0 To 255) As Byte
    Length As Integer
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("BinaryData"),
                Newline,
                Whitespace,
                Identifier ("Buffer"),
                LeftParenthesis,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("255"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                ByteKeyword,
                Newline,
                Whitespace,
                Identifier ("Length"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_comprehensive_example() {
        let source = r"
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
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                PublicKeyword,
                Whitespace,
                TypeKeyword,
                Whitespace,
                Identifier ("CustomerRecord"),
                Newline,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("CustomerID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                Whitespace,
                Identifier ("AccountNumber"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("20"),
                Newline,
                Whitespace,
                Newline,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("FirstName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("50"),
                Newline,
                Whitespace,
                Identifier ("LastName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("50"),
                Newline,
                Whitespace,
                Identifier ("Email"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Whitespace,
                MultiplicationOperator,
                Whitespace,
                IntegerLiteral ("100"),
                Newline,
                Whitespace,
                Newline,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("Balance"),
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
                Whitespace,
                Identifier ("CreditLimit"),
                Whitespace,
                AsKeyword,
                Whitespace,
                CurrencyKeyword,
                Newline,
                Whitespace,
                Identifier ("TransactionHistory"),
                LeftParenthesis,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                VariantKeyword,
                Newline,
                Whitespace,
                Newline,
                Whitespace,
                EndOfLineComment,
                Newline,
                Whitespace,
                Identifier ("CreatedDate"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DateKeyword,
                Newline,
                Whitespace,
                Identifier ("LastModified"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DateKeyword,
                Newline,
                Whitespace,
                Identifier ("IsActive"),
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn type_with_zero_based_array() {
        let source = r"
Type ZeroBasedData
    Items(0 To 9) As Integer
    Count As Long
End Type
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("ZeroBasedData"),
                Newline,
                Whitespace,
                Identifier ("Items"),
                LeftParenthesis,
                IntegerLiteral ("0"),
                Whitespace,
                ToKeyword,
                Whitespace,
                IntegerLiteral ("9"),
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                Whitespace,
                Identifier ("Count"),
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
        ]);
    }
}
