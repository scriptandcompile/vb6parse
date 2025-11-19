//! Declare statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Declare statements (external function/sub declarations).
//!
//! Declare statement syntax:
//! \[ Public | Private \] Declare { Sub | Function } name Lib "libname" \[ Alias "aliasname" \] \[ ( arglist ) \] \[ As type \]
//!
//! Sub statements are handled in the sub_statements module.
//! Function statements are handled in the function_statements module.
//! Dim/ReDim and general Variable declarations are handled in the array_statements module.
//! Property statements are handled in the property_statements module.
//! Parameter lists are handled in the parameters module.
//!
//! [Declare Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Visual Basic 6 Declare statement with syntax:
    ///
    /// \[ Public | Private \] Declare { Sub | Function } name Lib "libname" \[ Alias "aliasname" \] \[ ( arglist ) \] \[ As type \]
    ///
    /// The Declare statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the Declare statement is accessible to all other procedures in all modules. |
    /// | Private     | Optional | Indicates that the Declare statement is accessible only to other procedures in the module where it is declared. |
    /// | Sub         | Required | Indicates that the procedure doesn't return a value. |
    /// | Function    | Required | Indicates that the procedure returns a value that can be used in an expression. |
    /// | name        | Required | Name of the external procedure; follows standard variable naming conventions. |
    /// | Lib         | Required | Indicates that a DLL or code resource contains the procedure being declared. The Lib clause is required for all declarations. |
    /// | libname     | Required | Name of the DLL or code resource that contains the declared procedure. |
    /// | Alias       | Optional | Indicates that the procedure being called has another name in the DLL. This is useful when the external procedure name is the same as a keyword. |
    /// | aliasname   | Optional | Name of the procedure in the DLL or code resource. If the first character is not a number sign (#), aliasname is the name of the procedure's entry point in the DLL. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the procedure when it is called. |
    /// | type        | Optional | Data type of the value returned by a Function procedure; may be Byte, Boolean, Integer, Long, Currency, Single, Double, Decimal, Date, String, Object, Variant, or any user-defined type. |
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    pub(super) fn parse_declare_statement(&mut self) {
        // Declare statements are only valid in the header section
        self.builder
            .start_node(SyntaxKind::DeclareStatement.to_raw());

        // Consume optional Public/Private keyword
        if self.at_token(VB6Token::PublicKeyword) || self.at_token(VB6Token::PrivateKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Declare" keyword
        self.consume_token();

        // Consume any whitespace after "Declare"
        self.consume_whitespace();

        // Consume "Sub" or "Function" keyword
        if self.at_token(VB6Token::SubKeyword) || self.at_token(VB6Token::FunctionKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Sub/Function
        self.consume_whitespace();

        // Consume procedure name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before Lib
        self.consume_whitespace();

        // Consume "Lib" keyword
        if self.at_token(VB6Token::LibKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Lib
        self.consume_whitespace();

        // Consume library name string
        if self.at_token(VB6Token::StringLiteral) {
            self.consume_token();
        }

        // Consume any whitespace after library name
        self.consume_whitespace();

        // Consume optional Alias clause
        if self.at_token(VB6Token::AliasKeyword) {
            self.consume_token();

            // Consume any whitespace after Alias
            self.consume_whitespace();

            // Consume alias name string
            if self.at_token(VB6Token::StringLiteral) {
                self.consume_token();
            }

            // Consume any whitespace after alias name
            self.consume_whitespace();
        }

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (includes "As Type" if present for Function)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // DeclareStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn declare_function_simple() {
        // Test simple Declare Function without parameters
        let source = "Declare Function GetTickCount Lib \"kernel32\" () As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("Declare Function GetTickCount"));
        assert!(cst.text().contains("Lib"));
        assert!(cst.text().contains("kernel32"));
    }

    #[test]
    fn declare_sub_simple() {
        // Test simple Declare Sub without parameters
        let source = "Declare Sub Sleep Lib \"kernel32\" (ByVal dwMilliseconds As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("Declare Sub Sleep"));
        assert!(cst.text().contains("Lib"));
    }

    #[test]
    fn declare_public_function() {
        // Test Public Declare Function
        let source = "Public Declare Function BitBlt Lib \"gdi32\" (ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("Public Declare Function BitBlt"));
        assert!(cst.text().contains("gdi32"));
    }

    #[test]
    fn declare_private_function() {
        // Test Private Declare Function
        let source = "Private Declare Function GetPixel Lib \"gdi32\" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("Private Declare Function GetPixel"));
    }

    #[test]
    fn declare_with_alias() {
        // Test Declare with Alias clause
        let source = "Private Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("CopyMemory"));
        assert!(cst.text().contains("Alias"));
        assert!(cst.text().contains("RtlMoveMemory"));
    }

    #[test]
    fn declare_with_alias_and_params() {
        // Test Declare with Alias and multiple parameters
        let source = "Private Declare Function SendMessageTimeout Lib \"user32\" Alias \"SendMessageTimeoutA\" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("SendMessageTimeout"));
        assert!(cst.text().contains("SendMessageTimeoutA"));
    }

    #[test]
    fn declare_lib_with_dll_extension() {
        // Test Declare with .dll extension in library name
        let source = "Private Declare Sub ZeroMemory Lib \"kernel32.dll\" Alias \"RtlZeroMemory\" (ByRef Destination As Any, ByVal Length As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("kernel32.dll"));
    }

    #[test]
    fn declare_no_parameters() {
        // Test Declare Function with no parameters
        let source = "Public Declare Function GetLastError Lib \"kernel32\" () As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("GetLastError"));
    }

    #[test]
    fn declare_byval_byref_params() {
        // Test Declare with ByVal and ByRef parameters
        let source = "Private Declare Function CallWindowProcW Lib \"user32\" (ByRef lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("ByRef lpPrevWndFunc"));
        assert!(cst.text().contains("ByVal hwnd"));
    }

    #[test]
    fn declare_any_type() {
        // Test Declare with Any type parameters
        let source = "Private Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (Destination As Any, Source As Any, ByVal Length As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("As Any"));
    }

    #[test]
    fn declare_long_parameters() {
        // Test Declare with many parameters (like StretchBlt)
        let source = "Public Declare Function StretchBlt Lib \"GDI32\" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal ClipX As Long, ByVal ClipY As Long, ByVal RasterOp As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("StretchBlt"));
        assert!(cst.text().contains("GDI32"));
    }

    #[test]
    fn declare_sub_no_return_type() {
        // Test Declare Sub doesn't have return type
        let source = "Private Declare Sub GdiplusShutdown Lib \"GdiPlus.dll\" (ByVal mtoken As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("Declare Sub"));
        assert!(cst.text().contains("GdiplusShutdown"));
    }

    #[test]
    fn declare_function_string_return() {
        // Test Declare Function returning String
        let source = "Public Declare Function GetUserName Lib \"advapi32.dll\" Alias \"GetUserNameA\" (ByVal lpBuffer As String, nSize As Long) As Long\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("As String"));
    }

    #[test]
    fn declare_multiple_statements() {
        // Test multiple Declare statements in sequence
        let source = "Private Declare Function VirtualProtect Lib \"kernel32\" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long\nPrivate Declare Sub RtlMoveMemory Lib \"ntdll\" (ByVal pDst As Long, ByVal pSrc As Long, ByVal dwLength As Long)\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 2);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        if let Some(child) = cst.child_at(1) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("VirtualProtect"));
        assert!(cst.text().contains("RtlMoveMemory"));
    }

    #[test]
    fn declare_uppercase_lib() {
        // Test Declare with uppercase library name
        let source = "Public Declare Function SetPixelV Lib \"gdi32\" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.child_count(), 1);
        if let Some(child) = cst.child_at(0) {
            assert_eq!(child.kind, SyntaxKind::DeclareStatement);
        }
        assert!(cst.text().contains("As Byte"));
    }
}
