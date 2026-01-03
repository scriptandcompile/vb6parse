//! `Declare` statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 declaration statements:
//! - `Declare`: External function/sub declarations
//! - `Event`: Custom event declarations in classes
//! - `Implements`: Interface implementation declarations
//!
//! `Declare` statement syntax:
//! `\[ Public | Private \] Declare { Sub | Function } name Lib "libname" \[ Alias "aliasname" \] \[ ( arglist ) \] \[ As type \]`
//!
//! `Event` statement syntax:
//! `\[ Public \] Event eventname \[ ( arglist ) \]`
//!
//! `Implements` statement syntax:
//! `Implements interfacename`
//!
//! `Sub` statements are handled in the `sub_statements` module.
//! `Function` statements are handled in the `function_statements` module.
//! `Dim`/`ReDim` and general `Variable` declarations are handled in the `array_statements` module.
//! `Property` statements are handled in the `property_statements` module.
//! `Parameter` lists are handled in the `parameters` module.
//!
//! [Declare Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
//! [Event Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/event-statement)
//! [Implements Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/implements-statement)

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Visual Basic 6 `Declare` statement with syntax:
    ///
    /// `\[ Public | Private \] Declare { Sub | Function } name Lib "libname" \[ Alias "aliasname" \] \[ ( arglist ) \] \[ As type \]`
    ///
    /// The `Declare` statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the `Declare` statement is accessible to all other procedures in all modules. |
    /// | Private     | Optional | Indicates that the `Declare` statement is accessible only to other procedures in the module where it is declared. |
    /// | Sub         | Required | Indicates that the procedure doesn't return a value. |
    /// | Function    | Required | Indicates that the procedure returns a value that can be used in an expression. |
    /// | name        | Required | Name of the external procedure; follows standard variable naming conventions. |
    /// | Lib         | Required | Indicates that a DLL or code resource contains the procedure being declared. The `Lib` clause is required for all declarations. |
    /// | libname     | Required | Name of the DLL or code resource that contains the declared procedure. |
    /// | Alias       | Optional | Indicates that the procedure being called has another name in the DLL. This is useful when the external procedure name is the same as a keyword. |
    /// | aliasname   | Optional | Name of the procedure in the DLL or code resource. If the first character is not a number sign (#), aliasname is the name of the procedure's entry point in the DLL. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the procedure when it is called. |
    /// | type        | Optional | Data type of the value returned by a Function procedure; may be `Byte`, `Boolean`, `Integer`, `Long`, `Currency`, `Single`, `Double`, `Decimal`, `Date`, `String`, `Object`, `Variant`, or any user-defined type. |
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// `\[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \]`
    ///
    /// [Reference](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa243324(v=vs.60))
    pub(super) fn parse_declare_statement(&mut self) {
        // Declare statements are only valid in the header section
        self.builder
            .start_node(SyntaxKind::DeclareStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume optional Public/Private keyword
        if self.at_token(Token::PublicKeyword) || self.at_token(Token::PrivateKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Declare" keyword
        self.consume_token();

        // Consume any whitespace after "Declare"
        self.consume_whitespace();

        // Consume "Sub" or "Function" keyword
        if self.at_token(Token::SubKeyword) || self.at_token(Token::FunctionKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Sub/Function
        self.consume_whitespace();

        // Consume procedure name (keywords can be used as procedure names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume any whitespace before Lib
        self.consume_whitespace();

        // Consume "Lib" keyword
        if self.at_token(Token::LibKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Lib
        self.consume_whitespace();

        // Consume library name string
        if self.at_token(Token::StringLiteral) {
            self.consume_token();
        }

        // Consume any whitespace after library name
        self.consume_whitespace();

        // Consume optional Alias clause
        if self.at_token(Token::AliasKeyword) {
            self.consume_token();

            // Consume any whitespace after Alias
            self.consume_whitespace();

            // Consume alias name string
            if self.at_token(Token::StringLiteral) {
                self.consume_token();
            }

            // Consume any whitespace after alias name
            self.consume_whitespace();
        }

        // Parse parameter list if present
        if self.at_token(Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (includes "As Type" if present for Function)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // DeclareStatement
    }

    /// Parse a Visual Basic 6 `Event` statement with syntax:
    ///
    /// `\[ Public \] Event eventname \[ ( arglist ) \]`
    ///
    /// The `Event` statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public      | Optional | Indicates that the `Event` is accessible to all other procedures in all modules. Events are Public by default. Note that events can't be Private. |
    /// | eventname   | Required | Name of the event; follows standard variable naming conventions. |
    /// | arglist     | Optional | List of variables representing arguments that are passed to the event handler when the event occurs. |
    ///
    /// The `arglist` argument has the following syntax and parts:
    ///
    /// `\[ ByVal | ByRef \] varname \[ ( ) \] \[ As type \]`
    ///
    /// Remarks:
    /// - `Event` statements can appear only in class modules, form modules, and user controls.
    /// - `Event`s are raised using the `RaiseEvent` statement.
    /// - `Event`s declared with `Public` are available to all procedures in the same project.
    /// - `Event`s cannot be declared as `Private`, `Static`, or `Friend`.
    /// - `Event`s cannot have named arguments, `Optional` arguments, or `ParamArray` arguments.
    /// - `Event`s do not have return values.
    ///
    /// Examples:
    /// ```vb
    /// Public Event StatusChanged(ByVal NewStatus As String)
    /// Event DataReceived(ByVal Data() As Byte)
    /// Event Click()
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/event-statement)
    pub(super) fn parse_event_statement(&mut self) {
        // Event statements are only valid in class modules
        self.builder.start_node(SyntaxKind::EventStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume optional Public keyword
        if self.at_token(Token::PublicKeyword) {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume "Event" keyword
        self.consume_token();

        // Consume any whitespace after "Event"
        self.consume_whitespace();

        // Consume event name (keywords can be used as event names in VB6)
        if self.at_token(Token::Identifier) {
            self.consume_token();
        } else if self.at_keyword() {
            self.consume_token_as_identifier();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // EventStatement
    }

    /// Parse an Implements statement.
    ///
    /// VB6 Implements statement syntax:
    /// - Implements interfacename
    ///
    /// Specifies an interface or class that will be implemented in the class module in which it appears.
    ///
    /// The Implements statement syntax has this part:
    ///
    /// | Part          | Description |
    /// |---------------|-------------|
    /// | interfacename | Required. Name of an interface or class whose methods and properties will be implemented in the class containing the Implements statement. |
    ///
    /// Remarks:
    /// - The Implements statement is used only in class modules.
    /// - Once you have specified that a class implements an interface, you must provide a procedure in the class for each public procedure defined in the interface.
    /// - The procedure in the implementing class must have the same name and signature as the procedure in the interface.
    /// - A class module can implement more than one interface by including a separate Implements statement for each interface.
    /// - The interface must be defined in a separate class module.
    /// - You can't implement an interface within a single class module.
    ///
    /// Examples:
    /// ```vb
    /// ' In class module clsInterface:
    /// Public Sub DoSomething(x As Integer)
    /// End Sub
    ///
    /// ' In implementing class:
    /// Implements clsInterface
    ///
    /// Private Sub clsInterface_DoSomething(x As Integer)
    ///     ' Implementation code
    /// End Sub
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/implements-statement)
    pub(super) fn parse_implements_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::ImplementsStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "Implements" keyword
        self.consume_token();

        // Consume everything until newline (the interface name)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // ImplementsStatement
    }

    /// Parse an Object statement.
    ///
    /// VB6 Object statement syntax:
    /// - Object = "{UUID}#version#flags"; "filename"
    /// - Object = *\G{UUID}#version#flags; "filename"
    ///
    /// Declares external `ActiveX` controls and libraries that a form or user control depends on.
    ///
    /// The Object statement syntax has these parts:
    ///
    /// | Part     | Description |
    /// |----------|-------------|
    /// | UUID     | Required. The class ID (CLSID) of the `ActiveX` control, enclosed in braces. |
    /// | version  | Required. The version number of the control (e.g., "2.0"). |
    /// | flags    | Required. Additional flags (typically "0"). |
    /// | filename | Required. The filename of the OCX or DLL containing the control. |
    ///
    /// Remarks:
    /// - Object statements appear in form (.frm) and user control (.ctl) files.
    /// - They are placed at the module level, after VERSION and before BEGIN.
    /// - Multiple Object statements can appear to declare dependencies on multiple controls.
    /// - The format can use either "{UUID}" or "*\G{UUID}" prefix.
    /// - These declarations are automatically maintained by the VB6 IDE when you add controls to a form.
    ///
    /// Examples:
    /// ```vb
    /// Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
    /// Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
    /// Object = "*\G{00025600-0000-0000-C000-000000000046}#5.2#0"; "stdole2.tlb"
    /// ```
    pub(super) fn parse_object_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::ObjectStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "Object" keyword
        self.consume_token();

        // Consume everything until newline (=, UUID string, semicolon, filename)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // ObjectStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn declare_function_simple() {
        // Test simple Declare Function without parameters
        let source = "Declare Function GetTickCount Lib \"kernel32\" () As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetTickCount"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_sub_simple() {
        // Test simple Declare Sub without parameters
        let source = "Declare Sub Sleep Lib \"kernel32\" (ByVal dwMilliseconds As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Sleep"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("dwMilliseconds"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_public_function() {
        // Test Public Declare Function
        let source = "Public Declare Function BitBlt Lib \"gdi32\" (ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PublicKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("BitBlt"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"gdi32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hDstDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("nWidth"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("nHeight"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hSrcDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("xSrc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("ySrc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("dwRop"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_private_function() {
        // Test Private Declare Function
        let source = "Private Declare Function GetPixel Lib \"gdi32\" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetPixel"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"gdi32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_with_alias() {
        // Test Declare with Alias clause
        let source = "Private Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("CopyMemory"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                AliasKeyword,
                Whitespace,
                StringLiteral ("\"RtlMoveMemory\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Dest"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Source"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Bytes"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_with_alias_and_params() {
        // Test Declare with Alias and multiple parameters
        let source = "Private Declare Function SendMessageTimeout Lib \"user32\" Alias \"SendMessageTimeoutA\" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("SendMessageTimeout"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"user32\""),
                Whitespace,
                AliasKeyword,
                Whitespace,
                StringLiteral ("\"SendMessageTimeoutA\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hwnd"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("wMsg"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("wParam"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("lParam"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("fuFlags"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("uTimeout"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("lpdwResult"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_lib_with_dll_extension() {
        // Test Declare with .dll extension in library name
        let source = "Private Declare Sub ZeroMemory Lib \"kernel32.dll\" Alias \"RtlZeroMemory\" (ByRef Destination As Any, ByVal Length As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("ZeroMemory"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32.dll\""),
                Whitespace,
                AliasKeyword,
                Whitespace,
                StringLiteral ("\"RtlZeroMemory\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Destination"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Length"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_no_parameters() {
        // Test Declare Function with no parameters
        let source = "Public Declare Function GetLastError Lib \"kernel32\" () As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PublicKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetLastError"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_byval_byref_params() {
        // Test Declare with ByVal and ByRef parameters
        let source = "Private Declare Function CallWindowProcW Lib \"user32\" (ByRef lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("CallWindowProcW"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"user32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("lpPrevWndFunc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hwnd"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("msg"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("wParam"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("lParam"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_any_type() {
        // Test Declare with Any type parameters
        let source = "Private Declare Sub CopyMemory Lib \"kernel32\" Alias \"RtlMoveMemory\" (Destination As Any, Source As Any, ByVal Length As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("CopyMemory"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                AliasKeyword,
                Whitespace,
                StringLiteral ("\"RtlMoveMemory\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    Identifier ("Destination"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    Identifier ("Source"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Any"),
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Length"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_long_parameters() {
        // Test Declare with many parameters (like StretchBlt)
        let source = "Public Declare Function StretchBlt Lib \"GDI32\" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal ClipX As Long, ByVal ClipY As Long, ByVal RasterOp As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PublicKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("StretchBlt"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"GDI32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hDestDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("x"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("nWidth"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("nHeight"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hSrcDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("xSrc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("ySrc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("ClipX"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("ClipY"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("RasterOp"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_sub_no_return_type() {
        // Test Declare Sub doesn't have return type
        let source =
            "Private Declare Sub GdiplusShutdown Lib \"GdiPlus.dll\" (ByVal mtoken As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("GdiplusShutdown"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"GdiPlus.dll\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("mtoken"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_function_string_return() {
        // Test Declare Function returning String
        let source = "Public Declare Function GetUserName Lib \"advapi32.dll\" Alias \"GetUserNameA\" (ByVal lpBuffer As String, nSize As Long) As Long\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PublicKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetUserName"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"advapi32.dll\""),
                Whitespace,
                AliasKeyword,
                Whitespace,
                StringLiteral ("\"GetUserNameA\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("lpBuffer"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("nSize"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_multiple_statements() {
        // Test multiple Declare statements in sequence
        let source = "Private Declare Function VirtualProtect Lib \"kernel32\" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long\nPrivate Declare Sub RtlMoveMemory Lib \"ntdll\" (ByVal pDst As Long, ByVal pSrc As Long, ByVal dwLength As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("VirtualProtect"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"kernel32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("lpAddress"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("dwSize"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("flNewProtect"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("lpflOldProtect"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
            },
            DeclareStatement {
                PrivateKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("RtlMoveMemory"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"ntdll\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("pDst"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("pSrc"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("dwLength"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn declare_uppercase_lib() {
        // Test Declare with uppercase library name
        let source = "Public Declare Function SetPixelV Lib \"gdi32\" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DeclareStatement {
                PublicKeyword,
                Whitespace,
                DeclareKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("SetPixelV"),
                Whitespace,
                LibKeyword,
                Whitespace,
                StringLiteral ("\"gdi32\""),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("hDC"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("X"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Y"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("crColor"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                ByteKeyword,
                Newline,
            },
        ]);
    }

    // Event statement tests

    #[test]
    fn event_simple() {
        let source = "Event StatusChanged()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("StatusChanged"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_with_parameter() {
        let source = "Event DataReceived(ByVal Data As String)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("DataReceived"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_public() {
        let source = "Public Event Click()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                PublicKeyword,
                Whitespace,
                EventKeyword,
                Whitespace,
                Identifier ("Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_multiple_parameters() {
        let source = "Event ValueChanged(ByVal OldValue As Long, ByVal NewValue As Long)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("ValueChanged"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("OldValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("NewValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_with_array_parameter() {
        let source = "Event DataReceived(ByVal Data() As Byte)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("DataReceived"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Data"),
                    LeftParenthesis,
                    RightParenthesis,
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ByteKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_no_parameters() {
        let source = "Public Event Initialize()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                PublicKeyword,
                Whitespace,
                EventKeyword,
                Whitespace,
                Identifier ("Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_byref_parameter() {
        let source = "Event Modified(ByRef Cancel As Boolean)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("Modified"),
                ParameterList {
                    LeftParenthesis,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Cancel"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    BooleanKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_preserves_whitespace() {
        let source = "    Event    Test    (    )    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("Test"),
                Whitespace,
                ParameterList {
                    LeftParenthesis,
                    Whitespace,
                    RightParenthesis,
                },
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn event_complex_parameters() {
        let source = "Public Event ProgressUpdate(ByVal PercentComplete As Integer, ByVal Message As String, ByRef Cancel As Boolean)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                PublicKeyword,
                Whitespace,
                EventKeyword,
                Whitespace,
                Identifier ("ProgressUpdate"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("PercentComplete"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Message"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    ByRefKeyword,
                    Whitespace,
                    Identifier ("Cancel"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    BooleanKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_object_parameter() {
        let source = "Event ItemAdded(ByVal Item As Object)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("ItemAdded"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Item"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ObjectKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn multiple_event_declarations() {
        let source = "Event Click()\nEvent DblClick()\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("Click"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
            },
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("DblClick"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_variant_parameter() {
        let source = "Event DataChanged(ByVal NewData As Variant)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("DataChanged"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("NewData"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    VariantKeyword,
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn event_custom_type_parameter() {
        let source = "Event RecordChanged(ByVal Record As CustomerRecord)\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            EventStatement {
                EventKeyword,
                Whitespace,
                Identifier ("RecordChanged"),
                ParameterList {
                    LeftParenthesis,
                    ByValKeyword,
                    Whitespace,
                    Identifier ("Record"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("CustomerRecord"),
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn object_statement_single() {
        let source = r#"Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.frm", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            ObjectStatement {
                ObjectKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\""),
                Semicolon,
                Whitespace,
                StringLiteral ("\"mscomctl.ocx\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn object_statement_multiple() {
        let source = r#"VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.frm", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            VersionStatement {
                VersionKeyword,
                Whitespace,
                SingleLiteral,
                Newline,
            },
            ObjectStatement {
                ObjectKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\""),
                Semicolon,
                Whitespace,
                StringLiteral ("\"mscomctl.ocx\""),
                Newline,
            },
            ObjectStatement {
                ObjectKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteral ("\"{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0\""),
                Semicolon,
                Whitespace,
                StringLiteral ("\"COMDLG32.OCX\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn object_statement_with_backslash_g_prefix() {
        let source = "Object = *\\G{00025600-0000-0000-C000-000000000046}#5.2#0; \"stdole2.tlb\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.frm", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            ObjectStatement {
                ObjectKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                MultiplicationOperator,
                BackwardSlashOperator,
                Identifier ("G"),
                LeftCurlyBrace,
                IntegerLiteral ("00025600"),
                SubtractionOperator,
                IntegerLiteral ("0000"),
                SubtractionOperator,
                IntegerLiteral ("0000"),
                SubtractionOperator,
                Identifier ("C000"),
                SubtractionOperator,
                IntegerLiteral ("000000000046"),
                RightCurlyBrace,
                DateLiteral ("#5.2#"),
                IntegerLiteral ("0"),
                Semicolon,
                Whitespace,
                StringLiteral ("\"stdole2.tlb\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn object_variable_assignment_at_module_level() {
        // At module level, "Object = value" where value is not a GUID string
        // should NOT be parsed as an Object statement (Object statements have very specific format)
        // VB6 doesn't actually allow regular assignment statements at module level in reality,
        // but the parser should handle keywords as identifiers when followed by =
        let source = "Object = 5\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Unknown,
            Whitespace,
            AssignmentStatement {
                IdentifierExpression {
                    EqualityOperator,
                },
                Whitespace,
                NumericLiteralExpression {
                    IntegerLiteral ("5"),
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn object_variable_assignment_in_sub() {
        // Test that an assignment to a variable named "Object" inside a Sub
        // is parsed as an assignment, NOT an Object statement
        let source = "Sub Test()\n    Object = 5\nEnd Sub\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    Unknown,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            EqualityOperator,
                        },
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("5"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
