//! Declaration parsing for VB6 CST.
//!
//! This module handles parsing of VB6 declarations:
//! - Function statements
//! - Sub statements
//! - Property statements (Get, Let, Set)
//! - Parameter lists
//! 
//! Dim/ReDim and general Variable declarations are handled in the array_statements module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Visual Basic 6 subroutine with syntax:
    ///
    /// \[ Public | Private | Friend \] \[ Static \] Sub name \[ ( arglist ) \]
    /// \[ statements \]
    /// \[ Exit Sub \]
    /// \[ statements \]
    /// End Sub
    ///
    /// The Sub statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public   	  | Optional | Indicates that the Sub procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private statement, the procedure is not available outside the project. |
    /// | Private  	  | Optional | Indicates that the Sub procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend 	  | Optional | Used only in a class module. Indicates that the Sub procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static 	  | Optional | Indicates that the Sub procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Sub, even if they are used in the procedure. |
    /// | name 	      | Required | Name of the Sub; follows standard variable naming conventions. |
    /// | arglist 	  | Optional | List of variables representing arguments that are passed to the Sub procedure when it is called. Multiple variables are separated by commas. |
    /// | statements  | Optional | Any group of statements to be executed within the Sub procedure.
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sub-statement)
    pub(super) fn parse_sub_statement(&mut self) {
        // if we are now parsing a sub statement, we are no longer in the header.
        self.parsing_header = false;
        self.builder.start_node(SyntaxKind::SubStatement.to_raw());

        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword)
            || self.at_token(VB6Token::PrivateKeyword)
            || self.at_token(VB6Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Sub" keyword
        self.consume_token();

        // Consume any whitespace after "Sub"
        self.consume_whitespace();

        // Consume procedure name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End Sub"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::SubKeyword)
        });

        // Consume "End Sub" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Sub"
            self.consume_whitespace();

            // Consume "Sub"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // SubStatement
    }

    /// Parse a Visual Basic 6 function with syntax:
    ///
    /// \[ Public | Private | Friend \] \[ Static \] Function name \[ ( arglist ) \] \[ As type \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// \[ Exit Function \]
    /// \[ statements \]
    /// \[ name = expression \]
    /// End Function
    ///
    /// The Function statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Public   	  | Optional | Indicates that the Function procedure is accessible to all other procedures in all modules. If used in a module that contains an Option Private, the procedure is not available outside the project. |
    /// | Private  	  | Optional | Indicates that the Function procedure is accessible only to other procedures in the module where it is declared. |
    /// | Friend 	  | Optional | Used only in a class module. Indicates that the Function procedure is visible throughout the project, but not visible to a controller of an instance of an object. |
    /// | Static 	  | Optional | Indicates that the Function procedure's local variables are preserved between calls. The Static attribute doesn't affect variables that are declared outside the Function, even if they are used in the procedure. |
    /// | name 	      | Required | Name of the Function; follows standard variable naming conventions. |
    /// | arglist 	  | Optional | List of variables representing arguments that are passed to the Function procedure when it is called. Multiple variables are separated by commas. |
    /// | type 	      | Optional | Data type of the value returned by the Function procedure; may be Byte, Boolean, Integer, Long, Currency, Single, Double, Decimal (not currently supported), Date, String (except fixed length), Object, Variant, or any user-defined type. |
    /// | statements  | Optional | Any group of statements to be executed within the Function procedure.
    /// | expression  | Optional | Return value of the Function. |
    ///
    /// The arglist argument has the following syntax and parts:
    ///
    /// \[ Optional \] \[ ByVal | ByRef \] \[ ParamArray \] varname \[ ( ) \] \[ As type \] \[ = defaultvalue \]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/function-statement)
    pub(super) fn parse_function_statement(&mut self) {
        // if we are now parsing a function statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::FunctionStatement.to_raw());

        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword)
            || self.at_token(VB6Token::PrivateKeyword)
            || self.at_token(VB6Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Function" keyword
        self.consume_token();

        // Consume any whitespace after "Function"
        self.consume_whitespace();

        // Consume function name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (includes "As Type" if present)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End Function"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::FunctionKeyword)
        });

        // Consume "End Function" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Function"
            self.consume_whitespace();

            // Consume "Function"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // FunctionStatement
    }

    /// Parse a Property statement (Property Get, Property Let, or Property Set).
    ///
    /// VB6 Property statement syntax:
    /// - [Public | Private | Friend] [Static] Property Get name [(arglist)] [As type]
    /// - [Public | Private | Friend] [Static] Property Let name ([arglist,] value)
    /// - [Public | Private | Friend] [Static] Property Set name ([arglist,] value)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-get-statement)
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-let-statement)
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/property-set-statement)
    pub(super) fn parse_property_statement(&mut self) {
        // if we are now parsing a property statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::PropertyStatement.to_raw());

        // Consume optional Public/Private/Friend keyword
        if self.at_token(VB6Token::PublicKeyword)
            || self.at_token(VB6Token::PrivateKeyword)
            || self.at_token(VB6Token::FriendKeyword)
        {
            self.consume_token();

            // Consume any whitespace after visibility modifier
            self.consume_whitespace();
        }

        // Consume optional Static keyword
        if self.at_token(VB6Token::StaticKeyword) {
            self.consume_token();

            // Consume any whitespace after Static
            self.consume_whitespace();
        }

        // Consume "Property" keyword
        self.consume_token();

        // Consume any whitespace after "Property"
        self.consume_whitespace();

        // Consume Get/Let/Set keyword
        if self.at_token(VB6Token::GetKeyword)
            || self.at_token(VB6Token::LetKeyword)
            || self.at_token(VB6Token::SetKeyword)
        {
            self.consume_token();
        }

        // Consume any whitespace after Get/Let/Set
        self.consume_whitespace();

        // Consume property name
        if self.at_token(VB6Token::Identifier) {
            self.consume_token();
        }

        // Consume any whitespace before parameter list
        self.consume_whitespace();

        // Parse parameter list if present
        if self.at_token(VB6Token::LeftParenthesis) {
            self.parse_parameter_list();
        }

        // Consume everything until newline (includes "As Type" if present)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End Property"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::PropertyKeyword)
        });

        // Consume "End Property" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Property"
            self.consume_whitespace();

            // Consume "Property"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);

            // Consume the newline
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // PropertyStatement
    }

    /// Parse a parameter list: (param1 As Type, param2 As Type)
    pub(super) fn parse_parameter_list(&mut self) {
        self.builder.start_node(SyntaxKind::ParameterList.to_raw());

        // Consume "("
        self.consume_token();

        // Consume everything until ")"
        let mut depth = 1;
        while !self.is_at_end() && depth > 0 {
            if self.at_token(VB6Token::LeftParenthesis) {
                depth += 1;
            } else if self.at_token(VB6Token::RightParenthesis) {
                depth -= 1;
            }

            self.consume_token();

            if depth == 0 {
                break;
            }
        }

        self.builder.finish_node(); // ParameterList
    }

}
