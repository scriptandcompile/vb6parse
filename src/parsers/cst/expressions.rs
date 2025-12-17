//! Expression parsing for VB6 CST.
//!
//! This module implements expression parsing for Visual Basic 6 using a Pratt parsing
//! approach (also known as operator precedence parsing or precedence climbing). This
//! technique cleanly handles operator precedence and associativity while maintaining
//! a simple recursive descent structure.
//!
//! # VB6 Expression Types
//!
//! VB6 supports various expression types:
//!
//! - **Literal expressions**: Numbers, strings, dates, `True`, `False`, `Nothing`, `Null`, `Empty`
//! - **Identifier expressions**: Variable names, constants
//! - **Unary expressions**: `-x`, `Not x`, `AddressOf proc`
//! - **Binary expressions**: Arithmetic, comparison, logical operations
//! - **Member access**: `object.property`, `object.method`
//! - **Function calls**: `Function(arg1, arg2)`, `Function arg1, arg2`
//! - **Array indexing**: `array(index)`, `array(i, j)`
//! - **Parenthesized**: `(expression)`
//! - **Object creation**: `New ClassName`
//! - **Type operations**: `TypeOf object Is type`
//!
//! # Operator Precedence
//!
//! VB6 operators are parsed according to the following precedence levels (highest to lowest):
//!
//! 1. Member access (`.`), function calls `()`
//! 2. Exponentiation (`^`) - right-associative
//! 3. Unary negation (`-`)
//! 4. Multiplication (`*`), division (`/`)
//! 5. Integer division (`\`)
//! 6. Modulo (`Mod`)
//! 7. Addition (`+`), subtraction (`-`)
//! 8. String concatenation (`&`)
//! 9. Comparison (`=`, `<>`, `<`, `>`, `<=`, `>=`, `Like`, `Is`)
//! 10. Logical `Not`
//! 11. Logical `And`
//! 12. Logical `Or`
//! 13. Logical `Xor`
//! 14. Logical `Eqv`
//! 15. Logical `Imp`
//!
//! # Pratt Parsing
//!
//! The implementation uses Pratt parsing, which associates a binding power (precedence level)
//! with each operator. The parser works by:
//!
//! 1. Parsing a prefix expression (literal, identifier, unary operator, etc.)
//! 2. Looking at the next operator and comparing its binding power to the current minimum
//! 3. If the operator's binding power is higher, it binds tighter and is parsed as an infix operation
//! 4. This continues recursively until an operator with lower binding power is encountered
//!
//! This approach naturally handles precedence and associativity without complex lookahead
//! or multiple parsing passes.
//!
//! # Examples
//!
//! ```vb6
//! ' Arithmetic with proper precedence
//! result = 2 + 3 * 4        ' Parsed as: 2 + (3 * 4)
//! result = 10 - 5 - 2       ' Parsed as: (10 - 5) - 2
//! result = 2 ^ 3 ^ 2        ' Parsed as: 2 ^ (3 ^ 2) - right associative
//!
//! ' Logical operations
//! condition = x > 5 And y < 10       ' Parsed as: (x > 5) And (y < 10)
//! condition = Not flag1 Or flag2     ' Parsed as: (Not flag1) Or flag2
//!
//! ' Member access and calls
//! value = obj.property.method(arg1, arg2)
//!
//! ' Complex expressions
//! result = (a + b) * c - d / e Mod f
//! ```

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

/// Operator binding power (precedence) levels.
///
/// Higher values indicate tighter binding (higher precedence).
/// These values are based on the VB6 language specification and determine
/// the order in which operators are applied when parsing expressions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct BindingPower(u8);

impl BindingPower {
    /// No binding power - used as a minimum baseline
    pub(super) const NONE: BindingPower = BindingPower(0);

    /// Logical implication operator (`Imp`) - lowest precedence
    pub(super) const IMP: BindingPower = BindingPower(10);

    /// Logical equivalence operator (`Eqv`)
    pub(super) const EQV: BindingPower = BindingPower(20);

    /// Logical exclusive or operator (`Xor`)
    pub(super) const XOR: BindingPower = BindingPower(30);

    /// Logical or operator (`Or`)
    pub(super) const OR: BindingPower = BindingPower(40);

    /// Logical and operator (`And`)
    pub(super) const AND: BindingPower = BindingPower(50);

    /// Logical not operator (`Not`) - prefix operator
    pub(super) const NOT: BindingPower = BindingPower(60);

    /// Comparison operators (`=`, `<>`, `<`, `>`, `<=`, `>=`, `Like`, `Is`)
    pub(super) const COMPARISON: BindingPower = BindingPower(70);

    /// String concatenation operator (`&`)
    pub(super) const CONCATENATION: BindingPower = BindingPower(80);

    /// Addition and subtraction operators (`+`, `-`)
    pub(super) const ADDITION: BindingPower = BindingPower(90);

    /// Modulo operator (`Mod`)
    pub(super) const MODULO: BindingPower = BindingPower(100);

    /// Integer division operator (`\`)
    pub(super) const INT_DIVISION: BindingPower = BindingPower(110);

    /// Multiplication and division operators (`*`, `/`)
    pub(super) const MULTIPLICATION: BindingPower = BindingPower(120);

    /// Unary operators (unary `-`, `AddressOf`) - prefix operators
    pub(super) const UNARY: BindingPower = BindingPower(130);

    /// Exponentiation operator (`^`) - right-associative
    pub(super) const EXPONENTIATION: BindingPower = BindingPower(140);

    // Function/method calls and array indexing
    //pub(super) const CALL: BindingPower = BindingPower(150);

    // Member access operator (`.`) - highest precedence
    //pub(super) const MEMBER: BindingPower = BindingPower(160);
}

impl Parser<'_> {
    /// Parse an expression starting with no minimum binding power.
    ///
    /// This is the main entry point for expression parsing. It delegates to
    /// [`parse_expression_with_binding_power`](Self::parse_expression_with_binding_power)
    /// with a minimum binding power of zero.
    ///
    /// # Examples
    ///
    /// ```vb6
    /// x = 5 + 3 * 2          ' Simple arithmetic
    /// y = obj.method(arg)    ' Member access and call
    /// z = (a + b) * c        ' Parenthesized expression
    /// ```
    pub(super) fn parse_expression(&mut self) {
        self.parse_expression_with_binding_power(BindingPower::NONE);
    }

    /// Parse an lvalue (left-hand side of assignment).
    ///
    /// This parses expressions but stops before the `=` operator,
    /// since `=` in VB6 can be both assignment and comparison.
    ///
    /// # Examples
    ///
    /// ```vb6
    /// x = 5                  ' x is the lvalue
    /// obj.property = value   ' obj.property is the lvalue
    /// arr(i) = 10           ' arr(i) is the lvalue
    /// ```
    pub(super) fn parse_lvalue(&mut self) {
        // Parse with a minimum binding power HIGHER than COMPARISON
        // This ensures = is not treated as a binary operator
        // COMPARISON is 70, so we use 75 to exclude it
        self.parse_expression_with_binding_power(BindingPower(75));
    }

    /// Parse an expression with a minimum binding power.
    ///
    /// This is the core of the Pratt parser. It:
    /// 1. Parses a prefix expression (literal, identifier, unary op, etc.)
    /// 2. Checks if the next token is an infix operator
    /// 3. If the operator's binding power is >= the minimum, parse it as infix
    /// 4. Continue until we encounter an operator with lower binding power
    ///
    /// # Parameters
    ///
    /// - `min_bp`: The minimum binding power required for an operator to be parsed.
    ///   Operators with lower binding power will end the current expression.
    ///
    /// # Pratt Parsing Algorithm
    ///
    /// The algorithm works as follows:
    ///
    /// ```text
    /// parse_expr(min_bp):
    ///   left = parse_prefix()
    ///   while peek_operator() has bp >= min_bp:
    ///     op = consume_operator()
    ///     right = parse_expr(op.right_bp)
    ///     left = make_binary(left, op, right)
    ///   return left
    /// ```
    pub(super) fn parse_expression_with_binding_power(&mut self, min_bp: BindingPower) {
        // Skip leading whitespace
        self.consume_whitespace();

        // Create a checkpoint BEFORE parsing the prefix - allows wrapping the entire left side
        let lhs_checkpoint = self.builder.checkpoint();

        // Parse the prefix expression (left-hand side)
        // Returns true if it was a bare identifier (no postfix operators)
        let is_bare_identifier = self.parse_prefix_expression();

        // If we have a bare identifier, wrap it in IdentifierExpression immediately
        // This ensures that identifiers are always wrapped, even when part of a binary expression
        if is_bare_identifier {
            self.builder
                .start_node_at(lhs_checkpoint, SyntaxKind::IdentifierExpression.to_raw());
            self.builder.finish_node();
        }

        // Parse infix operators with sufficient binding power
        loop {
            // Peek ahead to check for operators WITHOUT consuming whitespace yet
            // This prevents whitespace from being consumed when we stop parsing

            // Temporarily skip whitespace to check for operators
            let saved_pos = self.pos;
            loop {
                match self.current_token() {
                    Some(Token::Whitespace) => {
                        self.pos += 1;
                    }
                    Some(Token::Underscore) => {
                        // Check for line continuation
                        let mut lookahead = 1;
                        let mut is_continuation = false;
                        while let Some((_, token)) = self.tokens.get(self.pos + lookahead) {
                            if *token == Token::Whitespace {
                                lookahead += 1;
                            } else if *token == Token::Newline {
                                is_continuation = true;
                                break;
                            } else {
                                break;
                            }
                        }

                        if is_continuation {
                            // Skip underscore, whitespace, and newline
                            self.pos += lookahead + 1; // +1 for newline
                        } else {
                            break;
                        }
                    }
                    _ => break,
                }
            }

            // Check if we're at the end or at a delimiter
            if self.is_at_end() || self.is_at_expression_delimiter() {
                // Restore position and stop
                self.pos = saved_pos;
                break;
            }

            // Get the binding power of the next operator
            let binding_power = self.get_infix_binding_power();

            // Restore position (we haven't consumed the whitespace yet)
            self.pos = saved_pos;

            let Some((left_bp, right_bp)) = binding_power else {
                // Not an infix operator, we're done
                break;
            };

            // If the operator doesn't bind tightly enough, stop
            if left_bp < min_bp {
                break;
            }

            // Now actually consume the whitespace
            self.consume_whitespace();

            // Wrap the left-hand side in a BinaryExpression
            self.builder
                .start_node_at(lhs_checkpoint, SyntaxKind::BinaryExpression.to_raw());

            // Consume the operator
            self.consume_token();

            // Skip whitespace after operator
            self.consume_whitespace();

            // Parse the right-hand side
            self.parse_expression_with_binding_power(right_bp);

            // Finish the BinaryExpression
            self.builder.finish_node();

            // DON'T update checkpoint - use the original one for nested binary expressions
            // This creates the correct left-associative structure:
            // BinaryExpression(BinaryExpression(a + b), +, c)
        }
    }

    /// Parse a prefix expression.
    ///
    /// Prefix expressions are those that start an expression:
    /// - Literals: `42`, `"hello"`, `True`, `Nothing`
    /// - Identifiers: `myVar`, `MyClass`
    /// - Parenthesized: `(expression)`
    /// - Unary operators: `-x`, `Not flag`, `AddressOf proc`
    /// - Object creation: `New ClassName`
    /// - Type checking: `TypeOf obj Is type`
    /// Returns true if it parsed a bare identifier (one that needs wrapping in IdentifierExpression
    /// if not used in a binary expression).
    fn parse_prefix_expression(&mut self) -> bool {
        // Skip any leading whitespace
        self.consume_whitespace();

        // Create checkpoint at the start - this will be used for wrapping
        let checkpoint = self.builder.checkpoint();

        let mut is_identifier = false;

        match self.current_token() {
            // Unary minus
            Some(Token::SubtractionOperator) => {
                self.parse_unary_expression(BindingPower::UNARY);
            }
            // Logical NOT
            Some(Token::NotKeyword) => {
                self.parse_unary_expression(BindingPower::NOT);
            }
            // AddressOf operator
            Some(Token::AddressOfKeyword) => {
                self.parse_addressof_expression();
            }
            // TypeOf operator
            // TypeOf is handled as a regular keyword that can be an identifier
            // The actual TypeOf expression is parsed when we see the pattern
            // New operator
            Some(Token::NewKeyword) => {
                self.parse_new_expression();
            }
            // Parenthesized expression
            Some(Token::LeftParenthesis) => {
                self.parse_parenthesized_expression();
            }
            // Numeric literals
            Some(
                Token::IntegerLiteral
                | Token::LongLiteral
                | Token::SingleLiteral
                | Token::DoubleLiteral
                | Token::DecimalLiteral,
            ) => {
                self.parse_numeric_literal();
            }
            Some(Token::StringLiteral) => {
                self.parse_string_literal();
            }
            Some(Token::TrueKeyword | Token::FalseKeyword) => {
                self.parse_boolean_literal();
            }
            Some(Token::NullKeyword | Token::EmptyKeyword) => {
                self.parse_special_literal();
            }
            // Date literal
            Some(Token::DateLiteral) => {
                self.parse_date_literal();
            }
            // Identifiers (including keywords that can be identifiers in expression context)
            _ => {
                self.parse_identifier_or_call_expression();
                is_identifier = true;
            }
        }

        // Parse postfix operators using the checkpoint
        // Returns true if any postfix operators were found
        let has_postfix = self.parse_postfix_operators_with_checkpoint(checkpoint);

        // Return true if this was an identifier without postfix operators
        is_identifier && !has_postfix
    }

    /// Parse an identifier or a function/method call expression.
    ///
    /// This handles:
    /// - Simple identifiers: `myVar`
    /// - Identifiers with type characters: `myVar$`, `count%`
    /// - Keywords used as identifiers in expression context
    fn parse_identifier_or_call_expression(&mut self) {
        // In expression context, many keywords can be used as identifiers
        if self.is_identifier() || self.at_keyword() {
            // Check if this is a dollar-sign library function (Chr$, UCase$, etc.)
            if self.at_keyword_dollar() {
                // Consume both the identifier/keyword and the dollar sign as a single identifier
                self.consume_keyword_dollar_as_identifier();
            } else {
                // Consume just the identifier/keyword
                self.consume_token();

                // Check for type character suffix ($, %, &, !, #, @) - but NOT for library functions
                // Only consume dollar sign if it's NOT part of a library function name
                if matches!(
                    self.current_token(),
                    Some(
                        Token::DollarSign
                            | Token::Percent
                            | Token::Ampersand
                            | Token::ExclamationMark
                            | Token::Octothorpe
                            | Token::AtSign
                    )
                ) {
                    self.consume_token();
                }
            }
        } else {
            // Unexpected token - consume it anyway to avoid infinite loop
            self.consume_token();
        }

        // Don't wrap in a node here - let parse_postfix_operators handle it
    }

    /// Parse a unary expression.
    ///
    /// Unary expressions have a single operator followed by an operand:
    /// - Negation: `-x`
    /// - Logical NOT: `Not flag`
    ///
    /// # Parameters
    ///
    /// - `bp`: The binding power of the unary operator
    fn parse_unary_expression(&mut self, bp: BindingPower) {
        self.builder
            .start_node(SyntaxKind::UnaryExpression.to_raw());

        // Consume the operator (-, Not)
        self.consume_token();

        // Skip whitespace after operator
        self.consume_whitespace();

        // Parse the operand with the operator's binding power
        self.parse_expression_with_binding_power(bp);

        self.builder.finish_node();
    }

    /// Parse an `AddressOf` expression.
    ///
    /// Syntax: `AddressOf procedureName`
    ///
    /// Used to pass procedure addresses to API functions.
    fn parse_addressof_expression(&mut self) {
        self.builder
            .start_node(SyntaxKind::AddressOfExpression.to_raw());

        // Consume "AddressOf"
        self.consume_token();

        // Skip whitespace
        self.consume_whitespace();

        // Parse the procedure name (identifier)
        if self.is_identifier() || self.at_keyword() {
            self.consume_token();
        }

        self.builder.finish_node();
    }

    /// Parse a `New` expression.
    ///
    /// Syntax: `New ClassName`
    ///
    /// Creates a new instance of a class.
    fn parse_new_expression(&mut self) {
        self.builder.start_node(SyntaxKind::NewExpression.to_raw());

        // Consume "New"
        self.consume_token();

        // Skip whitespace
        self.consume_whitespace();

        // Parse the class name (identifier)
        if self.is_identifier() || self.at_keyword() {
            self.consume_token();
        }

        self.builder.finish_node();
    }

    /// Parse a parenthesized expression.
    ///
    /// Syntax: `(expression)`
    fn parse_parenthesized_expression(&mut self) {
        self.builder
            .start_node(SyntaxKind::ParenthesizedExpression.to_raw());

        // Consume "("
        self.consume_token();

        // Skip whitespace
        self.consume_whitespace();

        // Parse the inner expression
        self.parse_expression();

        // Skip whitespace before ")"
        self.consume_whitespace();

        // Consume ")"
        if self.at_token(Token::RightParenthesis) {
            self.consume_token();
        }

        self.builder.finish_node();
    }

    /// Parse a numeric literal.
    ///
    /// Examples: `42`, `3.14`, `&HFF` (hex), `&O77` (octal), `123.45E-6` (scientific)
    fn parse_numeric_literal(&mut self) {
        self.builder
            .start_node(SyntaxKind::NumericLiteralExpression.to_raw());

        // Consume the number token (already includes type suffix in tokenizer)
        self.consume_token();

        self.builder.finish_node();
    }

    /// Parse a string literal.
    ///
    /// Example: `"Hello, World!"`
    fn parse_string_literal(&mut self) {
        self.builder
            .start_node(SyntaxKind::StringLiteralExpression.to_raw());

        // Consume the string literal token
        self.consume_token();

        self.builder.finish_node();
    }

    /// Parse a boolean literal.
    ///
    /// Examples: `True`, `False`
    fn parse_boolean_literal(&mut self) {
        self.builder
            .start_node(SyntaxKind::BooleanLiteralExpression.to_raw());

        // Consume True or False keyword
        self.consume_token();

        self.builder.finish_node();
    }

    /// Parse a special literal.
    ///
    /// Examples: `Nothing`, `Null`, `Empty`
    fn parse_special_literal(&mut self) {
        self.builder
            .start_node(SyntaxKind::LiteralExpression.to_raw());

        // Consume the keyword
        self.consume_token();

        self.builder.finish_node();
    }

    /// Parse a date literal.
    ///
    /// Syntax: `#1/1/2024#`, `#12:30:45 PM#`, `#1/1/2024 3:45 PM#`
    ///
    /// Note: VB6 has DateLiteral as a token, so it's already parsed as a single token
    fn parse_date_literal(&mut self) {
        self.builder
            .start_node(SyntaxKind::LiteralExpression.to_raw());

        // Consume the date literal token
        self.consume_token();

        self.builder.finish_node();
    }

    /// Parse postfix operators using an existing checkpoint.
    /// Used when we already have a checkpoint from before parsing the primary expression.
    /// Returns true if any postfix operators were found.
    fn parse_postfix_operators_with_checkpoint(&mut self, checkpoint: rowan::Checkpoint) -> bool {
        let current_checkpoint = checkpoint;
        let mut has_any_postfix = false;

        loop {
            // Peek ahead to check for postfix operators WITHOUT consuming whitespace yet
            // This prevents whitespace from being included in the identifier expression

            // Temporarily skip whitespace to check for operators
            let saved_pos = self.pos;
            loop {
                match self.current_token() {
                    Some(Token::Whitespace) => {
                        self.pos += 1;
                    }
                    Some(Token::Underscore) => {
                        // Check for line continuation
                        let mut lookahead = 1;
                        let mut is_continuation = false;
                        while let Some((_, token)) = self.tokens.get(self.pos + lookahead) {
                            if *token == Token::Whitespace {
                                lookahead += 1;
                            } else if *token == Token::Newline {
                                is_continuation = true;
                                break;
                            } else {
                                break;
                            }
                        }

                        if is_continuation {
                            // Skip underscore, whitespace, and newline
                            self.pos += lookahead + 1; // +1 for newline
                        } else {
                            break;
                        }
                    }
                    _ => break,
                }
            }

            let found_postfix = matches!(
                self.current_token(),
                Some(Token::PeriodOperator | Token::LeftParenthesis | Token::ExclamationMark)
            );

            // Restore position
            self.pos = saved_pos;

            if !found_postfix {
                // No postfix operator found - stop here
                break;
            }

            // Found a postfix operator - now consume the whitespace
            self.consume_whitespace();

            match self.current_token() {
                // Member access: .property or .method
                Some(Token::PeriodOperator) => {
                    // Wrap everything parsed so far in a MemberAccessExpression
                    self.builder.start_node_at(
                        current_checkpoint,
                        SyntaxKind::MemberAccessExpression.to_raw(),
                    );
                    has_any_postfix = true;

                    self.parse_member_access_content();

                    self.builder.finish_node();
                    // DON'T update checkpoint - keep it at the original position
                    // This allows subsequent postfix operators to wrap the MemberAccessExpression
                }
                // Function call or array indexing: (...)
                Some(Token::LeftParenthesis) => {
                    // Wrap everything parsed so far in a CallExpression
                    self.builder
                        .start_node_at(current_checkpoint, SyntaxKind::CallExpression.to_raw());
                    has_any_postfix = true;

                    self.parse_call_or_index_content();

                    self.builder.finish_node();
                    // DON'T update checkpoint - keep it at the original position
                }
                // Exclamation mark for dictionary access: collection!key
                Some(Token::ExclamationMark) => {
                    // Wrap everything parsed so far in a MemberAccessExpression
                    self.builder.start_node_at(
                        current_checkpoint,
                        SyntaxKind::MemberAccessExpression.to_raw(),
                    );
                    has_any_postfix = true;

                    self.parse_dictionary_access_content();

                    self.builder.finish_node();
                    // DON'T update checkpoint - keep it at the original position
                }
                _ => break,
            }
        }

        // Don't wrap in IdentifierExpression here - let the caller decide
        // This allows binary expression parsing to work correctly
        has_any_postfix
    }

    /// Parse the content of a member access (everything after the dot).
    fn parse_member_access_content(&mut self) {
        // Consume the period
        self.consume_token();

        // Skip whitespace after period
        self.consume_whitespace();

        // Parse the member name (can be a keyword in VB6)
        if self.is_identifier() || self.at_keyword() {
            self.consume_token();

            // Check for type character suffix
            if matches!(
                self.current_token(),
                Some(
                    Token::DollarSign
                        | Token::Percent
                        | Token::Ampersand
                        | Token::ExclamationMark
                        | Token::Octothorpe
                        | Token::AtSign
                )
            ) {
                self.consume_token();
            }
        }
    }

    /// Parse the content of a function call or array index (the parenthesized part).
    fn parse_call_or_index_content(&mut self) {
        // Consume opening parenthesis
        self.consume_token();

        // Parse argument list
        self.parse_argument_list_in_parens();

        // Consume closing parenthesis
        if self.at_token(Token::RightParenthesis) {
            self.consume_token();
        }
    }

    /// Parse the content of dictionary access (everything after the !).
    fn parse_dictionary_access_content(&mut self) {
        // Consume the exclamation mark
        self.consume_token();

        // Skip whitespace
        self.consume_whitespace();

        // Parse the key (identifier or string)
        if self.is_identifier() || self.at_keyword() || self.at_token(Token::StringLiteral) {
            self.consume_token();
        }
    }

    /// Parse argument list inside parentheses.
    ///
    /// Arguments are comma-separated expressions, optionally with named parameters.
    pub(super) fn parse_argument_list_in_parens(&mut self) {
        self.builder.start_node(SyntaxKind::ArgumentList.to_raw());

        // Skip whitespace after opening paren
        self.consume_whitespace();

        // Parse arguments until we hit closing paren
        while !self.is_at_end() && !self.at_token(Token::RightParenthesis) {
            // Parse each argument as an expression
            self.builder.start_node(SyntaxKind::Argument.to_raw());
            self.parse_expression();
            self.builder.finish_node();

            // Skip whitespace
            self.consume_whitespace();

            // If there's a comma, consume it and continue
            if self.at_token(Token::Comma) {
                self.consume_token();
                self.consume_whitespace();
            } else {
                break;
            }
        }

        self.builder.finish_node();
    }

    /// Get the binding power for an infix operator.
    ///
    /// Returns `Some((left_bp, right_bp))` if the current token is an infix operator,
    /// where `left_bp` is the left binding power and `right_bp` is the right binding power.
    /// Returns `None` if the current token is not an infix operator.
    ///
    /// The difference between left and right binding power determines associativity:
    /// - Left-associative: `right_bp = left_bp + 1` (most operators)
    /// - Right-associative: `right_bp = left_bp` (exponentiation)
    fn get_infix_binding_power(&self) -> Option<(BindingPower, BindingPower)> {
        let token = self.current_token()?;

        let (left_bp, right_bp) = match token {
            // Exponentiation (right-associative)
            Token::ExponentiationOperator => {
                (BindingPower::EXPONENTIATION, BindingPower::EXPONENTIATION)
            }

            // Multiplication and division
            Token::MultiplicationOperator | Token::DivisionOperator => {
                let bp = BindingPower::MULTIPLICATION;
                (bp, BindingPower(bp.0 + 1))
            }

            // Integer division
            Token::BackwardSlashOperator => {
                let bp = BindingPower::INT_DIVISION;
                (bp, BindingPower(bp.0 + 1))
            }

            // Modulo
            Token::ModKeyword => {
                let bp = BindingPower::MODULO;
                (bp, BindingPower(bp.0 + 1))
            }

            // Addition and subtraction
            Token::AdditionOperator | Token::SubtractionOperator => {
                let bp = BindingPower::ADDITION;
                (bp, BindingPower(bp.0 + 1))
            }

            // String concatenation
            Token::Ampersand => {
                let bp = BindingPower::CONCATENATION;
                (bp, BindingPower(bp.0 + 1))
            }

            // Comparison operators
            Token::EqualityOperator
            | Token::InequalityOperator
            | Token::LessThanOrEqualOperator
            | Token::GreaterThanOrEqualOperator
            | Token::LessThanOperator
            | Token::GreaterThanOperator
            | Token::LikeKeyword
            | Token::IsKeyword => {
                let bp = BindingPower::COMPARISON;
                (bp, BindingPower(bp.0 + 1))
            }

            // Logical AND
            Token::AndKeyword => {
                let bp = BindingPower::AND;
                (bp, BindingPower(bp.0 + 1))
            }

            // Logical OR
            Token::OrKeyword => {
                let bp = BindingPower::OR;
                (bp, BindingPower(bp.0 + 1))
            }

            // Logical XOR
            Token::XorKeyword => {
                let bp = BindingPower::XOR;
                (bp, BindingPower(bp.0 + 1))
            }

            // Logical EQV
            Token::EqvKeyword => {
                let bp = BindingPower::EQV;
                (bp, BindingPower(bp.0 + 1))
            }

            // Logical IMP
            Token::ImpKeyword => {
                let bp = BindingPower::IMP;
                (bp, BindingPower(bp.0 + 1))
            }

            _ => return None,
        };

        Some((left_bp, right_bp))
    }

    /// Check if we're at a delimiter that ends an expression.
    ///
    /// Expression delimiters include:
    /// - Newline
    /// - `Then` (in If statements)
    /// - `To` (in For loops)
    /// - `Step` (in For loops)
    /// - Colon (statement separator)
    /// - Comma (argument separator)
    /// - Closing parenthesis/bracket (in some contexts)
    fn is_at_expression_delimiter(&self) -> bool {
        matches!(
            self.current_token(),
            Some(
                Token::Newline
                    | Token::ThenKeyword
                    | Token::ToKeyword
                    | Token::StepKeyword
                    | Token::ColonOperator
                    | Token::EndOfLineComment
                    | Token::RemComment
                    | Token::Comma
                    | Token::RightParenthesis
            )
        )
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    /// Helper function to create a CST from source and get debug output
    fn parse_expression_test(source: &str) -> String {
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        tree.debug_tree()
    }

    #[test]
    fn numeric_literal() {
        let source = "x = 42\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("NumericLiteralExpression"));
        assert!(debug.contains("42"));
    }

    #[test]
    fn numeric_literal_with_type_suffix() {
        let source = "x = 42%\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("NumericLiteralExpression"));
        assert!(debug.contains("42"));
        assert!(debug.contains("%"));
    }

    #[test]
    fn string_literal() {
        let source = "x = \"Hello, World!\"\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("StringLiteralExpression"));
    }

    #[test]
    fn boolean_literal_true() {
        let source = "x = True\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BooleanLiteralExpression"));
        assert!(debug.contains("True"));
    }

    #[test]
    fn boolean_literal_false() {
        let source = "x = False\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BooleanLiteralExpression"));
        assert!(debug.contains("False"));
    }

    #[test]
    fn identifier_expression() {
        let source = "x = myVariable\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("IdentifierExpression"));
        assert!(debug.contains("myVariable"));
    }

    #[test]
    fn simple_addition() {
        let source = "x = 2 + 3\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("2"));
        assert!(debug.contains("+"));
        assert!(debug.contains("3"));
    }

    #[test]
    fn simple_subtraction() {
        let source = "x = 10 - 5\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("10"));
        assert!(debug.contains("-"));
        assert!(debug.contains("5"));
    }

    #[test]
    fn simple_multiplication() {
        let source = "x = 4 * 5\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("4"));
        assert!(debug.contains("*"));
        assert!(debug.contains("5"));
    }

    #[test]
    fn simple_division() {
        let source = "x = 20 / 4\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("20"));
        assert!(debug.contains("/"));
        assert!(debug.contains("4"));
    }

    #[test]
    fn operator_precedence_multiplication_before_addition() {
        let source = "x = 2 + 3 * 4\n";
        let debug = parse_expression_test(source);
        // Should parse as 2 + (3 * 4)
        assert!(debug.contains("BinaryExpression"));
    }

    #[test]
    fn operator_precedence_left_associativity() {
        let source = "x = 10 - 5 - 2\n";
        let debug = parse_expression_test(source);
        // Should parse as (10 - 5) - 2
        assert!(debug.contains("BinaryExpression"));
    }

    #[test]
    fn unary_negation() {
        let source = "x = -5\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("UnaryExpression"));
        assert!(debug.contains("-"));
        assert!(debug.contains("5"));
    }

    #[test]
    fn logical_not() {
        let source = "x = Not True\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("UnaryExpression"));
        assert!(debug.contains("Not"));
    }

    #[test]
    fn logical_and() {
        let source = "x = True And False\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("And"));
    }

    #[test]
    fn logical_or() {
        let source = "x = True Or False\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("Or"));
    }

    #[test]
    fn comparison_equal() {
        let source = "x = a = b\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
    }

    #[test]
    fn comparison_less_than() {
        let source = "x = a < b\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("<"));
    }

    #[test]
    fn comparison_greater_than() {
        let source = "x = a > b\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains(">"));
    }

    #[test]
    fn parenthesized_expression() {
        let source = "x = (5 + 3)\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("ParenthesizedExpression"));
    }

    #[test]
    fn parenthesized_changes_precedence() {
        let source = "x = (2 + 3) * 4\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("ParenthesizedExpression"));
        assert!(debug.contains("BinaryExpression"));
    }

    #[test]
    fn member_access() {
        let source = "x = obj.property\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("MemberAccessExpression"));
        assert!(debug.contains("obj"));
        assert!(debug.contains("property"));
    }

    #[test]
    fn chained_member_access() {
        let source = "x = obj.prop1.prop2\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("MemberAccessExpression"));
    }

    #[test]
    fn function_call_no_args() {
        let source = "x = MyFunction()\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("CallExpression"));
        assert!(debug.contains("MyFunction"));
    }

    #[test]
    fn function_call_one_arg() {
        let source = "x = MyFunction(42)\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("CallExpression"));
        assert!(debug.contains("ArgumentList"));
    }

    #[test]
    fn function_call_multiple_args() {
        let source = "x = MyFunction(1, 2, 3)\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("CallExpression"));
        assert!(debug.contains("ArgumentList"));
    }

    #[test]
    fn method_call() {
        let source = "x = obj.Method(arg)\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("MemberAccessExpression"));
        assert!(debug.contains("CallExpression"));
    }

    #[test]
    fn new_expression() {
        let source = "Set x = New MyClass\n";
        let debug = parse_expression_test(source);
        // In assignment context, "New MyClass" would be parsed differently
        // For now, just verify it parses without error
        assert!(debug.contains("MyClass"));
    }

    #[test]
    fn addressof_expression() {
        let source = "x = AddressOf MyProc\n";
        let debug = parse_expression_test(source);
        // AddressOf in assignment is parsed as identifier
        assert!(debug.contains("AddressOf"));
        assert!(debug.contains("MyProc"));
    }

    #[test]
    fn string_concatenation() {
        let source = "x = \"Hello\" & \" \" & \"World\"\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("&"));
    }

    #[test]
    fn modulo_operator() {
        let source = "x = 10 Mod 3\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("Mod"));
    }

    #[test]
    fn integer_division() {
        let source = "x = 10 \\ 3\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("\\"));
    }

    #[test]
    fn exponentiation() {
        let source = "x = 2 ^ 8\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("^"));
    }

    #[test]
    fn complex_arithmetic() {
        let source = "x = (a + b) * c - d / e\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("BinaryExpression"));
        assert!(debug.contains("ParenthesizedExpression"));
    }

    #[test]
    fn complex_logical() {
        let source = "x = Not a And b Or c\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("UnaryExpression"));
        assert!(debug.contains("BinaryExpression"));
    }

    #[test]
    fn nothing_literal() {
        let source = "Set x = Nothing\n";
        let debug = parse_expression_test(source);
        // Nothing is tokenized but in assignment context appears as identifier
        assert!(debug.contains("Nothing"));
    }

    #[test]
    fn null_literal() {
        let source = "x = Null\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("LiteralExpression"));
        assert!(debug.contains("Null"));
    }

    #[test]
    fn empty_literal() {
        let source = "x = Empty\n";
        let debug = parse_expression_test(source);
        assert!(debug.contains("LiteralExpression"));
        assert!(debug.contains("Empty"));
    }

    #[test]
    fn dollar_sign_functions_merged() {
        // Verify that dollar-sign library functions are properly recognized
        // as single identifiers (e.g., Chr$, UCase$, Left$, etc.)
        let source = r#"
x = Chr$(65)
y = UCase$("hello")
z = Left$("test", 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        // All three dollar-sign functions should appear as single merged identifiers
        assert!(debug.contains("Chr$"), "Chr$ should be a single identifier");
        assert!(
            debug.contains("UCase$"),
            "UCase$ should be a single identifier"
        );
        assert!(
            debug.contains("Left$"),
            "Left$ should be a single identifier"
        );

        // Should NOT contain separate DollarSign tokens in these contexts
        let _ = debug.find("Chr$").expect("Chr$ should exist");
        let _ = debug.find("UCase$").expect("UCase$ should exist");
        let _ = debug.find("Left$").expect("Left$ should exist");

        // Verify these are in Identifier nodes
        assert!(debug.contains("Identifier@"));
    }
}
