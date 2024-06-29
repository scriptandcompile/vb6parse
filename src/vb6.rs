use winnow::{
    token::{one_of, take_till, take_while},
    PResult, Parser,
};

/// Parses a VB6 end-of-line comment.
///
/// The comment starts with a single quote and continues until the end of the line.
/// But it does not include the single quote, the carriage return character, the newline character,
/// and it does not consume the carriage return or newline character.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// The comment without the single quote, carriage return, and newline characters.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::eol_comment_parse;
///
/// let mut input = "' This is a comment\r\n".as_bytes();
/// let comment = eol_comment_parse(&mut input).unwrap();
///
/// assert_eq!(comment, b" This is a comment");
/// ```
pub fn eol_comment_parse<'a>(input: &mut &'a [u8]) -> PResult<&'a [u8]> {
    '\''.parse_next(input)?;

    let comment = take_till(0.., ('\r', '\n')).parse_next(input)?;

    Ok(comment)
}

/// Parses a VB6 variable name.
///
/// The variable name starts with a letter and can contain letters, numbers, and underscores.
///
/// # Arguments
///
/// * `input` - The input to parse.
///
/// # Returns
///
/// The VB6 variable name.
///
/// # Example
///
/// ```rust
/// use vb6parse::vb6::variable_name_parse;
///
/// let mut input = "variable_name".as_bytes();
/// let variable_name = variable_name_parse(&mut input).unwrap();
///
/// assert_eq!(variable_name, b"variable_name");
/// ```
pub fn variable_name_parse<'a>(input: &mut &'a [u8]) -> PResult<&'a [u8]> {
    let variable_name = (
        one_of(('a'..='z', 'A'..='Z')),
        take_while(1.., ('_', 'a'..='z', 'A'..='Z', '0'..='9')),
    )
        .recognize()
        .parse_next(input)?;

    Ok(variable_name)
}
