#![warn(clippy::pedantic)]

use core::panic;

use crate::vb6::eol_comment_parse;

use winnow::{
    ascii::{digit1, line_ending, space0, space1, Caseless},
    combinator::{opt, repeat_till},
    error::{ContextError, ErrMode, ParserError, StrContext, StrContextValue},
    token::{literal, take_while},
    PResult, Parser,
};

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassHeader<'a> {
    pub version: VB6ClassVersion,
    pub multi_use: bool,             // (0/-1) false/true
    pub persistable: bool,           // (0/-1) false/true
    pub data_binding_behavior: bool, // (0/-1) false/true - vbNone
    pub data_source_behavior: bool,  // (0/-1) false/true - vbNone
    pub mts_transaction_mode: bool, // (0/-1) false/true - If false then this is NotAnMTSObject / if true this is a MTSObject
    pub attributes: VB6FileAttributes<'a>,
    pub option_explicit: bool, // Option Explicit, or Option Explicit ON, or Option Explicit OFF. This must be before any other line in the file (except comments)
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassAttributes<'a> {
    pub name: &'a [u8],          // Attribute VB_Name = "Organism"
    pub global_name_space: bool, // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: bool,         // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: bool,   // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: bool,           // (True/False) Attribute VB_Exposed = False
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassFile<'a> {
    pub header: VB6ClassHeader<'a>,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassVersion {
    pub major: u16,
    pub minor: u16,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FileAttributes<'a> {
    pub name: &'a [u8],          // Attribute VB_Name = "Organism"
    pub global_name_space: bool, // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: bool,         // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: bool,   // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: bool,           // (True/False) Attribute VB_Exposed = False
}

impl<'a> VB6ClassFile<'a> {
    pub fn parse(input: &'a [u8]) -> Result<Self, ErrMode<ContextError>> {
        let mut input = input;

        let header = class_header_parse(&mut input)?;

        Ok(VB6ClassFile { header })
    }
}

fn class_header_parse<'a>(input: &mut &'a [u8]) -> PResult<VB6ClassHeader<'a>> {
    // VERSION #.# CLASS
    // BEGIN
    //  key = value 'comment
    //  ...
    // END

    let version = version_parse.parse_next(input)?;

    keyword_parse("BEGIN").parse_next(input)?;

    line_ending
        .context(StrContext::Label("Newline expected after BEGIN keyword."))
        .parse_next(input)?;

    let mut multi_use = false;
    let mut persistable = false;
    let mut data_binding_behavior = false;
    let mut data_source_behavior = false;
    let mut mts_transaction_mode = false;

    let (collection, _): (Vec<(&[u8], &[u8])>, &[u8]) =
        repeat_till(0.., key_value_line_parse(b"="), keyword_parse("END")).parse_next(input)?;

    for pair in collection.iter() {
        let (key, value) = *pair;

        match key {
            b"Persistable" => {
                persistable = value == b"-1";
            }
            b"MultiUse" => {
                multi_use = value == b"-1";
            }
            b"DataBindingBehavior" => {
                data_binding_behavior = value == b"-1";
            }
            b"DataSourceBehavior" => {
                data_source_behavior = value == b"-1";
            }
            b"MTSTransactionMode" => {
                mts_transaction_mode = value == b"-1";
            }
            _ => {
                panic!("Unknown key found in class header.");
            }
        }
    }

    line_ending
        .context(StrContext::Label("Newline expected after END."))
        .parse_next(input)?;

    let attributes = attributes_parse.parse_next(input)?;

    Ok(VB6ClassHeader {
        version,
        multi_use,
        persistable,
        data_binding_behavior,
        data_source_behavior,
        mts_transaction_mode,
        attributes,
        option_explicit: false,
    })
}

fn attributes_parse<'a>(input: &mut &'a [u8]) -> PResult<VB6FileAttributes<'a>> {
    let _ = space0::<&[u8], ContextError>(input);

    let mut name = Option::None;
    let mut global_name_space = false;
    let mut creatable = false;
    let mut pre_declared_id = false;
    let mut exposed = false;

    while let Ok((_, (key, value))) =
        (keyword_parse("Attribute"), key_value_parse(b"=")).parse_next(input)
    {
        line_ending
            .context(StrContext::Label(
                "Newline expected after Class File Attribute line.",
            ))
            .parse_next(input)?;

        match key {
            b"VB_Name" => {
                name = Some(value);
            }
            b"VB_GlobalNameSpace" => {
                global_name_space = value == b"True";
            }
            b"VB_Creatable" => {
                creatable = value == b"True";
            }
            b"VB_PredeclaredId" => {
                pre_declared_id = value == b"True";
            }
            b"VB_Exposed" => {
                exposed = value == b"True";
            }
            _ => {
                panic!("Unknown key found in class attributes.");
            }
        }
    }

    if name.is_none() {
        let error = ParserError::assert(input, "VB_Name attribute not found.");

        return Err(ErrMode::Cut(error));
    }

    Ok(VB6FileAttributes {
        name: name.unwrap(),
        global_name_space,
        creatable,
        pre_declared_id,
        exposed,
    })
}

fn keyword_parse<'a>(keyword: &'a str) -> impl Parser<&'a [u8], &'a [u8], ContextError> {
    move |input: &mut &'a [u8]| {
        let keyword_value = Caseless(keyword)
            .context(StrContext::Expected(StrContextValue::Description(
                "Unable to match Keyword {keyword}",
            )))
            .parse_next(input)?;

        Ok(keyword_value)
    }
}

fn key_value_parse<'a>(
    divider: &'a [u8],
) -> impl Parser<&'a [u8], (&'a [u8], &'a [u8]), ContextError> {
    move |input: &mut &'a [u8]| {
        let _ = space0::<&[u8], ContextError>.parse_next(input);

        let key = take_while(1.., ('_', '"', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9'))
            .parse_next(input)?;

        let _ = space0::<&[u8], ContextError>.parse_next(input);

        literal(divider).parse_next(input)?;

        let _ = space0::<&[u8], ContextError>.parse_next(input);

        opt("\"").parse_next(input)?;

        let value =
            take_while(1.., ('_', '-', '+', 'a'..='z', 'A'..='Z', '0'..='9')).parse_next(input)?;

        opt("\"").parse_next(input)?;

        let _ = space0::<&[u8], ContextError>.parse_next(input);

        Ok((key, value))
    }
}

fn key_value_line_parse<'a>(
    divider: &'a [u8],
) -> impl Parser<&'a [u8], (&'a [u8], &'a [u8]), ContextError> {
    move |input: &mut &'a [u8]| {
        let (key, value) = key_value_parse(&divider).parse_next(input)?;

        eol_comment_parse.parse_next(input)?;

        line_ending
            .context(StrContext::Label("newline not found"))
            .parse_next(input)?;

        Ok((key, value))
    }
}

fn version_parse<'a>(input: &mut &'a [u8]) -> PResult<VB6ClassVersion> {
    keyword_parse("VERSION").parse_next(input)?;

    space1(input)?;

    let major_digits = digit1
        .context(StrContext::Label("Major version number not found."))
        .parse_next(input)?;

    let major_version =
        u16::from_str_radix(bstr::BStr::new(major_digits).to_string().as_str(), 10).unwrap();

    b".".context(StrContext::Label("Version decimal character not found."))
        .parse_next(input)?;

    let minor_digits = digit1
        .context(StrContext::Label("Minor version number not found."))
        .parse_next(input)?;

    let minor_version =
        u16::from_str_radix(bstr::BStr::new(minor_digits).to_string().as_str(), 10).unwrap();

    space1(input)?;

    keyword_parse("CLASS").parse_next(input)?;

    space0(input)?;

    opt(line_ending.context(StrContext::Label("Newline expected after CLASS keyword.")))
        .parse_next(input)?;

    Ok(VB6ClassVersion {
        major: major_version,
        minor: minor_version,
    })
}
