#![warn(clippy::pedantic)]

use core::panic;

use crate::vb6::{eol_comment_parse, keyword_parse, vb6_parse, VB6Token};

use winnow::{
    ascii::{digit1, line_ending, space0, space1},
    combinator::{opt, repeat_till},
    error::{ContextError, ErrMode, ParserError, StrContext},
    token::{literal, take_while},
    PResult, Parser,
};

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum FileUsage {
    MultiUse,  // -1 (true)
    SingleUse, // 0 (false)
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum Persistance {
    Persistable,    // -1 (true)
    NonPersistable, // 0 (false)
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub enum MtsStatus {
    NotAnMTSObject, // 0 (false)
    MTSObject,      // -1 (true)
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassHeader<'a> {
    pub version: VB6ClassVersion,
    pub multi_use: FileUsage,            // (0/-1) multi use / single use
    pub persistable: Persistance,        // (0/-1) NonParsistable / Persistable
    pub data_binding_behavior: bool,     // (0/-1) false/true - vbNone
    pub data_source_behavior: bool,      // (0/-1) false/true - vbNone
    pub mts_transaction_mode: MtsStatus, // (0/-1) NotAnMTSObject / MTSObject
    pub attributes: VB6FileAttributes<'a>,
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
    pub tokens: Vec<VB6Token<'a>>,
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

        let tokens = vb6_parse(&mut input)?;

        Ok(VB6ClassFile { header, tokens })
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

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistance::NonPersistable;
    let mut data_binding_behavior = false;
    let mut data_source_behavior = false;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    let (collection, _): (Vec<(&[u8], &[u8])>, &[u8]) =
        repeat_till(0.., key_value_line_parse(b"="), keyword_parse("END")).parse_next(input)?;

    for pair in &collection {
        let (key, value) = *pair;

        match key {
            b"Persistable" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == b"-1" {
                    persistable = Persistance::Persistable;
                } else {
                    persistable = Persistance::NonPersistable;
                }
            }
            b"MultiUse" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == b"-1" {
                    multi_use = FileUsage::MultiUse;
                } else {
                    multi_use = FileUsage::SingleUse;
                }
            }
            b"DataBindingBehavior" => {
                // -1 is 'true' and 0 is 'false' in VB6
                data_binding_behavior = value == b"-1";
            }
            b"DataSourceBehavior" => {
                // -1 is 'true' and 0 is 'false' in VB6
                data_source_behavior = value == b"-1";
            }
            b"MTSTransactionMode" => {
                // -1 is 'true' and 0 is 'false' in VB6
                if value == b"-1" {
                    mts_transaction_mode = MtsStatus::MTSObject;
                } else {
                    mts_transaction_mode = MtsStatus::NotAnMTSObject;
                }
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

fn version_parse(input: &mut &[u8]) -> PResult<VB6ClassVersion> {
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
