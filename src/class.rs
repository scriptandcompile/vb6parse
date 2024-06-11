#![warn(clippy::pedantic)]

use std::{mem::take, string::ParseError};

use bstr::ByteSlice;
use winnow::{
    ascii::line_ending, combinator::{alt, rest, separated_pair}, error::{ContextError, ErrMode, ParserError}, stream::{AsBStr, AsChar}, token::{literal, take_till, take_while}, PResult, Parser
};

use crate::errors::VB6ClassParseError;

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassFile<'a> {
    pub version: VB6ClassVersion,
    pub multi_use: bool,                    // (0/-1) false/true
    pub persistable: bool,                  // (0/-1) false/true
    pub data_binding_behavior: bool,        // (0/-1) false/true - vbNone
    pub data_source_behavior: bool,         // (0/-1) false/true - vbNone
    pub mts_transaction_mode: bool,         // (0/-1) false/true - If false then this is NotAnMTSObject / if true this is a MTSObject
    pub attributes: VB6FileAttributes<'a>,
    pub option_explicit: bool               // Option Explicit, or Option Explicit ON, or Option Explicit OFF. This must be before any other line in the file (except comments)
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ClassVersion {
    pub major: u16,
    pub minor: u16,
}

#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6FileAttributes<'a> {
    pub name: &'a [u8],             // Attribute VB_Name = "Organism"
    pub global_name_space: bool,    // (True/False) Attribute VB_GlobalNameSpace = False
    pub creatable: bool,            // (True/False) Attribute VB_Creatable = True
    pub pre_declared_id: bool,      // (True/False) Attribute VB_PredeclaredId = False
    pub exposed: bool,              // (True/False) Attribute VB_Exposed = False
}

impl<'a> VB6ClassFile<'a> {
    pub fn parse(input: &'a [u8]) -> Result<Self, VB6ClassParseError<&'a [u8]>> {
        
        let mut input = input;

        //let version = literal(b"VERSION").parse_next(&mut input) else {
            //panic!("not working yet. I've got to figure this thing out.");
            //let r = literal.err();

            //let error = VB6ClassParseError::ClassVersionInformationNotFound {
                //line: line.unwrap(),
                //line_location: 0
            //};
            
            //return Err(error);
        //};

        //
        Ok(VB6ClassFile {
            version: VB6ClassVersion {
                major: 0,
                minor: 0
            },
            multi_use: false,
            persistable: false,
            data_binding_behavior: false,
            data_source_behavior: false,
            mts_transaction_mode: false,
            attributes: VB6FileAttributes {
                name: b"",
                global_name_space: false,
                creatable: false,
                pre_declared_id: false,
                exposed: false
            },
            option_explicit: false
        })
    }
}

// fn true_false_parse<'a>(input: &mut &'a [u8]) -> PResult<bool, VB6ClassParseError<'a>> {
//     // 0 is false...and -1 is true.
//     // Why vb6? What are you like this? Who hurt you?
//     let result = alt(('0'.value(false), "-1".value(true))).parse_next(input)?;

//     Ok(result)
// }