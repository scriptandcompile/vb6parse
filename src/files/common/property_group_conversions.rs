//! This module provides standard Rust trait implementations for converting between
//! `PropertyGroup` and other types.
//!
//! Uses `TryFrom` for fallible conversions from `PropertyGroup` and `From` for
//! infallible conversions to `PropertyGroup`.

use crate::files::common::PropertyGroup;
use crate::language::Font;

use std::collections::HashMap;

use either::Either;

/// Fallible conversion from `PropertyGroup` to `Font`
impl TryFrom<&PropertyGroup> for Font {
    type Error = String; // In a real implementation, you would likely want a more specific error type

    fn try_from(group: &PropertyGroup) -> Result<Self, Self::Error> {
        if !group.name.eq_ignore_ascii_case("Font") {
            return Err(format!(
                "Expected PropertyGroup name 'Font', found '{}'",
                group.name
            ));
        }

        let name = group
            .properties
            .get("Name")
            .and_then(|name| name.as_ref().left())
            .ok_or_else(|| "Missing 'Name' property".to_string())?;

        let size = group
            .properties
            .get("Size")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Size' property".to_string())?
            .parse::<f32>()
            .map_err(|_| "Invalid 'Size' property value".to_string())?;

        let charset = group
            .properties
            .get("Charset")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Charset' property".to_string())?
            .parse::<i32>()
            .map_err(|_| "Invalid 'Charset' property value".to_string())?;

        let weight = group
            .properties
            .get("Weight")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Weight' property".to_string())?
            .parse::<i32>()
            .map_err(|_| "Invalid 'Weight' property value".to_string())?;

        let underline = group
            .properties
            .get("Underline")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Underline' property".to_string())?;

        let italic = group
            .properties
            .get("Italic")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Italic' property".to_string())?;

        let strikethrough = group
            .properties
            .get("Strikethrough")
            .and_then(|v| v.as_ref().left())
            .ok_or_else(|| "Missing 'Strikethrough' property".to_string())?;

        Ok(Font {
            name: name.clone(),
            size,
            charset,
            weight,
            underline: parse_vb6_bool(underline),
            italic: parse_vb6_bool(italic),
            strikethrough: parse_vb6_bool(strikethrough),
        })
    }
}

/// Infallible conversion from `Font` to `PropertyGroup`
impl From<&Font> for PropertyGroup {
    fn from(font: &Font) -> Self {
        let mut properties = HashMap::new();

        properties.insert("Name".to_string(), Either::Left(font.name.clone()));
        properties.insert("Size".to_string(), Either::Left(font.size.to_string()));
        properties.insert(
            "Charset".to_string(),
            Either::Left(font.charset.to_string()),
        );
        properties.insert("Weight".to_string(), Either::Left(font.weight.to_string()));
        properties.insert(
            "Underline".to_string(),
            Either::Left(if font.underline { "-1" } else { "0" }.to_string()),
        );
        properties.insert(
            "Italic".to_string(),
            Either::Left(if font.italic { "-1" } else { "0" }.to_string()),
        );
        properties.insert(
            "Strikethrough".to_string(),
            Either::Left(if font.strikethrough { "-1" } else { "0" }.to_string()),
        );

        PropertyGroup {
            name: "Font".to_string(),
            guid: None, // Font can have GUID in VB6, could extract if needed
            properties,
        }
    }
}

// Helper function
fn parse_vb6_bool(s: &str) -> bool {
    matches!(s, "-1" | "True" | "true")
}
