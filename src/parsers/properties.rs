use std::borrow::Cow;
use std::collections::HashMap;
use std::convert::TryFrom;
use std::iter::Iterator;

use bstr::BStr;
use bstr::ByteSlice;
use num_enum::TryFromPrimitive;

use crate::language::StartUpPosition;
use crate::language::VB6Color;

pub struct Properties<'a> {
    key_value_store: HashMap<&'a BStr, Cow<'a, [u8]>>,
}

pub struct PropertiesIter<'a> {
    iter: std::collections::hash_map::Iter<'a, &'a BStr, Cow<'a, [u8]>>,
}

impl<'a> Properties<'a> {
    pub fn iter(&self) -> PropertiesIter<'_> {
        PropertiesIter {
            iter: self.key_value_store.iter(),
        }
    }
}

impl<'a> Iterator for PropertiesIter<'a> {
    type Item = (&'a BStr, &'a [u8]);

    fn next(&mut self) -> Option<Self::Item> {
        self.iter.next().map(|(key, value)| (*key, value.as_ref()))
    }
}

impl<'a> Clone for Properties<'a> {
    fn clone(&self) -> Self {
        Properties {
            key_value_store: self.key_value_store.clone(),
        }
    }
}

impl<'a> Default for Properties<'a> {
    fn default() -> Self {
        Properties::new()
    }
}

impl<'a> Properties<'a> {
    pub fn new() -> Self {
        Properties {
            key_value_store: HashMap::new(),
        }
    }

    pub fn insert(&mut self, property_key: &'a BStr, value: &'a [u8]) {
        self.key_value_store
            .insert(property_key, Cow::Borrowed(value));
    }

    pub fn insert_resource(&mut self, property_key: &'a BStr, value: Vec<u8>) {
        self.key_value_store.insert(property_key, Cow::Owned(value));
    }

    pub fn len(&self) -> usize {
        self.key_value_store.len()
    }

    pub fn is_empty(&self) -> bool {
        self.key_value_store.is_empty()
    }

    pub fn contains_key(&self, property_key: &BStr) -> bool {
        self.key_value_store.contains_key(property_key)
    }

    pub fn get_keys(&self) -> Vec<&'a BStr> {
        self.key_value_store.keys().copied().collect()
    }

    pub fn remove(&mut self, property_key: &BStr) -> Option<Cow<'a, [u8]>> {
        self.key_value_store.remove(property_key)
    }

    pub fn clear(&mut self) {
        self.key_value_store.clear();
    }

    #[must_use]
    pub fn get(&self, property_key: &BStr) -> Option<&[u8]> {
        if !self.key_value_store.contains_key(property_key) {
            return None;
        }

        Some(&self.key_value_store[property_key])
    }

    #[must_use]
    pub fn get_bool(&self, property_key: &BStr, default: bool) -> bool {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        match &self.key_value_store[property_key] {
            Cow::Borrowed(b) => match b {
                &b"0" => false,
                &b"1" | &b"-1" => true,
                _ => default,
            },
            Cow::Owned(b) => match b.as_slice() {
                b"0" => false,
                b"1" | b"-1" => true,
                _ => default,
            },
        }
    }

    #[must_use]
    pub fn get_color(&self, property_key: &BStr, default: VB6Color) -> VB6Color {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        let Ok(property_ascii) = self.key_value_store[property_key].to_str() else {
            // If conversion fails, return default color
            return default;
        };

        match VB6Color::from_hex(property_ascii) {
            Ok(color) => color,
            Err(_) => default,
        }
    }

    #[must_use]
    pub fn get_i32(&self, property_key: &BStr, default: i32) -> i32 {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        let Ok(property_ascii) = self.key_value_store[property_key].to_str() else {
            // If conversion fails, return default value
            return default;
        };

        match property_ascii.parse::<i32>() {
            Ok(value) => value,
            Err(_) => default,
        }
    }

    #[must_use]
    pub fn get_property<T>(&self, property_key: &BStr, default: T) -> T
    where
        T: TryFrom<i32> + TryFromPrimitive,
    {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        let Ok(property_ascii) = self.key_value_store[property_key].to_str() else {
            // If conversion fails, return default value
            return default;
        };

        match property_ascii.parse::<i32>() {
            Ok(value) => T::try_from(value).unwrap_or(default),
            Err(_) => default,
        }
    }

    #[must_use]
    pub fn get_startup_position(
        &self,
        property_key: &BStr,
        default: StartUpPosition,
    ) -> StartUpPosition {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        let Ok(property_ascii) = self.key_value_store[property_key].to_str() else {
            // If conversion fails, return default position
            return default;
        };

        match property_ascii.parse::<i32>() {
            Ok(value) => match value {
                0 => {
                    let client_height = self.get_i32(b"ClientHeight".into(), 3000);
                    let client_width = self.get_i32(b"ClientWidth".into(), 3000);
                    let client_top = self.get_i32(b"ClientTop".into(), 200);
                    let client_left = self.get_i32(b"ClientLeft".into(), 100);

                    StartUpPosition::Manual {
                        client_height,
                        client_width,
                        client_top,
                        client_left,
                    }
                }
                1 => StartUpPosition::CenterOwner,
                2 => StartUpPosition::CenterScreen,
                // 3 is the default value for Windows, but we also want
                // to default to WindowsDefault if the value is not found.
                // I've just commented this out since leaving it in will
                // cause Clippy to complain about the unknown value and the
                // default value in the match arm being the same.
                //
                // 3 => StartUpPosition::WindowsDefault,
                _ => StartUpPosition::WindowsDefault,
            },
            Err(_) => StartUpPosition::WindowsDefault,
        }
    }

    #[must_use]
    pub fn get_option<T>(&'a self, property_key: &'a BStr, default: Option<T>) -> Option<T>
    where
        T: TryFrom<&'a str>,
    {
        if !self.key_value_store.contains_key(property_key) {
            return default;
        }

        let property_ascii = &self.key_value_store[property_key];

        let property_ascii = match property_ascii {
            Cow::Borrowed(b) => match b.to_str() {
                Ok(value) => value,
                Err(_) => return default,
            },
            Cow::Owned(value) => match value.to_str() {
                Ok(value) => value,
                Err(_) => return default,
            },
        };

        match T::try_from(property_ascii) {
            Ok(value) => Some(value),
            Err(_) => default,
        }
    }
}
