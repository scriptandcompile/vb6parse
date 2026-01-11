//! Module for parsing and handling VB6 property groups and key-value properties.
//!
//! This module defines structures and methods to represent and manipulate
//! property groups and key-value properties typically found in VB6 project files.
//!
//! It includes functionality for serialization, type conversion, and iteration over properties.
//!
//! # Examples
//! ```rust
//! use vb6parse::files::common::{PropertyGroup, Properties};
//! use vb6parse::language::Color;
//! use vb6parse::language::color::VB_RED;
//! use std::collections::HashMap;
//! use either::Either;
//! use uuid::Uuid;
//! let mut properties = HashMap::new();
//! properties.insert("BackColor".to_string(), Either::Left(VB_RED.to_vb_string()));
//! let group = PropertyGroup {
//!     name: "FormProperties".to_string(),
//!     guid: Some(Uuid::parse_str("123e4567-e89b-12d3-a456-426614174000").unwrap()),
//!     properties,
//! };
//! assert_eq!(group.name, "FormProperties");
//! assert_eq!(group.guid.unwrap().to_string(), "123e4567-e89b-12d3-a456-426614174000");
//! assert_eq!(group.properties.get("BackColor").unwrap(), &Either::Left(VB_RED.to_vb_string()));
//! ```
//!

use std::collections::HashMap;
use std::convert::TryFrom;
use std::iter::Iterator;

use either::Either;
use num_enum::TryFromPrimitive;
use serde::Serialize;
use uuid::Uuid;

use crate::language::Color;
use crate::language::StartUpPosition;

/// A group of properties, which may contain nested property groups.
///
/// # Examples
///
/// ```rust
/// # pub fn main() -> Result<(), Box<dyn std::error::Error>> {
///     use vb6parse::files::common::PropertyGroup;
///     use vb6parse::language::Color;
///     use vb6parse::language::color::VB_RED;
///     use std::collections::HashMap;
///     use either::Either;
///     use uuid::Uuid;
///
///     let mut properties = HashMap::new();
///
///     properties.insert("BackColor".to_string(), Either::Left(VB_RED.to_vb_string()));
///
///     let group = PropertyGroup {
///         name: "FormProperties".to_string(),
///         guid: Some(Uuid::parse_str("123e4567-e89b-12d3-a456-426614174000")?),
///         properties,
///     };
///
///     assert_eq!(group.name, "FormProperties");
///     assert_eq!(group.guid.expect("Expected GUID").to_string(), "123e4567-e89b-12d3-a456-426614174000");
///     assert_eq!(group.properties.get("BackColor").expect("Expected 'BackColor'property"), &Either::Left(VB_RED.to_vb_string()));
///     # Ok(())
/// # }
/// ```
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct PropertyGroup {
    /// The name of the property group.
    pub name: String,
    /// An optional GUID associated with the property group.
    pub guid: Option<Uuid>,
    /// A map of property names to their values or nested property groups.
    pub properties: HashMap<String, Either<String, PropertyGroup>>,
}

/// Serialize implementation for `PropertyGroup`.
impl Serialize for PropertyGroup {
    /// Serializes the `PropertyGroup` into a structured format.
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::Serializer,
    {
        use serde::ser::SerializeStruct;

        let mut state = serializer.serialize_struct("VB6PropertyGroup", 3)?;

        state.serialize_field("name", &self.name)?;

        if let Some(guid) = &self.guid {
            state.serialize_field("guid", &guid.to_string())?;
        } else {
            state.serialize_field("guid", &"None")?;
        }

        state.serialize_field("properties", &self.properties)?;

        state.end()
    }
}

/// A collection of key-value properties typically found in VB6 project files.
/// The keys and values are stored as strings, but utility methods are provided
/// to retrieve values in various types such as `bool`, `i32`, `Color`, and enums.
///
/// This is a thin wrapper around a `HashMap<String, String>` with added convenience methods.
///
/// # Examples
///
/// ```rust
/// use vb6parse::files::common::Properties;
/// let mut props = Properties::new();
/// props.insert("ClientWidth", "800");
/// let width = props.get_i32("ClientWidth", 600);
/// assert_eq!(width, 800);
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Properties {
    key_value_store: HashMap<String, String>,
}

impl AsRef<HashMap<String, String>> for Properties {
    fn as_ref(&self) -> &HashMap<String, String> {
        &self.key_value_store
    }
}

/// An iterator over the key-value pairs in a `Properties` collection.
///
/// # Examples
///
/// ```rust
/// use vb6parse::files::common::Properties;
/// let mut props = Properties::new();
/// props.insert("Key1", "Value1");
/// props.insert("Key2", "Value2");
/// for (key, value) in props.iter() {
///    println!("{}: {}", key, value);
/// }
/// ```
pub struct PropertiesIter<'a> {
    iter: std::collections::hash_map::Iter<'a, String, String>,
}

/// Iterator implementation for `PropertiesIter`.
///
/// # Examples
///
/// ```rust
/// use vb6parse::files::common::Properties;
/// let mut props = Properties::new();
/// props.insert("Key1", "Value1");
/// props.insert("Key2", "Value2");
/// let mut iter = props.iter();
/// while let Some((key, value)) = iter.next() {
///    println!("{}: {}", key, value);
/// }
/// ```
impl Properties {
    /// Returns an iterator over the key-value pairs in the `Properties` collection.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// props.insert("Key2", "Value2");
    /// let mut iter = props.iter();
    /// while let Some((key, value)) = iter.next() {
    ///    println!("{}: {}", key, value);
    /// }
    /// ```
    #[must_use]
    pub fn iter(&self) -> PropertiesIter<'_> {
        PropertiesIter {
            iter: self.key_value_store.iter(),
        }
    }
}

impl<'a> IntoIterator for &'a Properties {
    type Item = (&'a String, &'a String);
    type IntoIter = PropertiesIter<'a>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

/// Iterator implementation for `PropertiesIter`.
///
/// # Examples
///
/// ```rust
/// use vb6parse::files::common::Properties;
/// let mut props = Properties::new();
/// props.insert("Key1", "Value1");
/// props.insert("Key2", "Value2");
/// let mut iter = props.iter();
/// while let Some((key, value)) = iter.next() {
///    println!("{}: {}", key, value);
/// }
/// ```
impl<'a> Iterator for PropertiesIter<'a> {
    type Item = (&'a String, &'a String);

    /// Returns the next key-value pair in the iterator.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// let mut iter = props.iter();
    /// if let Some((key, value)) = iter.next() {
    ///    assert_eq!(key, "Key1");
    ///    assert_eq!(value, "Value1");
    /// }
    /// ```
    fn next(&mut self) -> Option<Self::Item> {
        self.iter.next()
    }
}

impl<K, V> FromIterator<(K, V)> for Properties
where
    K: Into<String>,
    V: Into<String>,
{
    fn from_iter<T: IntoIterator<Item = (K, V)>>(iter: T) -> Self {
        let mut props = Properties::new();
        for (k, v) in iter {
            props.insert(k, v);
        }
        props
    }
}

impl<K, V> Extend<(K, V)> for Properties
where
    K: Into<String>,
    V: Into<String>,
{
    fn extend<T: IntoIterator<Item = (K, V)>>(&mut self, iter: T) {
        for (k, v) in iter {
            self.insert(k, v);
        }
    }
}

impl Properties {
    /// Creates a new, empty `Properties` collection.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let props = Properties::new();
    /// assert!(props.is_empty());
    /// ```
    #[must_use]
    pub fn new() -> Self {
        Properties {
            key_value_store: HashMap::new(),
        }
    }

    /// Inserts a key-value pair into the `Properties` collection.
    /// If the key already exists, its value is updated.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// assert_eq!(props.get("Key1"), Some(&"Value1".to_string()));
    /// ```
    pub fn insert(&mut self, property_key: impl Into<String>, value: impl Into<String>) {
        self.key_value_store
            .insert(property_key.into(), value.into());
    }

    /// Returns the number of key-value pairs in the `Properties` collection.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// assert_eq!(props.len(), 0);
    /// props.insert("Key1", "Value1");
    /// assert_eq!(props.len(), 1);
    /// ```
    #[must_use]
    pub fn len(&self) -> usize {
        self.key_value_store.len()
    }

    /// Checks if the `Properties` collection is empty.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// assert!(props.is_empty());
    /// props.insert("Key1", "Value1");
    /// assert!(!props.is_empty());
    /// ```
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.key_value_store.is_empty()
    }

    /// Checks if the `Properties` collection contains the specified key.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// assert!(props.contains_key("Key1"));
    /// assert!(!props.contains_key("Key2"));
    /// ```
    #[must_use]
    pub fn contains_key(&self, property_key: &str) -> bool {
        self.key_value_store.contains_key(property_key)
    }

    /// Returns a vector of all keys in the `Properties` collection.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// props.insert("Key2", "Value2");
    /// let keys = props.keys();
    /// assert_eq!(keys.len(), 2);
    /// assert!(keys.contains(&&"Key1".to_string()));
    /// assert!(keys.contains(&&"Key2".to_string()));
    /// ```
    #[must_use]
    pub fn keys(&self) -> Vec<&String> {
        self.key_value_store.keys().collect()
    }

    /// Removes a key-value pair from the `Properties` collection by key.
    /// Returns the removed value if the key existed.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// let removed = props.remove("Key1");
    /// assert_eq!(removed, Some("Value1".to_string()));
    /// assert!(!props.contains_key("Key1"));
    /// ```
    pub fn remove(&mut self, property_key: &str) -> Option<String> {
        self.key_value_store.remove(property_key)
    }

    /// Clears all key-value pairs from the `Properties` collection.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// props.clear();
    /// assert!(props.is_empty());
    /// ```
    pub fn clear(&mut self) {
        self.key_value_store.clear();
    }

    /// Retrieves the value associated with the specified key as a string slice.
    /// Returns `None` if the key does not exist.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("Key1", "Value1");
    /// let value = props.get("Key1");
    /// assert_eq!(value, Some(&"Value1".to_string()));
    /// let missing = props.get("Key2");
    /// assert_eq!(missing, None);
    /// ```
    #[must_use]
    pub fn get(&self, property_key: &str) -> Option<&String> {
        self.key_value_store.get(property_key)
    }

    /// Retrieves the value associated with the specified key as a boolean.
    /// Interprets "0" as `false`, "1" or "-1" as `true`.
    /// Returns the provided default value if the key does not exist or cannot be parsed.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("IsEnabled", "1");
    /// let is_enabled = props.get_bool("IsEnabled", false);
    /// assert_eq!(is_enabled, true);
    /// let is_disabled = props.get_bool("IsDisabled", true);
    /// assert_eq!(is_disabled, true); // default used
    /// ```
    #[must_use]
    pub fn get_bool(&self, property_key: &str, default: bool) -> bool {
        match self.key_value_store.get(property_key) {
            Some(value) => match value.as_str() {
                "0" => false,
                "1" | "-1" => true,
                _ => default,
            },
            None => default,
        }
    }

    /// Retrieves the value associated with the specified key as a `Color`.
    /// Parses the value as a hexadecimal color string.
    /// Returns the provided default color if the key does not exist or cannot be parsed.
    ///
    /// # Examples
    ///
    /// ```rust
    /// # pub fn main() -> Result<(), Box<dyn std::error::Error>> {
    ///     use vb6parse::files::common::Properties;
    ///     use vb6parse::language::color::VB_WHITE;
    ///     use vb6parse::language::Color;
    ///
    ///     let mut props = Properties::new();
    ///     props.insert("BackgroundColor", "&H00FFFFFF&");
    ///
    ///     let color = props.get_color("BackgroundColor", VB_WHITE);
    ///     assert_eq!(color, Color::from_hex("&H00FFFFFF&")?);
    ///
    ///     let default_color = props.get_color("ForegroundColor", VB_WHITE);
    ///     assert_eq!(default_color, VB_WHITE); // default used
    ///     # Ok(())
    /// # }
    /// ```
    #[must_use]
    pub fn get_color(&self, property_key: &str, default: Color) -> Color {
        match self.key_value_store.get(property_key) {
            Some(value) => Color::from_hex(value).unwrap_or(default),
            None => default,
        }
    }

    /// Retrieves the value associated with the specified key as an `i32`.
    /// Returns the provided default value if the key does not exist or cannot be parsed.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("ClientWidth", "800");
    ///
    /// let width = props.get_i32("ClientWidth", 600);
    /// assert_eq!(width, 800);
    ///
    /// let default_width = props.get_i32("ClientHeight", 600);
    /// assert_eq!(default_width, 600); // default used
    /// ```
    #[must_use]
    pub fn get_i32(&self, property_key: &str, default: i32) -> i32 {
        match self.key_value_store.get(property_key) {
            Some(value) => value.parse::<i32>().unwrap_or(default),
            None => default,
        }
    }

    /// Retrieves the value associated with the specified key as a type `T`
    /// that can be converted from `i32` and implements `TryFromPrimitive`.
    ///
    /// Returns the provided default value if the key does not exist or cannot be parsed/converted.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// use num_enum::TryFromPrimitive;
    ///
    /// #[derive(Debug, PartialEq, TryFromPrimitive)]
    /// #[repr(i32)]
    /// enum ExampleEnum {
    ///     VariantA = 0,
    ///     VariantB = 1,
    /// }
    ///
    /// let mut props = Properties::new();
    /// props.insert("ExampleKey", "1");
    ///
    /// let value: ExampleEnum = props.get_property("ExampleKey", ExampleEnum::VariantA);
    /// assert_eq!(value, ExampleEnum::VariantB);
    ///
    /// let default_value: ExampleEnum = props.get_property("MissingKey", ExampleEnum::VariantA);
    /// assert_eq!(default_value, ExampleEnum::VariantA); // default used
    /// ```
    #[must_use]
    pub fn get_property<T>(&self, property_key: &str, default: T) -> T
    where
        T: TryFrom<i32> + TryFromPrimitive,
    {
        match self.key_value_store.get(property_key) {
            Some(value) => value
                .parse::<i32>()
                .ok()
                .and_then(|v| T::try_from(v).ok())
                .unwrap_or(default),
            None => default,
        }
    }

    /// Retrieves the startup position configuration from the properties.
    /// Interprets the value associated with the specified key to determine
    /// the startup position of a window.
    ///
    /// Returns the provided default `StartUpPosition` if the key does not exist
    /// or cannot be parsed.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// use vb6parse::language::StartUpPosition;
    /// let mut props = Properties::new();
    /// props.insert("StartUpPosition", "1");
    ///
    /// let position = props.get_startup_position("StartUpPosition", StartUpPosition::WindowsDefault);
    /// assert_eq!(position, StartUpPosition::CenterOwner);
    ///
    /// let default_position = props.get_startup_position("MissingKey", StartUpPosition::WindowsDefault);
    /// assert_eq!(default_position, StartUpPosition::WindowsDefault); // default used
    /// ```
    #[must_use]
    pub fn get_startup_position(
        &self,
        property_key: &str,
        default: StartUpPosition,
    ) -> StartUpPosition {
        match self.key_value_store.get(property_key) {
            Some(value) => {
                match value.parse::<i32>() {
                    Ok(0) => {
                        let client_height = self.get_i32("ClientHeight", 3000);
                        let client_width = self.get_i32("ClientWidth", 3000);
                        let client_top = self.get_i32("ClientTop", 200);
                        let client_left = self.get_i32("ClientLeft", 100);

                        StartUpPosition::Manual {
                            client_height,
                            client_width,
                            client_top,
                            client_left,
                        }
                    }
                    Ok(1) => StartUpPosition::CenterOwner,
                    Ok(2) => StartUpPosition::CenterScreen,
                    // 3 is the default value for Windows, but we also want
                    // to default to WindowsDefault if the value is not found.
                    // I've just commented this out since leaving it in will
                    // cause Clippy to complain about the unknown value and the
                    // default value in the match arm being the same.
                    //
                    // Ok(3) => StartUpPosition::WindowsDefault,
                    _ => StartUpPosition::WindowsDefault,
                }
            }
            None => default,
        }
    }

    /// Retrieves the value associated with the specified key as a type `T`
    /// that can be converted from a string using `TryFrom<&str>`.
    ///
    /// Returns the provided default value if the key does not exist or cannot be parsed/converted.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use vb6parse::files::common::Properties;
    /// let mut props = Properties::new();
    /// props.insert("SomeKey", "SomeValue");
    ///
    /// let value: Option<String> = props.get_option("SomeKey", Some("DefaultValue".to_string()));
    /// assert_eq!(value, Some("SomeValue".to_string()));
    ///
    /// let default_value: Option<String> = props.get_option("MissingKey", Some("DefaultValue".to_string()));
    /// assert_eq!(default_value, Some("DefaultValue".to_string())); // default used
    /// ```
    #[must_use]
    pub fn get_option<T>(&self, property_key: &str, default: Option<T>) -> Option<T>
    where
        T: for<'a> TryFrom<&'a str>,
    {
        match self.key_value_store.get(property_key) {
            Some(value) => T::try_from(value.as_str()).ok().or(default),
            None => default,
        }
    }
}
