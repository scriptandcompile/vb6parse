//! Custom control properties.
//!
//! This module defines the `CustomControlProperties` struct, which represents
//! the properties of a custom control in a VB6 form. Custom controls are
//! non-standard controls that are not part of the standard VB6 controls, and
//! can include third-party controls and user-defined controls.
//!
//! The `CustomControlProperties` struct stores the properties of a custom
//! control in a `HashMap<String, String>`, where the keys are the property
//! names and the values are the property values.
//!
//! This struct is used as an enum variant of
//! [`ControlKind::CustomControl`](crate::language::controls::ControlKind::Custom).
//!
//! The `tag`, `name`, and `index` of the control are not included in this
//! struct, but are instead part of the parent
//! [`Control`](crate::language::controls::Control) struct.
//!

use crate::parsers::Properties;

use std::collections::HashMap;

/// Properties for a `Custom` control.
///
/// This is used to represent a non-standard control that is not part of the
/// standard VB6 controls. This can include third-party controls and user-defined
/// controls.
///
/// This is used as an enum variant of
/// [`ControlKind::CustomControl`](crate::language::controls::ControlKind::Custom).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Clone, Default, serde::Serialize)]
pub struct CustomControlProperties {
    /// A store for the properties of the custom control.
    pub property_store: HashMap<String, String>,
}

impl From<Properties> for CustomControlProperties {
    fn from(prop: Properties) -> Self {
        let mut custom_prop = CustomControlProperties::default();

        for (key, value) in prop.iter() {
            custom_prop
                .property_store
                .insert(key.clone(), value.clone());
        }

        custom_prop
    }
}

impl CustomControlProperties {
    /// Get the number of properties in the custom control.
    #[must_use]
    pub fn len(&self) -> usize {
        self.property_store.len()
    }

    /// Check if the custom control has no properties.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.property_store.is_empty()
    }
}
