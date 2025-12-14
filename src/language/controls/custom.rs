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
    property_store: HashMap<String, String>,
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
    #[must_use]
    pub fn len(&self) -> usize {
        self.property_store.len()
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.property_store.is_empty()
    }
}
