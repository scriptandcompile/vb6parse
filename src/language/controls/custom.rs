use crate::parsers::Properties;

use bstr::BString;
use std::collections::HashMap;

/// Properties for a `Custom` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::CustomControl`](crate::language::controls::VB6ControlKind::Custom).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Clone, serde::Serialize)]
pub struct CustomControlProperties {
    property_store: HashMap<BString, Vec<u8>>,
}

impl Default for CustomControlProperties {
    fn default() -> Self {
        CustomControlProperties {
            property_store: HashMap::new(),
        }
    }
}

impl<'a> From<Properties<'a>> for CustomControlProperties {
    fn from(prop: Properties<'a>) -> Self {
        let mut custom_prop = CustomControlProperties::default();

        for (key, value) in prop.iter() {
            custom_prop
                .property_store
                .insert(key.to_owned(), value.to_owned());
        }

        custom_prop
    }
}

impl CustomControlProperties {
    pub fn len(&self) -> usize {
        self.property_store.len()
    }
}
