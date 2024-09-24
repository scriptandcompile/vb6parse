use std::collections::HashMap;

use crate::{
    errors::VB6ErrorKind,
    parsers::form::{build_bool_property, build_i32_property},
};

use bstr::BStr;

/// Properties for a `Timer` control.
///
/// This is used as an enum variant of
/// [`VB6ControlKind::Timer`](crate::language::controls::VB6ControlKind::Timer).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`VB6Control`](crate::language::controls::VB6Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct TimerProperties {
    pub enabled: bool,
    pub interval: i32,
    pub left: i32,
    pub top: i32,
}

impl Default for TimerProperties {
    fn default() -> Self {
        TimerProperties {
            enabled: true,
            interval: 0,
            left: 0,
            top: 0,
        }
    }
}

impl TimerProperties {
    pub fn construct_control(properties: &HashMap<&BStr, &BStr>) -> Result<Self, VB6ErrorKind> {
        let mut timer_properties = TimerProperties::default();

        timer_properties.enabled =
            build_bool_property(properties, BStr::new("Enabled"), timer_properties.enabled);
        timer_properties.interval =
            build_i32_property(properties, BStr::new("Interval"), timer_properties.interval);
        timer_properties.left =
            build_i32_property(properties, BStr::new("Left"), timer_properties.left);

        Ok(timer_properties)
    }
}
