//! Properties for a `Timer` control.
//!
//! This is used as an enum variant of
//! [`ControlKind::Timer`](crate::language::controls::ControlKind::Timer).
//! tag, name, and index are not included in this struct, but instead are part
//! of the parent [`Control`](crate::language::controls::Control) struct.
//!

use crate::language::controls::Activation;

use crate::parsers::Properties;

/// Properties for a `Timer` control.
///
/// This is used as an enum variant of
/// [`ControlKind::Timer`](crate::language::controls::ControlKind::Timer).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [`Control`](crate::language::controls::Control) struct.
#[derive(Debug, PartialEq, Eq, Clone, serde::Serialize)]
pub struct TimerProperties {
    /// Whether the timer is enabled.
    pub enabled: Activation,
    /// The timer interval in milliseconds.
    pub interval: i32,
    /// The left position of the timer control.
    pub left: i32,
    /// The top position of the timer control.
    pub top: i32,
}

impl Default for TimerProperties {
    fn default() -> Self {
        TimerProperties {
            enabled: Activation::Enabled,
            interval: 0,
            left: 0,
            top: 0,
        }
    }
}

impl From<Properties> for TimerProperties {
    fn from(prop: Properties) -> Self {
        let mut timer_prop = TimerProperties::default();

        timer_prop.enabled = prop.get_property("Enabled", timer_prop.enabled);
        timer_prop.interval = prop.get_i32("Interval", timer_prop.interval);
        timer_prop.left = prop.get_i32("Left", timer_prop.left);

        timer_prop
    }
}
