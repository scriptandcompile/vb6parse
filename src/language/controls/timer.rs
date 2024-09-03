/// Properties for a Timer control. This is used as an enum variant of
/// [VB6ControlKind::Timer](crate::language::controls::VB6ControlKind::Timer).
/// tag, name, and index are not included in this struct, but instead are part
/// of the parent [VB6Control](crate::language::controls::VB6Control) struct.
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
