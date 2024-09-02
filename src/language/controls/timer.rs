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
