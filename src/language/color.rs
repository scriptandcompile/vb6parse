use crate::errors::FormErrorKind;

/// `Colors` are 24 bits with 8 bits for red, green, and blue.
///
/// `Colors` are stored and used within VB6 as text formatted as '&H00BBGGRR&'.
/// If, instead, the value begins with '&H80' such as in '&H80000000&', then
/// the color is a system color. and the value is not the elements of the color,
/// but rather the index of a system color.
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub enum Color {
    /// A color represented by red, green, and blue values.
    /// The values are 8 bits each.
    /// The values are stored in the order of red, green, blue.
    /// This is the same as calling `Color::new(red, green, blue)`.
    RGB {
        /// The red value.
        red: u8,
        /// The green value.
        green: u8,
        /// The blue value.
        blue: u8,
    },
    /// A system color represented by an index.
    /// The index is the index of the system color.
    /// This is the same as calling `Color::system(index)`.
    System {
        /// The system color index.
        index: u8,
    },
}

impl Color {
    /// Converts the `Color` to a VB6 formatted string.
    ///
    /// # Returns
    ///
    /// A string formatted as '&H00BBGGRR&' for RGB colors
    /// or '&H800000II&' for system colors.
    #[must_use]
    pub fn to_vb_string(&self) -> String {
        match self {
            Color::RGB { red, green, blue } => {
                format!("&H00{blue:02X}{green:02X}{red:02X}&")
            }
            Color::System { index } => {
                format!("&H80{index:06X}&")
            }
        }
    }
}

/// A `Color` with red, green, and blue values of 0x00.
/// This is the same as calling `Color::new(0x00, 0x00, 0x00)`.
/// This corresponds to the VB6 color constant `vbBlack`.
#[allow(dead_code)]
pub const VB_BLACK: Color = Color::RGB {
    red: 0x00,
    green: 0x00,
    blue: 0x00,
};

/// A `Color` with red, green, and blue values of 0xFF.
/// This is the same as calling `Color::new(0xFF, 0xFF, 0xFF)`.
/// This corresponds to the VB6 color constant `vbWhite`.
#[allow(dead_code)]
pub const VB_WHITE: Color = Color::RGB {
    red: 0xFF,
    green: 0xFF,
    blue: 0xFF,
};

/// A `Color` with a red value of 0xFF and a green and blue value of 0x00.
/// This is the same as calling `Color::new(0xFF, 0x00, 0x00)`.
/// This corresponds to the VB6 color constant `vbRed`.
#[allow(dead_code)]
pub const VB_RED: Color = Color::RGB {
    red: 0xFF,
    green: 0x00,
    blue: 0x00,
};

/// A `Color` with a red value of 0x00 and a green value of
/// 0xFF and a blue value of 0x00.
///
/// This is the same as calling `Color::new(0x00, 0xFF, 0x00)`.
/// This corresponds to the VB6 color constant `vbGreen`.
#[allow(dead_code)]
pub const VB_GREEN: Color = Color::RGB {
    red: 0x00,
    green: 0xFF,
    blue: 0x00,
};

/// A `Color` with red and green values of 0x00 and a blue value of 0xFF.
///
/// This is the same as calling `Color::new(0x00, 0x00, 0xFF)`.
/// This corresponds to the VB6 color constant `vbBlue`.
#[allow(dead_code)]
pub const VB_BLUE: Color = Color::RGB {
    red: 0x00,
    green: 0x00,
    blue: 0xFF,
};

/// A `Color` with a red and green value of 0xFF and a blue value of 0x00.
///
/// This is the same as calling `Color::new(0xFF, 0xFF, 0x00)`.
/// This corresponds to the VB6 color constant `vbYellow`.
#[allow(dead_code)]
pub const VB_YELLOW: Color = Color::RGB {
    red: 0xFF,
    green: 0xFF,
    blue: 0x00,
};

/// A `Color` with a red and blue value of 0xFF and a green value of 0x00.
///
/// This is the same as calling `Color::new(0xFF, 0x00, 0xFF)`.
/// This corresponds to the VB6 color constant `vbMagenta`.
#[allow(dead_code)]
pub const VB_MAGENTA: Color = Color::RGB {
    red: 0xFF,
    green: 0x00,
    blue: 0xFF,
};

/// A `Color` with a red value of 0x00 and a green and blue value of 0xFF.
///
/// This is the same as calling `Color::new(0x00, 0xFF, 0xFF)`.
/// This corresponds to the VB6 color constant `vbCyan`.
#[allow(dead_code)]
pub const VB_CYAN: Color = Color::RGB {
    red: 0x00,
    green: 0xFF,
    blue: 0xFF,
};

/// Scrollbar color
#[allow(dead_code)]
pub const VB_SCROLL_BARS: Color = Color::System { index: 0x00 };

/// Desktop color
#[allow(dead_code)]
pub const VB_DESKTOP: Color = Color::System { index: 0x01 };

/// Color of the title bar for the active window
#[allow(dead_code)]
pub const VB_ACTIVE_TITLE_BAR: Color = Color::System { index: 0x02 };

/// Color of the title bar for the inactive window
#[allow(dead_code)]
pub const VB_INACTIVE_TITLE_BAR: Color = Color::System { index: 0x03 };

/// Menu background color
#[allow(dead_code)]
pub const VB_MENU_BAR: Color = Color::System { index: 0x04 };

/// Window background color
#[allow(dead_code)]
pub const VB_WINDOW_BACKGROUND: Color = Color::System { index: 0x05 };

/// Window frame color
#[allow(dead_code)]
pub const VB_WINDOW_FRAME: Color = Color::System { index: 0x06 };

/// Color of text on menus
#[allow(dead_code)]
pub const VB_MENU_TEXT: Color = Color::System { index: 0x07 };

/// Color of text in windows
#[allow(dead_code)]
pub const VB_WINDOW_TEXT: Color = Color::System { index: 0x08 };

/// Color of text in caption, size box, and scroll arrow
#[allow(dead_code)]
pub const VB_TITLE_BAR_TEXT: Color = Color::System { index: 0x09 };

/// Border color of active window
#[allow(dead_code)]
pub const VB_ACTIVE_BORDER: Color = Color::System { index: 0x0A };

/// Border color of inactive window
#[allow(dead_code)]
pub const VB_INACTIVE_BORDER: Color = Color::System { index: 0x0B };

/// Background color of multiple document interface (MDI) applications
#[allow(dead_code)]
pub const VB_APPLICATION_WORKSPACE: Color = Color::System { index: 0x0C };

/// Background color of items selected in a control
#[allow(dead_code)]
pub const VB_HIGHLIGHT: Color = Color::System { index: 0x0D };

/// Text color of items selected in a control
#[allow(dead_code)]
pub const VB_HIGHLIGHT_TEXT: Color = Color::System { index: 0x0E };

/// Color of shading on the face of command buttons
#[allow(dead_code)]
pub const VB_BUTTON_FACE: Color = Color::System { index: 0x0F };

/// Color of shading on the face of command buttons
#[allow(dead_code)]
pub const VB_3D_FACE: Color = Color::System { index: 0x0F };

/// Lightest shadow color for 3-D display elements
#[allow(dead_code)]
pub const VB_3D_SHADOW: Color = Color::System { index: 0x10 };

/// Color of shading on the edge of command buttons
#[allow(dead_code)]
pub const VB_BUTTON_SHADOW: Color = Color::System { index: 0x10 };

/// Grayed (disabled) text
#[allow(dead_code)]
pub const VB_GRAY_TEXT: Color = Color::System { index: 0x11 };

/// Text color on push buttons
#[allow(dead_code)]
pub const VB_BUTTON_TEXT: Color = Color::System { index: 0x12 };

/// Color of text in an inactive caption
#[allow(dead_code)]
pub const VB_INACTIVE_CAPTION_TEXT: Color = Color::System { index: 0x13 };

/// Highlight color for 3-D display elements
#[allow(dead_code)]
pub const VB_3D_HIGHLIGHT: Color = Color::System { index: 0x14 };

/// Darkest shadow color for 3-D display elements.
#[allow(dead_code)]
pub const VB_3D_DK_SHADOW: Color = Color::System { index: 0x15 };

/// Second lightest 3-D color after vb3DHighlight
#[allow(dead_code)]
pub const VB_3D_LIGHT: Color = Color::System { index: 0x16 };

/// Color of text in tool tips
#[allow(dead_code)]
pub const VB_INFO_TEXT: Color = Color::System { index: 0x17 };

/// Color of text in tool tips
#[allow(dead_code)]
pub const VB_MSG_BOX: Color = Color::System { index: 0x17 };

/// Background color of tool tips
#[allow(dead_code)]
pub const VB_INFO_BACKGROUND: Color = Color::System { index: 0x18 };

/// Background color of tool tips
#[allow(dead_code)]
pub const VB_MSG_BOX_TEXT: Color = Color::System { index: 0x18 };

impl Color {
    /// Creates a new `Color`.
    ///
    /// # Arguments
    ///
    /// * `red` - The red value.
    /// * `green` - The green value.
    /// * `blue` - The blue value.
    ///
    /// # Returns
    ///
    /// A new RGB `Color`.
    #[must_use]
    pub fn new(red: u8, green: u8, blue: u8) -> Self {
        Color::RGB { red, green, blue }
    }

    /// Creates a new `Color` that represents a system color.
    /// The index is the index of the system color.
    /// This is the same as calling `Color::System { index }`.
    ///
    /// # Arguments
    /// * `index` - The index of the system color.
    ///
    /// # Returns
    ///
    /// A new system `Color`.
    #[must_use]
    pub fn system(index: u8) -> Self {
        Color::System { index }
    }

    /// Creates a new `Color` with an RGB value.
    /// This is the same as calling `Color::new(red, green, blue)`.
    ///
    /// # Arguments
    ///
    /// * `red` - The red value.
    /// * `green` - The green value.
    /// * `blue` - The blue value.
    ///
    /// # Returns
    ///
    /// A new RGB `Color`.
    ///
    /// # Example
    ///
    /// ```rust
    /// use std::matches;
    ///
    /// use vb6parse::language::Color;
    ///
    /// let color = Color::rgb(0xFF, 0x33, 0x12);
    ///
    /// assert!(matches!(color, Color::RGB { .. } ));
    /// assert_eq!(color, Color::RGB { red: 0xFF, green: 0x33, blue: 0x12 });
    /// ```
    #[must_use]
    pub fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Color::RGB { red, green, blue }
    }

    /// Parses a `Color`.
    ///
    /// The color is represented as a 24-bit RGB value.
    /// The red, green, and blue values are each 8-bits.
    /// This is stored and used in VB6 as a formatted hex text value &H00BBGGRR&.
    ///
    /// # Arguments
    ///
    /// * `input` - The input to parse.
    ///
    /// # Errors
    ///
    /// If the input is not a valid hex color of either a
    /// '&H00BBGGRR&' or '&H800000II&' format, an error is returned.
    ///
    /// The '&H00BBGGRR&' format is a 24-bit RGB color in the order of blue, green, red.
    /// where each element is in hex format.
    ///
    /// The '&H800000II&' format is a system color with the index value of 'II'.
    /// Again, where 'II' is in hex format.
    ///
    /// # Returns
    ///
    /// The `Color`.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::language::Color;
    ///
    /// // Of course, VB6 being as it is...
    /// // the color is stored in a 'special' order.
    /// // blue, green, red
    /// let mut input = "&H00BBCCFF&";
    /// let color = Color::from_hex(&input).unwrap();
    ///
    /// assert!(matches!(color, Color::RGB { .. } ));
    /// assert_eq!(color, Color::RGB { red: 0xFF, green: 0xCC, blue: 0xBB });
    /// ```
    pub fn from_hex(input: &str) -> Result<Color, FormErrorKind> {
        let kind_ascii = &input[2..4];

        let kind =
            u8::from_str_radix(kind_ascii, 16).map_err(|_| FormErrorKind::HexColorParseError)?;

        if kind == 0x80 {
            // System color
            let index = u8::from_str_radix(&input[8..10], 16)
                .map_err(|_| FormErrorKind::HexColorParseError)?;
            return Ok(Color::system(index));
        } else if kind != 0x00 {
            return Err(FormErrorKind::HexColorParseError);
        }

        let blue_ascii = &input[4..6];
        let green_ascii = &input[6..8];
        let red_ascii = &input[8..10];

        let blue =
            u8::from_str_radix(blue_ascii, 16).map_err(|_| FormErrorKind::HexColorParseError)?;
        let green =
            u8::from_str_radix(green_ascii, 16).map_err(|_| FormErrorKind::HexColorParseError)?;
        let red =
            u8::from_str_radix(red_ascii, 16).map_err(|_| FormErrorKind::HexColorParseError)?;

        Ok(Color::new(red, green, blue))
    }
}
