use crate::errors::VB6ErrorKind;

/// Represents a VB6 color.
/// The color is represented as a 24-bit RGB value.
/// The red, green, and blue values are each 8-bits.
/// This is stored and used within VB6 as text formatted as &H00BBGGRR& value.
/// if, instead, the value begins with '&H80' such as in '&H80000000&', then
/// the color is a system color. and the value is not the elements of the color,
/// but rather the index of a system color.
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub enum VB6Color {
    RGB {
        /// The red value.
        red: u8,
        /// The green value.
        green: u8,
        /// The blue value.
        blue: u8,
    },
    System {
        /// The system color index.
        index: u8,
    },
}

/// A VB6Color with red, green, and blue values of 0x00.
/// This is the same as calling `VB6Color::new(0x00, 0x00, 0x00)`.
/// This corresponds to the VB6 color constant `vbBlack`.
#[allow(dead_code)]
pub const VB_BLACK: VB6Color = VB6Color::RGB {
    red: 0x00,
    green: 0x00,
    blue: 0x00,
};

/// A VB6Color with red, green, and blue values of 0xFF.
/// This is the same as calling `VB6Color::new(0xFF, 0xFF, 0xFF)`.
/// This corresponds to the VB6 color constant `vbWhite`.
#[allow(dead_code)]
pub const VB_WHITE: VB6Color = VB6Color::RGB {
    red: 0xFF,
    green: 0xFF,
    blue: 0xFF,
};

/// A VB6Color with a red value of 0xFF and a green and blue value of 0x00.
/// This is the same as calling `VB6Color::new(0xFF, 0x00, 0x00)`.
/// This corresponds to the VB6 color constant `vbRed`.
#[allow(dead_code)]
pub const VB_RED: VB6Color = VB6Color::RGB {
    red: 0xFF,
    green: 0x00,
    blue: 0x00,
};

/// A VB6Color with a red value of 0x00 and a green value of
/// 0xFF and a blue value of 0x00.
/// This is the same as calling `VB6Color::new(0x00, 0xFF, 0x00)`.
/// This corresponds to the VB6 color constant `vbGreen`.
#[allow(dead_code)]
pub const VB_GREEN: VB6Color = VB6Color::RGB {
    red: 0x00,
    green: 0xFF,
    blue: 0x00,
};

/// A VB6Color with red and green values of 0x00 and a blue value of 0xFF.
/// This is the same as calling `VB6Color::new(0x00, 0x00, 0xFF)`.
/// This corresponds to the VB6 color constant `vbBlue`.
#[allow(dead_code)]
pub const VB_BLUE: VB6Color = VB6Color::RGB {
    red: 0x00,
    green: 0x00,
    blue: 0xFF,
};

/// A VB6Color with a red and green value of 0xFF and a blue value of 0x00.
/// This is the same as calling `VB6Color::new(0xFF, 0xFF, 0x00)`.
/// This corresponds to the VB6 color constant `vbYellow`.
#[allow(dead_code)]
pub const VB_YELLOW: VB6Color = VB6Color::RGB {
    red: 0xFF,
    green: 0xFF,
    blue: 0x00,
};

/// A VB6Color with a red and blue value of 0xFF and a green value of 0x00.
/// This is the same as calling `VB6Color::new(0xFF, 0x00, 0xFF)`.
/// This corresponds to the VB6 color constant `vbMagenta`.
#[allow(dead_code)]
pub const VB_MAGENTA: VB6Color = VB6Color::RGB {
    red: 0xFF,
    green: 0x00,
    blue: 0xFF,
};

/// A VB6Color with a red value of 0x00 and a green and blue value of 0xFF.
/// This is the same as calling `VB6Color::new(0x00, 0xFF, 0xFF)`.
/// This corresponds to the VB6 color constant `vbCyan`.
#[allow(dead_code)]
pub const VB_CYAN: VB6Color = VB6Color::RGB {
    red: 0x00,
    green: 0xFF,
    blue: 0xFF,
};

/// Darkest shadow color for 3-@ display elements.
#[allow(dead_code)]
pub const VB_3D_DK_SHADOW: VB6Color = VB6Color::System { index: 0x15 };

/// Highlight color for 3-D display elements
#[allow(dead_code)]
pub const VB_3D_HIGHLIGHT: VB6Color = VB6Color::System { index: 0x14 };

/// Second lightest 3-D color after vb3DHighlight
#[allow(dead_code)]
pub const VB_3D_LIGHT: VB6Color = VB6Color::System { index: 0x16 };

/// Lightest shadow color for 3-D display elements
#[allow(dead_code)]
pub const VB_3D_SHADOW: VB6Color = VB6Color::System { index: 0x10 };

/// Border color of active window
#[allow(dead_code)]
pub const VB_ACTIVE_BORDER: VB6Color = VB6Color::System { index: 0x0A };

/// Color of the title bar for the active window
#[allow(dead_code)]
pub const VB_ACTIVE_TITLE_BAR: VB6Color = VB6Color::System { index: 0x02 };

/// Background color of multiple document interface (MDI) applications
#[allow(dead_code)]
pub const VB_APPLICATION_WORKSPACE: VB6Color = VB6Color::System { index: 0x0C };

/// Color of shading on the face of command buttons
#[allow(dead_code)]
pub const VB_BUTTON_FACE: VB6Color = VB6Color::System { index: 0x0F };

/// Color of shading on the face of command buttons
#[allow(dead_code)]
pub const VB_3D_FACE: VB6Color = VB6Color::System { index: 0x0F };

/// Color of shading on the edge of command buttons
#[allow(dead_code)]
pub const VB_BUTTON_SHADOW: VB6Color = VB6Color::System { index: 0x10 };

/// Text color on push buttons
#[allow(dead_code)]
pub const VB_BUTTON_TEXT: VB6Color = VB6Color::System { index: 0x12 };

/// Desktop color
#[allow(dead_code)]
pub const VB_DESKTOP: VB6Color = VB6Color::System { index: 0x01 };

/// Grayed (disabled) text
#[allow(dead_code)]
pub const VB_GRAY_TEXT: VB6Color = VB6Color::System { index: 0x11 };

/// Background color of items selected in a control
#[allow(dead_code)]
pub const VB_HIGHLIGHT: VB6Color = VB6Color::System { index: 0x0D };

/// Text color of items selected in a control
#[allow(dead_code)]
pub const VB_HIGHLIGHT_TEXT: VB6Color = VB6Color::System { index: 0x0E };

/// Border color of inactive window
#[allow(dead_code)]
pub const VB_INACTIVE_BORDER: VB6Color = VB6Color::System { index: 0x0B };

/// Color of text in an inactive caption
#[allow(dead_code)]
pub const VB_INACTIVE_CAPTION_TEXT: VB6Color = VB6Color::System { index: 0x13 };

/// Color of the title bar for the inactive window
#[allow(dead_code)]
pub const VB_INACTIVE_TITLE_BAR: VB6Color = VB6Color::System { index: 0x03 };

/// Background color of ToolTips
#[allow(dead_code)]
pub const VB_INFO_BACKGROUND: VB6Color = VB6Color::System { index: 0x18 };

/// Background color of ToolTips
#[allow(dead_code)]
pub const VB_MSG_BOX_TEXT: VB6Color = VB6Color::System { index: 0x18 };

/// Color of text in ToolTips
#[allow(dead_code)]
pub const VB_INFO_TEXT: VB6Color = VB6Color::System { index: 0x17 };

/// Color of text in ToolTips
#[allow(dead_code)]
pub const VB_MSG_BOX: VB6Color = VB6Color::System { index: 0x17 };

/// Menu background color
#[allow(dead_code)]
pub const VB_MENU_BAR: VB6Color = VB6Color::System { index: 0x04 };

/// Color of text on menus
#[allow(dead_code)]
pub const VB_MENU_TEXT: VB6Color = VB6Color::System { index: 0x07 };

/// Scrollbar color
#[allow(dead_code)]
pub const VB_SCROLL_BARS: VB6Color = VB6Color::System { index: 0x00 };

/// Color of text in caption, size box, and scroll arrow
#[allow(dead_code)]
pub const VB_TITLE_BAR_TEXT: VB6Color = VB6Color::System { index: 0x09 };

/// Window background color
#[allow(dead_code)]
pub const VB_WINDOW_BACKGROUND: VB6Color = VB6Color::System { index: 0x05 };

/// Window frame color
#[allow(dead_code)]
pub const VB_WINDOW_FRAME: VB6Color = VB6Color::System { index: 0x06 };

/// Color of text in windows
#[allow(dead_code)]
pub const VB_WINDOW_TEXT: VB6Color = VB6Color::System { index: 0x08 };

impl VB6Color {
    /// Creates a new VB6Color.
    ///
    /// # Arguments
    ///
    /// * `red` - The red value.
    /// * `green` - The green value.
    /// * `blue` - The blue value.
    ///
    /// # Returns
    ///
    /// A new VB6Color.
    pub fn new(red: u8, green: u8, blue: u8) -> Self {
        VB6Color::RGB { red, green, blue }
    }

    /// Creates a new VB6Color that represents a system color.
    /// The index is the index of the system color.
    /// This is the same as calling `VB6Color::System { index }`.
    ///
    /// # Arguments
    /// * `index` - The index of the system color.
    ///
    /// # Returns
    ///
    /// A new VB6Color of a system type.
    pub fn system(index: u8) -> Self {
        VB6Color::System { index }
    }

    /// Creates a new VB6Color with an alpha value of 0xFF.
    /// This is the same as calling `VB6Color::new(0xFF, red, green, blue)`.
    /// This is useful when you don't need to specify an alpha value.
    ///
    /// # Arguments
    ///
    /// * `red` - The red value.
    /// * `green` - The green value.
    /// * `blue` - The blue value.
    ///
    /// # Returns
    ///
    /// A new VB6Color.
    ///
    /// # Example
    ///
    /// ```rust
    /// use std::matches;
    ///
    /// use vb6parse::language::VB6Color;
    ///
    /// let color = VB6Color::rgb(0xFF, 0x00, 0x00);
    ///
    /// assert!(matches!(color, VB6Color::RGB { .. } ));
    /// assert_eq!(color, VB6Color::RGB { red: 0xFF, green: 0x00, blue: 0x00 });
    /// ```
    pub fn rgb(red: u8, green: u8, blue: u8) -> Self {
        VB6Color::RGB { red, green, blue }
    }

    /// Parses a VB6 color.
    ///
    /// The color is represented as a 32-bit RGB value.
    /// The red, green, and blue values are each 8-bits.
    /// This is stored and used in VB6 as a formatted hex text value &HAABBGGRR&.
    ///
    /// # Arguments
    ///
    /// * `input` - The input to parse.
    ///
    /// # Returns
    ///
    /// The VB6 color.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::{language::VB6Color, parsers::VB6Stream};
    ///
    /// // Of course, VB6 being as it is...
    /// // the color is stored in a 'special' order.
    /// // blue, green, red
    /// let mut input = "&H00BBCCFF&";
    /// let color = VB6Color::from_hex(&input).unwrap();
    ///
    /// assert!(matches!(color, VB6Color::RGB { .. } ));
    /// assert_eq!(color, VB6Color::RGB { red: 0xFF, green: 0xCC, blue: 0xBB });
    /// ```
    pub fn from_hex<'a>(input: &str) -> Result<VB6Color, VB6ErrorKind> {
        let kind_ascii = &input[2..4];

        let kind =
            u8::from_str_radix(kind_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;

        if kind == 0x80 {
            // System color
            let index = u8::from_str_radix(&input[8..10], 16)
                .map_err(|_| VB6ErrorKind::HexColorParseError)?;
            return Ok(VB6Color::system(index));
        } else if kind != 0x00 {
            return Err(VB6ErrorKind::HexColorParseError);
        }

        let blue_ascii = &input[4..6];
        let green_ascii = &input[6..8];
        let red_ascii = &input[8..10];

        let blue =
            u8::from_str_radix(blue_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;
        let green =
            u8::from_str_radix(green_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;
        let red =
            u8::from_str_radix(red_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;

        Ok(VB6Color::new(red, green, blue))
    }
}
