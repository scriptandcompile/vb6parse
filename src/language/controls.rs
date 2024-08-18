use crate::errors::VB6ErrorKind;

use bstr::{BStr, ByteSlice};

/// Represents a VB6 control.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6Control<'a> {
    pub common: VB6ControlCommonInformation<'a>,
    pub kind: VB6ControlKind<'a>,
}

/// Represents a VB6 control common information.
#[derive(Debug, PartialEq, Eq, Clone)]
pub struct VB6ControlCommonInformation<'a> {
    pub name: &'a BStr,
    pub caption: &'a BStr,
    pub back_color: VB6Color,
}

/// Represents a VB6 color.
/// The color is represented as a 32-bit RGB value.
/// The red, green, and blue values are each 8-bits.
/// This is stored and used within VB6 as a &HAABBGGRR value.
#[derive(Debug, PartialEq, Clone, Eq)]
pub struct VB6Color {
    pub alpha: u8,
    pub red: u8,
    pub green: u8,
    pub blue: u8,
}

impl VB6Color {
    /// Creates a new VB6Color.
    ///
    /// # Arguments
    ///
    /// * `alpha` - The alpha value.
    /// * `red` - The red value.
    /// * `green` - The green value.
    /// * `blue` - The blue value.
    ///
    /// # Returns
    ///
    /// A new VB6Color.
    pub fn new(alpha: u8, red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha,
            red,
            green,
            blue,
        }
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
    /// use vb6parse::language::VB6Color;
    ///
    /// let color = VB6Color::rgb(0xFF, 0x00, 0x00);
    ///
    /// assert_eq!(color.alpha, 0xFF);
    /// assert_eq!(color.red, 0xFF);
    /// assert_eq!(color.green, 0x00);
    /// assert_eq!(color.blue, 0x00);
    /// ```
    pub fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha: 0xFF,
            red,
            green,
            blue,
        }
    }

    /// Parses a VB6 color.
    ///
    /// The color is represented as a 32-bit RGB value.
    /// The red, green, and blue values are each 8-bits.
    /// This is stored and used within VB6 as a &HAABBGGRR value.
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
    /// use vb6parse::language::VB6Color;
    /// use vb6parse::vb6stream::VB6Stream;
    ///
    /// // Of course, VB6 being as it is...
    /// // the color is stored in a 'special' order.
    /// // alpha, blue, green, red
    /// let mut input = "&HAABBCCFF&";
    /// let color = VB6Color::from_hex(&input).unwrap();
    ///
    /// assert_eq!(color.alpha, 0xAA);
    /// assert_eq!(color.red, 0xFF);
    /// assert_eq!(color.green, 0xCC);
    /// assert_eq!(color.blue, 0xBB);
    /// ```
    pub fn from_hex<'a>(input: &str) -> Result<VB6Color, VB6ErrorKind> {
        let alpha_ascii = &input[2..4];
        let blue_ascii = &input[4..6];
        let green_ascii = &input[6..8];
        let red_ascii = &input[8..10];

        let alpha =
            u8::from_str_radix(alpha_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;
        let blue =
            u8::from_str_radix(blue_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;
        let green =
            u8::from_str_radix(green_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;
        let red =
            u8::from_str_radix(red_ascii, 16).map_err(|_| VB6ErrorKind::HexColorParseError)?;

        Ok(VB6Color::new(alpha, red, green, blue))
    }
}

/// Represents a VB6 control kind.
/// A VB6 control kind is an enumeration of the different kinds of
/// standard VB6 controls.
#[derive(Debug, PartialEq, Eq, Clone)]
pub enum VB6ControlKind<'a> {
    CommandButton {},
    TextBox {},
    CheckBox {},
    Line {},
    Label {},
    Frame {},
    PictureBox {},
    ComboBox {},
    HScrollBar {},
    VScrollBar {},
    Menu {
        caption: &'a BStr,
        controls: Vec<VB6Control<'a>>,
    },
    Form {
        controls: Vec<VB6Control<'a>>,
    },
}
