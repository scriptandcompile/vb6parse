use crate::errors::VB6ErrorKind;

/// Represents a VB6 color.
/// The color is represented as a 32-bit RGB value.
/// The red, green, and blue values are each 8-bits.
/// This is stored and used within VB6 as text formatted as &HAABBGGRR& value.
#[derive(Debug, PartialEq, Clone, Eq)]
pub struct VB6Color {
    /// The alpha value.
    pub alpha: u8,
    /// The red value.
    pub red: u8,
    /// The green value.
    pub green: u8,
    /// The blue value.
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
