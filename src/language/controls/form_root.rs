//! Top-level form root types for VB6 form files.
//!
//! This module defines the type-safe representation of form file roots.
//! VB6 form files (`.frm`, `.ctl`, `.dob`) must have either a `VB.Form` or
//! `VB.MDIForm` at the top level - it's not possible to have any other control
//! type as the root element.
//!
//! The [`FormRoot`] enum enforces this constraint at the type level, preventing
//! invalid states from being representable in the API.

use std::fmt::{Display, Formatter};

use serde::Serialize;

use super::form::FormProperties;
use super::mdiform::MDIFormProperties;
use super::{Control, MenuControl};

/// Top-level form root - can only be a Form or MDIForm.
///
/// This enum enforces the VB6 constraint that form files must have either
/// a standard `VB.Form` or a `VB.MDIForm` as the top-level element.
///
/// # Examples
///
/// ```
/// use vb6parse::files::FormFile;
/// use vb6parse::io::SourceFile;
/// use vb6parse::language::FormRoot;
///
/// let source = r#"VERSION 5.00
/// Begin VB.Form Form1
///    Caption = "Hello World"
/// End
/// "#;
///
/// let source_file = SourceFile::decode_with_replacement("test.frm", source.as_bytes()).expect("Failed to decode");
/// let (form_file, _) = FormFile::parse(&source_file).unpack();
/// let form_file = form_file.expect("Failed to parse form file");
///
/// match &form_file.form {
///     FormRoot::Form(form) => {
///         assert_eq!(form.name, "Form1");
///         assert_eq!(form.properties.caption, "Hello World");
///     }
///     FormRoot::MDIForm(_) => panic!("Expected a Form, not an MDIForm"),
/// }
/// ```
#[derive(Debug, PartialEq, Clone, Serialize)]
pub enum FormRoot {
    /// Standard VB6 Form
    Form(Form),
    /// Multi-Document Interface Form
    MDIForm(MDIForm),
}

impl FormRoot {
    /// Get the form name (works for both Form and MDIForm).
    ///
    /// # Returns
    ///
    /// The name of the form as a string slice.
    ///
    /// # Examples
    ///
    /// ```
    /// use vb6parse::language::{FormRoot, Form};
    /// use vb6parse::language::FormProperties;
    ///
    /// let form = Form {
    ///     name: "MainForm".to_string(),
    ///     tag: String::new(),
    ///     index: 0,
    ///     properties: FormProperties::default(),
    ///     controls: Vec::new(),
    ///     menus: Vec::new(),
    /// };
    ///
    /// let root = FormRoot::Form(form);
    /// assert_eq!(root.name(), "MainForm");
    /// ```
    #[must_use]
    pub fn name(&self) -> &str {
        match self {
            FormRoot::Form(f) => &f.name,
            FormRoot::MDIForm(m) => &m.name,
        }
    }

    /// Get a mutable reference to the form name.
    ///
    /// # Returns
    ///
    /// A mutable reference to the form name string.
    #[must_use]
    pub fn name_mut(&mut self) -> &mut String {
        match self {
            FormRoot::Form(f) => &mut f.name,
            FormRoot::MDIForm(m) => &mut m.name,
        }
    }

    /// Get a reference to child controls.
    ///
    /// # Returns
    ///
    /// A slice of the child controls.
    #[must_use]
    pub fn controls(&self) -> &[Control] {
        match self {
            FormRoot::Form(f) => &f.controls,
            FormRoot::MDIForm(m) => &m.controls,
        }
    }

    /// Get a mutable reference to child controls.
    ///
    /// # Returns
    ///
    /// A mutable reference to the vector of child controls.
    #[must_use]
    pub fn controls_mut(&mut self) -> &mut Vec<Control> {
        match self {
            FormRoot::Form(f) => &mut f.controls,
            FormRoot::MDIForm(m) => &mut m.controls,
        }
    }

    /// Get a reference to menus.
    ///
    /// # Returns
    ///
    /// A slice of menu controls.
    #[must_use]
    pub fn menus(&self) -> &[MenuControl] {
        match self {
            FormRoot::Form(f) => &f.menus,
            FormRoot::MDIForm(m) => &m.menus,
        }
    }

    /// Check if this is a standard Form (not an MDIForm).
    ///
    /// # Returns
    ///
    /// `true` if this is a `FormRoot::Form`, `false` otherwise.
    ///
    /// # Examples
    ///
    /// ```
    /// use vb6parse::language::{FormRoot, Form};
    /// use vb6parse::language::FormProperties;
    ///
    /// let form = Form {
    ///     name: "Form1".to_string(),
    ///     tag: String::new(),
    ///     index: 0,
    ///     properties: FormProperties::default(),
    ///     controls: Vec::new(),
    ///     menus: Vec::new(),
    /// };
    ///
    /// let root = FormRoot::Form(form);
    /// assert!(root.is_form());
    /// assert!(!root.is_mdi_form());
    /// ```
    #[must_use]
    pub fn is_form(&self) -> bool {
        matches!(self, FormRoot::Form(_))
    }

    /// Check if this is an MDIForm.
    ///
    /// # Returns
    ///
    /// `true` if this is a `FormRoot::MDIForm`, `false` otherwise.
    #[must_use]
    pub fn is_mdi_form(&self) -> bool {
        matches!(self, FormRoot::MDIForm(_))
    }
}

impl Display for FormRoot {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            FormRoot::Form(form) => write!(f, "Form: {}", form.name),
            FormRoot::MDIForm(mdi_form) => write!(f, "MDIForm: {}", mdi_form.name),
        }
    }
}

/// A standard VB6 Form with its controls and menus.
///
/// This represents a complete VB6 form, including its properties, child controls,
/// and menu structure. Forms are the primary container for UI elements in VB6
/// applications.
///
/// # Examples
///
/// ```
/// use vb6parse::language::{Form, FormProperties};
///
/// let form = Form {
///     name: "MainForm".to_string(),
///     tag: String::new(),
///     index: 0,
///     properties: FormProperties::default(),
///     controls: Vec::new(),
///     menus: Vec::new(),
/// };
///
/// assert_eq!(form.name, "MainForm");
/// assert_eq!(form.controls.len(), 0);
/// ```
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct Form {
    /// Form name (from Begin statement or VB_Name attribute).
    pub name: String,
    /// Tag value (arbitrary string data associated with the form).
    pub tag: String,
    /// Form index (typically 0 for top-level forms).
    pub index: i32,
    /// Form-specific properties (caption, size, colors, etc.).
    pub properties: FormProperties,
    /// Child controls contained in the form.
    pub controls: Vec<Control>,
    /// Form menus.
    pub menus: Vec<MenuControl>,
}

impl Display for Form {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "Form: {} ({} controls, {} menus)",
            self.name,
            self.controls.len(),
            self.menus.len()
        )
    }
}

/// An MDI (Multi-Document Interface) Form with its controls and menus.
///
/// MDI forms are special container forms that can host multiple child forms
/// (MDI child forms). They provide a framework for applications that work with
/// multiple documents simultaneously.
///
/// # Examples
///
/// ```
/// use vb6parse::language::{MDIForm, MDIFormProperties};
///
/// let mdi_form = MDIForm {
///     name: "MDIMain".to_string(),
///     tag: String::new(),
///     index: 0,
///     properties: MDIFormProperties::default(),
///     controls: Vec::new(),
///     menus: Vec::new(),
/// };
///
/// assert_eq!(mdi_form.name, "MDIMain");
/// assert_eq!(mdi_form.controls.len(), 0);
/// ```
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct MDIForm {
    /// MDI Form name (from Begin statement or VB_Name attribute).
    pub name: String,
    /// Tag value (arbitrary string data associated with the form).
    pub tag: String,
    /// Form index (typically 0).
    pub index: i32,
    /// MDI Form-specific properties (caption, size, colors, etc.).
    pub properties: MDIFormProperties,
    /// Child controls contained in the MDI form.
    pub controls: Vec<Control>,
    /// MDI Form menus.
    pub menus: Vec<MenuControl>,
}

impl Display for MDIForm {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "MDIForm: {} ({} controls, {} menus)",
            self.name,
            self.controls.len(),
            self.menus.len()
        )
    }
}
