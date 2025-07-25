//#![warn(missing_docs)]

//! # Summary
//!
//! `VB6Parse` is a library for parsing Visual Basic 6 (VB6) code. It is a
//! foundational library for tools and utilities that parse / analyse / convert
//! VB6 code. It is designed to be used as a base library for other tools and utilities.
//!
//! ## Design Goals
//!
//! `VB6Parse` is designed to be a fast and efficient library for parsing VB6 code.
//! Despite focusing on speed, ease of use has a high priority. While it should
//! be possible for the library to be used to create things like real-time syntax
//! highlighting, a language server, an interpreeter, or a high speed compiler,
//! the primary goal is focused around offline analysis, legacy utilities,
//! and tools that convert VB6 code to more modern languages.

//! ## Project File Parsing
//!
//! To load a VB6 project file, you can use the `VB6Project::parse` method.
//! This method takes a `file_name` and a byte slice as input, and returns a
//! `VB6Project` struct that contains the parsed information.
//!
//! The `file_name` is required for error reporting, while the byte slice
//! contains the contents of the project file. Because of the age of VB6
//! it's possible that the file is encoded in a non-standard encoding so the
//! byte slice is used instead of a string. The library will attempt to
//! decode the bytes assuming it is in windows-1252 encoding.
//!
//! The parser will attempt to detect non-english character encodings and report
//! an error if too many invalid characters are found. This does limit the library
//! to currently support 'predominantly english' source code. This is a limitation
//! which may be lifted in the future.
//!
//! ```rust
//! use vb6parse::*;
//!
//! let input = r#"Type=Exe
//! Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation
//! Object={00020430-0000-0000-C000-000000000046}#2.0#0; stdole2.tlb
//! Module=Module1; Module1.bas
//! Class=Class1; Class1.cls
//! Form=Form1.frm
//! Form=Form2.frm
//! UserControl=UserControl1.ctl
//! UserDocument=UserDocument1.uds
//! ExeName32="Project1.exe"
//! Command32=""
//! Path32=""
//! Name="Project1"
//! HelpContextID="0"
//! CompatibleMode="0"
//! MajorVer=1
//! MinorVer=0
//! RevisionVer=0
//! AutoIncrementVer=0
//! StartMode=0
//! Unattended=0
//! Retained=0
//! ThreadPerObject=0
//! MaxNumberOfThreads=1
//! DebugStartupOption=0
//! NoControlUpgrade=0
//! ServerSupportFiles=0
//! VersionCompanyName="Company Name"
//! VersionFileDescription="File Description"
//! VersionLegalCopyright="Copyright"
//! VersionLegalTrademarks="Trademark"
//! VersionProductName="Product Name"
//! VersionComments="Comments"
//! CompilationType=0
//! OptimizationType=0
//! FavorPentiumPro(tm)=0
//! CodeViewDebugInfo=0
//! NoAliasing=0
//! BoundsCheck=0
//! OverflowCheck=0
//! FlPointCheck=0
//! FDIVCheck=0
//! UnroundedFP=0
//! CondComp=""
//! ResFile32=""
//! IconForm=""
//! Startup="Form1"
//! HelpFile=""
//! Title="Project1"
//!
//! [MS Transaction Server]
//! AutoRefresh=1
//! "#;
//!
//!
//! let project_source_file = match SourceFile::decode_with_replacement("project1.vbp", input.as_bytes()) {
//!     Ok(source_file) => source_file,
//!     Err(e) => {
//!         e.print();
//!         panic!("failed to decode project source code.");
//!     }
//! };
//!
//! let result = VB6Project::parse(&project_source_file);
//!
//! if result.has_failures() {
//!     for failure in result.failures {
//!         failure.print();
//!     }
//!     panic!("Project parse had failures");
//! }
//!
//! let project = result.unwrap();;
//!
//! assert_eq!(project.project_type, CompileTargetType::Exe);
//! assert_eq!(project.references.len(), 1);
//! assert_eq!(project.objects.len(), 1);
//! assert_eq!(project.modules.len(), 1);
//! assert_eq!(project.classes.len(), 1);
//! assert_eq!(project.designers.len(), 0);
//! assert_eq!(project.forms.len(), 2);
//! assert_eq!(project.user_controls.len(), 1);
//! assert_eq!(project.user_documents.len(), 1);
//! assert_eq!(project.properties.startup, "Form1");
//! assert_eq!(project.properties.title, "Project1");
//! assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
//! ```
//!
//! Note that in the example above, the `VB6Project::parse` method is used to parse
//! the project file. The `VB6Project` struct contains the parsed information
//! about the project, including the project type, references, objects, modules,
//! classes, forms, user controls, etc. These references are not actually loaded
//! or parsed. This makes it possible to read a project file in parts or to
//! read a project file without having to load all the files in the project.
//!
//! ## Form File Parsing
//!
//! To load a VB6 form file, you can use the `VB6Form::parse` method. This
//! method takes a `file_name`, a byte slice as input, and a `resource_resolver`,
//! and returns a `VB6Form` struct that contains the parsed information. This
//! pattern is very similar to the `VB6Project::parse` method and is repeated
//! throughout the library. This makes it easier to work with dynamically
//! created VB6 code or test date without having to load the data from disk.
//!
//! A `resource_resolver` is a function that takes a `file_name` and an `offset`
//! and returns a Result containing the resource data. This is used to resolve
//! form resources such as images, icons, and other resources that are not
//! included in the form file itself.
//!
//! Included in the library is a default `resource_file_resolver` function
//! that can be used to resolve resources from a file. Breaking out the
//! resource resolver allows for more flexibility in how resources are
//! resolved as well as allowing for easier testing of the parser itself with
//! a mock resource resolver.
//!
//! ```rust
//! use vb6parse::parsers::VB6FormFile;
//! use vb6parse::parsers::resource_file_resolver;
//!
//! let input = b"VERSION 5.00\r
//! Begin VB.Form frmExampleForm\r
//!    BackColor       =   &H80000005&\r
//!    Caption         =   \"example form\"\r
//!    ClientHeight    =   6210\r
//!    ClientLeft      =   60\r
//!    ClientTop       =   645\r
//!    ClientWidth     =   9900\r
//!    BeginProperty Font\r
//!       Name            =   \"Arial\"\r
//!       Size            =   8.25\r
//!       Charset         =   0\r
//!       Weight          =   400\r
//!       Underline       =   0   'False\r
//!       Italic          =   0   'False\r
//!       Strikethrough   =   0   'False\r
//!    EndProperty\r
//!    LinkTopic       =   \"Form1\"\r
//!    ScaleHeight     =   414\r
//!    ScaleMode       =   3  'Pixel\r
//!    ScaleWidth      =   660\r
//!    StartUpPosition =   2  'CenterScreen\r
//!    Begin VB.Menu mnuFile\r
//!       Caption         =   \"&File\"\r
//!       Begin VB.Menu mnuOpenImage\r
//!          Caption         =   \"&Open image\"\r
//!       End\r
//!    End\r
//! End\r
//! Attribute VB_Name = \"frmExampleForm\"\r
//! ";
//!
//! let result = VB6FormFile::parse_with_resolver("form_parse.frm", &mut input.as_ref(), resource_file_resolver);
//!
//! assert!(result.is_ok());
//! assert_eq!(result.unwrap().form.name, "frmExampleForm");
//! ```

pub mod errors;
pub mod language;
pub mod parsers;
pub mod sourcefile;
pub mod sourcestream;

pub use crate::errors::*;
pub use crate::language::*;
pub use crate::parsers::*;
pub use crate::sourcefile::*;
