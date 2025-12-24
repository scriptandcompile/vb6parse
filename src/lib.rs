#![warn(missing_docs)]
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
//! highlighting, a language server, an interpreter, or a high speed compiler,
//! the primary goal is focused around offline analysis, legacy utilities,
//! and tools that convert VB6 code to more modern languages.

//! ## Project File Parsing
//!
//! To load a VB6 project file, you can use the `Project::parse` method.
//! This method takes a `SourceFile` as input, and returns a
//! `Project` struct that contains the parsed information.
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
//! let result = ProjectFile::parse(&project_source_file);
//!
//! if result.has_failures() {
//!     for failure in result.failures() {
//!         failure.print();
//!     }
//!     panic!("Project parse had failures");
//! }
//!
//! let project = result.unwrap();;
//!
//! assert_eq!(project.project_type, CompileTargetType::Exe);
//! assert_eq!(project.references().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.objects().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.modules().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.classes().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.designers().collect::<Vec<_>>().len(), 0);
//! assert_eq!(project.forms().collect::<Vec<_>>().len(), 2);
//! assert_eq!(project.user_controls().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.user_documents().collect::<Vec<_>>().len(), 1);
//! assert_eq!(project.properties.startup, "Form1");
//! assert_eq!(project.properties.title, "Project1");
//! assert_eq!(project.properties.exe_32_file_name, "Project1.exe");
//! ```
//!
//! Note that in the example above, the `ProjectFile::parse` method is used to parse
//! the project file. The `ProjectFile` struct contains the parsed information
//! about the project, including the project type, references, objects, modules,
//! classes, forms, user controls, etc. These references are not actually loaded
//! or parsed. This makes it possible to read a project file in parts or to
//! read a project file without having to load all the files in the project.
//!
//! ## Form File Parsing
//!
//! To load a VB6 form file, you can use the `FormFile::parse` method. This
//! pattern is very similar to the `ProjectFile::parse` method and is repeated
//! throughout the library.
//!
//! ```rust
//! use vb6parse::parsers::FormFile;
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
//! let source = vb6parse::SourceFile::decode("frmExampleForm.frm", input).expect("Failed to decode source file");
//! let result = FormFile::parse(&source);
//!
//! assert!(result.has_result());
//! assert_eq!(result.unwrap().attributes.name, "frmExampleForm");
//! ```

pub mod errors;
pub mod language;
pub mod parsers;
pub mod sourcefile;
pub mod sourcestream;
pub mod tokenize;
pub mod tokenstream;

pub use crate::errors::*;
pub use crate::language::*;
pub use crate::parsers::parse;
pub use crate::parsers::*;
pub use crate::parsers::{ConcreteSyntaxTree, SerializableTree, SyntaxKind};
pub use crate::sourcefile::*;
pub use crate::sourcestream::SourceStream;
pub use crate::tokenize::tokenize;
pub use crate::tokenstream::TokenStream;
