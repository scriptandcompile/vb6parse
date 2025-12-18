//! VB6 Class File Parser Module
//!
//! This module provides functionality to parse VB6 class files (.cls).
//! It defines the `ClassFile` struct representing a VB6 class file,
//! along with methods to parse the file and extract relevant information
//! such as version, properties, and attributes.
//!

pub mod properties;

use std::collections::HashMap;

use serde::Serialize;

use crate::{
    errors::ClassErrorKind,
    parsers::{
        class::properties::{ClassHeader, ClassProperties},
        cst::{parse, serialize_cst},
        header::{extract_version, Creatable, Exposed, FileAttributes, NameSpace, PreDeclaredID},
        SyntaxKind,
    },
    tokenize::tokenize,
    ConcreteSyntaxTree, ParseResult, SourceFile,
};

/// Represents a VB6 class file.
/// A VB6 class file contains a header and a concrete syntax tree.
///
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
///
/// The cst contains the concrete syntax tree of the code of the class file.
#[derive(Debug, PartialEq, Clone, Serialize)]
pub struct ClassFile {
    /// The header of the class file.
    pub header: ClassHeader,
    /// The concrete syntax tree of the class file.
    /// This excludes nodes that are already represented in the header.
    #[serde(serialize_with = "serialize_cst")]
    pub cst: ConcreteSyntaxTree,
}

impl ClassFile {
    /// Parses a VB6 class file from a byte slice.
    ///
    /// # Arguments
    ///
    /// * `input` The byte slice to parse.
    ///
    /// # Returns
    ///
    /// A result containing the parsed VB6 class file or an error.
    ///
    /// # Errors
    ///
    /// An error will be returned if the input is not a valid VB6 class file.
    ///
    /// # Example
    ///
    /// ```rust
    /// use vb6parse::parsers::ClassFile;
    /// use vb6parse::SourceFile;
    ///
    /// let input = b"VERSION 1.0 CLASS
    /// BEGIN
    ///   MultiUse = -1  'True
    ///   Persistable = 0  'NotPersistable
    ///   DataBindingBehavior = 0  'vbNone
    ///   DataSourceBehavior = 0  'vbNone
    ///   MTSTransactionMode = 0  'NotAnMTSObject
    /// END
    /// Attribute VB_Name = \"Something\"
    /// Attribute VB_GlobalNameSpace = False
    /// Attribute VB_Creatable = True
    /// Attribute VB_PredeclaredId = False
    /// Attribute VB_Exposed = False
    /// ";
    ///
    /// let result = SourceFile::decode_with_replacement("class_parse.cls", input);
    ///
    /// let source_file = match result {
    ///     Ok(source_file) => source_file,
    ///     Err(e) => panic!("Failed to decode source file 'class_parse.cls': {e:?}"),
    /// };
    ///
    /// let result = ClassFile::parse(&source_file);
    ///
    /// assert!(result.has_result());
    /// ```
    #[must_use]
    pub fn parse(source_file: &SourceFile) -> ParseResult<'_, Self, ClassErrorKind<'_>> {
        let mut input = source_file.get_source_stream();

        let mut failures = vec![];

        // Parse tokens and create CST
        let token_stream_result = tokenize(&mut input);

        if token_stream_result.has_failures() {
            for failure in token_stream_result.failures {
                failures.push(failure.into());
            }
        }

        let Some(token_stream) = token_stream_result.result else {
            return ParseResult {
                result: None,
                failures,
            };
        };

        // Parse CST
        let cst = parse(token_stream);

        // Extract version from CST
        let Some(version) = extract_version(&cst) else {
            let error = source_file
                .get_source_stream()
                .generate_error(ClassErrorKind::VersionKeywordMissing);
            failures.push(error);

            return ParseResult {
                result: None,
                failures,
            };
        };

        // Extract properties from CST
        let properties = extract_properties(&cst);

        // Extract attributes from CST
        let attributes = extract_attributes(&cst);

        let header = ClassHeader {
            version,
            attributes,
            properties,
        };

        // Filter out nodes that are already extracted to avoid duplication
        // For class files, we remove:
        // - VersionStatement (already in header.version)
        // - PropertiesBlock (BEGIN...END - already in header.properties)
        // - AttributeStatement nodes (already in header.attributes)
        let filtered_cst = cst.without_kinds(&[
            SyntaxKind::VersionStatement,
            SyntaxKind::PropertiesBlock,
            SyntaxKind::AttributeStatement,
        ]);

        ParseResult {
            result: Some(ClassFile {
                header,
                cst: filtered_cst,
            }),
            failures,
        }
    }
}

/// Extract FileAttributes from AttributeStatement nodes in the CST
fn extract_attributes(cst: &crate::parsers::ConcreteSyntaxTree) -> FileAttributes {
    let mut name = String::new();
    let mut global_name_space = NameSpace::Local;
    let mut creatable = Creatable::True;
    let mut predeclared_id = PreDeclaredID::False;
    let mut exposed = Exposed::False;
    let mut description: Option<String> = None;
    let mut ext_key: HashMap<String, String> = HashMap::new();

    // Find all AttributeStatement nodes
    let attr_statements: Vec<_> = cst
        .children()
        .into_iter()
        .filter(|c| c.kind == SyntaxKind::AttributeStatement)
        .collect();

    for attr_stmt in attr_statements {
        // Navigate through the child tokens of the AttributeStatement
        // Expected structure: AttributeKeyword, Whitespace, Identifier, Whitespace, EqualityOperator, Whitespace, Value, Newline

        let mut key = String::new();
        let mut value = String::new();
        let mut found_equals = false;

        for child in &attr_stmt.children {
            if !child.is_token {
                continue; // Skip non-token children
            }

            match child.kind {
                SyntaxKind::AttributeKeyword => {
                    // Skip the "Attribute" keyword
                    continue;
                }
                SyntaxKind::Identifier => {
                    if !found_equals {
                        // This is the attribute key (e.g., "VB_Name")
                        key = child.text.trim().to_string();
                    }
                }
                SyntaxKind::EqualityOperator => {
                    found_equals = true;
                }
                SyntaxKind::StringLiteral => {
                    if found_equals {
                        // This is the string value - remove surrounding quotes
                        value = child.text.trim().trim_matches('"').to_string();
                    }
                }
                SyntaxKind::TrueKeyword => {
                    if found_equals {
                        value = "True".to_string();
                    }
                }
                SyntaxKind::FalseKeyword => {
                    if found_equals {
                        value = "False".to_string();
                    }
                }
                SyntaxKind::IntegerLiteral | SyntaxKind::LongLiteral => {
                    if found_equals {
                        value = child.text.trim().to_string();
                    }
                }
                SyntaxKind::SubtractionOperator => {
                    if found_equals && value.is_empty() {
                        value.push('-');
                    }
                }
                _ => {}
            }
        }

        // Process the extracted key-value pair
        if !key.is_empty() {
            match key.as_str() {
                "VB_Name" => {
                    name = value;
                }
                "VB_GlobalNameSpace" => {
                    global_name_space = if value == "True" || value == "-1" {
                        NameSpace::Global
                    } else {
                        NameSpace::Local
                    };
                }
                "VB_Creatable" => {
                    creatable = if value == "True" || value == "-1" {
                        Creatable::True
                    } else {
                        Creatable::False
                    };
                }
                "VB_PredeclaredId" => {
                    predeclared_id = if value == "True" || value == "-1" {
                        PreDeclaredID::True
                    } else {
                        PreDeclaredID::False
                    };
                }
                "VB_Exposed" => {
                    exposed = if value == "True" || value == "-1" {
                        Exposed::True
                    } else {
                        Exposed::False
                    };
                }
                "VB_Description" => {
                    description = Some(value);
                }
                "VB_Ext_KEY" => {
                    // VB_Ext_KEY attributes have comma-separated values
                    // Format: VB_Ext_KEY = "key" ,"value"
                    // We need to parse the comma-separated string values
                    // For now, store the raw value
                    ext_key.insert(key.clone(), value);
                }
                _ => {
                    // Unknown attribute, could add to ext_key or ignore
                }
            }
        }
    }

    FileAttributes {
        name,
        global_name_space,
        creatable,
        predeclared_id,
        exposed,
        description,
        ext_key,
    }
}

/// Extract VB6ClassProperties from PropertiesBlock nodes in the CST
fn extract_properties(cst: &crate::parsers::ConcreteSyntaxTree) -> ClassProperties {
    use crate::parsers::class::properties::{
        DataBindingBehavior, DataSourceBehavior, FileUsage, MtsStatus, Persistence,
    };

    let mut multi_use = FileUsage::MultiUse;
    let mut persistable = Persistence::NotPersistable;
    let mut data_binding_behavior = DataBindingBehavior::None;
    let mut data_source_behavior = DataSourceBehavior::None;
    let mut mts_transaction_mode = MtsStatus::NotAnMTSObject;

    // Find the PropertiesBlock node
    let properties_blocks: Vec<_> = cst
        .children()
        .into_iter()
        .filter(|c| c.kind == SyntaxKind::PropertiesBlock)
        .collect();

    if properties_blocks.is_empty() {
        return ClassProperties::default();
    }

    let properties_block = &properties_blocks[0];

    // Find all Property nodes within the PropertiesBlock
    let property_nodes: Vec<_> = properties_block
        .children
        .iter()
        .filter(|c| c.kind == SyntaxKind::Property)
        .collect();

    for prop_node in property_nodes {
        let mut key = String::new();
        let mut value = String::new();
        let mut found_equals = false;

        for child in &prop_node.children {
            if !child.is_token {
                continue;
            }

            match child.kind {
                SyntaxKind::PropertyKey => {
                    // This is a nested node, get its text
                    if let Some(first_child) = child.children.first() {
                        key = first_child.text.trim().to_string();
                    }
                }
                SyntaxKind::EqualityOperator => {
                    found_equals = true;
                }
                SyntaxKind::PropertyValue => {
                    // This is a nested node, get all its text
                    if found_equals {
                        for val_child in &child.children {
                            if val_child.is_token {
                                match val_child.kind {
                                    SyntaxKind::IntegerLiteral | SyntaxKind::LongLiteral => {
                                        value.push_str(val_child.text.trim());
                                    }
                                    SyntaxKind::SubtractionOperator => {
                                        value.push('-');
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
        }

        // Process the extracted key-value pair
        if !key.is_empty() && !value.is_empty() {
            match key.as_str() {
                "MultiUse" => {
                    multi_use = if value == "-1" {
                        FileUsage::MultiUse
                    } else {
                        FileUsage::SingleUse
                    };
                }
                "Persistable" => {
                    persistable = if value == "-1" {
                        Persistence::Persistable
                    } else {
                        Persistence::NotPersistable
                    };
                }
                "DataBindingBehavior" => {
                    data_binding_behavior = match value.as_str() {
                        "0" => DataBindingBehavior::None,
                        "1" => DataBindingBehavior::Simple,
                        "2" => DataBindingBehavior::Complex,
                        _ => DataBindingBehavior::None,
                    };
                }
                "DataSourceBehavior" => {
                    data_source_behavior = match value.as_str() {
                        "0" => DataSourceBehavior::None,
                        "1" => DataSourceBehavior::DataSource,
                        _ => DataSourceBehavior::None,
                    };
                }
                "MTSTransactionMode" => {
                    mts_transaction_mode = match value.as_str() {
                        "0" => MtsStatus::NotAnMTSObject,
                        "1" => MtsStatus::NoTransactions,
                        "2" => MtsStatus::RequiresTransaction,
                        "3" => MtsStatus::UsesTransaction,
                        "4" => MtsStatus::RequiresNewTransaction,
                        _ => MtsStatus::NotAnMTSObject,
                    };
                }
                _ => {}
            }
        }
    }

    ClassProperties {
        multi_use,
        persistable,
        data_binding_behavior,
        data_source_behavior,
        mts_transaction_mode,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sourcefile;

    #[test]
    fn class_file_valid() {
        let class_bytes = r#"VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior = 0  'vbNone
    MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = \"Something\"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = \"SavedWithClassBuilder6\" ,\"Yes\"
Attribute VB_Ext_KEY = \"Saved\" ,\"False\"

Option Explicit
"#;

        let result = SourceFile::decode_with_replacement("test.cls", &class_bytes.as_bytes());

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let result = ClassFile::parse(&source_file);

        if result.has_failures() {
            for failure in result.failures {
                failure.print();
            }

            panic!("Class parse had failures");
        }

        assert!(result.has_result());
    }

    #[test]
    fn class_file_invalid() {
        // These should be '\r\n', or at worst '\n', but not '\r'.
        // '\r' is not even remotely a valid line ending.
        let class_bytes = b"VERSION 1.0 CLASS\r
                    BEGIN\r
                        MultiUse = -1  'True\r
                        Persistable = 0  'NotPersistable\r
                        DataBindingBehavior = 0  'vbNone\r
                        DataSourceBehavior = 0  'vbNone\r
                        MTSTransactionMode = 0  'NotAnMTSObject\r
                    END\r
                    Attribute VB_Name = \"Something\"\r
                    Attribute VB_GlobalNameSpace = False\r
                    Attribute VB_Creatable = True\r
                    Attribute VB_PredeclaredId = False\r
                    Attribute VB_Exposed = False\r
                    Attribute VB_Description = \"Description text\"\r
                    \r
                    Option Explicit\r";

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let result = ClassFile::parse(&source_file);

        assert!(result.has_failures());
    }

    #[test]
    fn class_header_valid() {
        let input = b"VERSION 1.0 CLASS\r
    BEGIN\r
        MultiUse = -1  'True\r
        Persistable = 0  'NotPersistable\r
        DataBindingBehavior = 0  'vbNone\r
        DataSourceBehavior = 0  'vbNone\r
        MTSTransactionMode = 0  'NotAnMTSObject\r
    END\r
    Attribute VB_Name = \"Something\"\r
    Attribute VB_GlobalNameSpace = False\r
    Attribute VB_Creatable = True\r
    Attribute VB_PredeclaredId = False\r
    Attribute VB_Exposed = False";

        let sourcefile =
            sourcefile::SourceFile::decode_with_replacement("test.cls", input).unwrap();

        let result = ClassFile::parse(&sourcefile);

        assert!(result.has_result());
    }

    #[test]
    fn version_valid() {
        let class_bytes = b"VERSION 1.0 CLASS\r\n";

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let mut source_stream = source_file.get_source_stream();
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        let version = extract_version(&cst);

        assert!(version.is_some());
        let version = version.unwrap();
        assert_eq!(version.major, 1);
        assert_eq!(version.minor, 0);
    }

    #[test]
    fn version_invalid() {
        // 'VERSION' isn't correct - this will fail to tokenize properly
        let class_bytes = b"VERION 1.0 CLASS";

        let result = SourceFile::decode_with_replacement("test.cls", class_bytes);

        let source_file = match result {
            Ok(source_file) => source_file,
            Err(e) => panic!("Failed to decode source file 'test.cls': {e:?}"),
        };

        let mut source_stream = source_file.get_source_stream();
        let token_stream = tokenize(&mut source_stream).unwrap();
        let cst = parse(token_stream);

        let version = extract_version(&cst);

        // Should be None because there's no VERSION keyword
        assert!(version.is_none());
    }
}
