//! VB6 Class File Properties Module
//!
//! This module defines the properties and attributes of a VB6 class file (.cls).
//! It includes enums and structs to represent various class file properties
//! such as COM usage, persistability, MTS status, data binding behavior, and more.
//!
//! These properties are typically found in the header of a VB6 class file
//! and are not normally visible in the code editor region. They are only
//! visible in the file property explorer.

use serde::Serialize;

use crate::parsers::header::{FileAttributes, FileFormatVersion};

/// Represents the COM usage of a class file.
/// Only available when the class is part of an `ActiveX` DLL project that is both
/// public and creatable.
///
/// Determines whether the class can be used by multiple clients or a single client.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum FileUsage {
    /// In a `COM` object a `SingleUse` class object will be created for each client.
    /// This value is stored as 0 (false) in the file.
    SingleUse = 0, // 0 (false)
    /// In a `COM` object a `MultiUse` class object will be created for all clients.
    /// This value is stored as -1 (true) in the file.
    #[default]
    MultiUse = -1,
}

/// Represents the persistability of a file.
///
/// Only available when the class is part of an `ActiveX` DLL project that is both
/// public and creatable.
///
/// Determines whether the class can be saved to disk.
///
/// If it is `Persistable`, then four procedures: `InitProperties`, `ReadProperties`, and
/// `WriteProperties` events, and the `PropertyChanged` method are automatically
/// added to the class module.
///
/// Without these procedures, the class cannot be saved to disk.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum Persistence {
    /// The class property cannot be saved to a file in a property bag.
    /// This value is stored as 0 (false) in the file.
    #[default]
    NotPersistable = 0,
    /// The class property can be saved to a file in a property bag.
    /// This value is stored as -1 (true) in the file.
    Persistable = -1,
}

/// Represents the MTS status of a file.
///
/// Only available when the class is part of an activeX DLL project. This should
/// be set to values other than `NotAnMTSObject` (0) if the class is to be used as
/// a Microsoft Transaction Server component.
///
/// Maps directly to the MTS transaction mode attribute in Microsoft Transaction
/// Server.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum MtsStatus {
    /// This class is not an MTS component.
    /// This value is stored as 0 in the file.
    /// This is the default value.
    #[default]
    NotAnMTSObject = 0,
    /// This class is an MTS component but does not support transactions.
    /// This value is stored as 1 in the file.
    NoTransactions = 1,
    /// This class is an MTS component and requires a transaction.
    /// This value is stored as 2 in the file.
    RequiresTransaction = 2,
    /// This class is an MTS component and uses a transaction.
    /// This value is stored as 3 in the file.
    UsesTransaction = 3,
    /// This class is an MTS component and requires a new transaction.
    /// This value is stored as 4 in the file.
    RequiresNewTransaction = 4,
}

/// Determines if a class can act as a `DataSource` for VB6 `DataBinding`.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum DataSourceBehavior {
    /// The class does not support acting as a `DataSource`.
    /// This value is stored as 0 in the file.
    #[default]
    None = 0,
    /// The class supports acting as a `DataSource`.
    /// This value is stored as 1 in the file.
    DataSource = 1,
}

/// Determines the default VB6 `DataBinding` behavior.
///
/// Only available when the class is part of an `ActiveX` DLL project that is both
/// public and creatable.
///
/// Used to specify whether the class supports `DataBinding` and the level of
/// `DataBinding` support.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub enum DataBindingBehavior {
    /// The class does not support `DataBinding`.
    /// This value is stored as 0 in the file.
    #[default]
    None = 0,
    /// The class supports simple `DataBinding`.
    /// This value is stored as 1 in the file.
    Simple = 1,
    /// The class supports complex `DataBinding`.
    /// This value is stored as 2 in the file.
    Complex = 2,
}

/// The properties of a VB6 class file is the list of key/value pairs
/// found between the `BEGIN` and `END` lines in the header.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone, Serialize, Default)]
pub struct ClassProperties {
    /// The COM usage of the class file.
    pub multi_use: FileUsage,
    /// The persistability of the class file.
    pub persistable: Persistence,
    /// The data binding behavior of the class file.
    pub data_binding_behavior: DataBindingBehavior,
    /// The data source behavior of the class file.
    pub data_source_behavior: DataSourceBehavior,
    /// The MTS transaction mode of the class file.
    pub mts_transaction_mode: MtsStatus,
}

/// Represents the header of a VB6 class file.
///
/// The header contains the version, multi use, persistable, data binding behavior,
/// data source behavior, and MTS transaction mode.
/// The header also contains the attributes of the class file.
///
/// None of these values are normally visible in the code editor region.
/// They are only visible in the file property explorer.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct ClassHeader {
    /// The version of the VB6 file format.
    pub version: FileFormatVersion,
    /// The properties of the class file.
    pub properties: ClassProperties,
    /// The attributes of the class file.
    pub attributes: FileAttributes,
}
