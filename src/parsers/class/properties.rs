use serde::Serialize;

use crate::parsers::header::{VB6FileAttributes, VB6FileFormatVersion};

/// Represents the COM usage of a class file.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum FileUsage {
    // In a COM object a MultiUse class object will be created for all clients.
    // This value is stored as -1 (true) in the file.
    MultiUse = -1,
    // In a COM object a SingleUse class object will be created for each client.
    // This value is stored as 0 (false) in the file.
    SingleUse = 0, // 0 (false)
}

/// Represents the persistability of a file.
///
/// Only available when the class is part of an activeX DLL project that is both
/// public and creatable.
///
/// Determines whether the class can be saved to disk.
///
/// If it is `Persistable`, then four procedures: `InitProperties`, `ReadProperties`, and
/// `WriteProperties` events, and the `PropertyChanged` method are automatically
/// added to the class module.
///
/// Without these procedures, the class cannot be saved to disk.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum Persistance {
    // The class property cannot be saved to a file in a property bag.
    // This value is stored as 0 (false) in the file.
    NonPersistable = 0,
    // The class property can be saved to a file in a property bag.
    // This value is stored as -1 (true) in the file.
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
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum MtsStatus {
    // This class is not an MTS component.
    // This value is stored as 0 in the file.
    // This is the default value.
    NotAnMTSObject = 0,
    // This class is an MTS component but does not support transactions.
    // This value is stored as 1 in the file.
    NoTransactions = 1,
    // This class is an MTS component and requires a transaction.
    // This value is stored as 2 in the file.
    RequiresTransaction = 2,
    // This class is an MTS component and uses a transaction.
    // This value is stored as 3 in the file.
    UsesTransaction = 3,
    // This class is an MTS component and requires a new transaction.
    // This value is stored as 4 in the file.
    RequiresNewTransaction = 4,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum DataSourceBehavior {
    // The class does not support acting as a Data Source.
    // This value is stored as 0 in the file.
    None = 0,
    // The class supports acting as a Data Source.
    // This value is stored as 1 in the file.
    DataSource = 1,
}

#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub enum DataBindingBehavior {
    // The class does not support data binding.
    // This value is stored as 0 in the file.
    None = 0,
    // The class supports simple data binding.
    // This value is stored as 1 in the file.
    Simple = 1,
    // The class supports complex data binding.
    // This value is stored as 2 in the file.
    Complex = 2,
}

/// The properties of a VB6 class file is the list of key/value pairs
/// found between the BEGIN and END lines in the header.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ClassProperties {
    // (0/-1) multi use / single use
    pub multi_use: FileUsage,
    // (0/1) NonParsistable / Persistable
    pub persistable: Persistance,
    // (0/1/2) vbNone / vbSimple / vbComplex
    pub data_binding_behavior: DataBindingBehavior,
    // (0/1) vbNone / vbDataSource
    pub data_source_behavior: DataSourceBehavior,
    // (0/1/2/3/4) NotAnMTSObject / NoTransactions / RequiresTransaction / UsesTransaction / RequiresNewTransaction
    pub mts_transaction_mode: MtsStatus,
}

/// Represents the version of a VB6 class file.
/// The class version contains a major and minor version number.
#[derive(Debug, PartialEq, Eq, Clone, Serialize)]
pub struct VB6ClassVersion {
    pub major: u8,
    pub minor: u8,
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
pub struct VB6ClassHeader<'a> {
    pub version: VB6FileFormatVersion,
    pub properties: VB6ClassProperties,
    pub attributes: VB6FileAttributes<'a>,
}
