
# Goals:

VB6Parse aims to be a complete, end-to-end parser library for VB6. Including, but not limited to:

* (*.vbp) VB6 project files.
* (*.bas) VB6 module files.
* (*.vbw) VB6 windows files for determining editor windows when they are opened. 
* (*.frm) VB6 Forms.
* (*.frx) VB6 Form Resource files.
* (*.dsx) VB6 Data Environment files.
* (*.dsr) VB6 Data Environment Resource files.
* (*.cls) VB6 Class files.
* (*.ttx) Crystal Report files.
* (*.ctl) User Control files.
* (*.dob) User Document files.

## Current support:

First work has focused on the (vbp) project files since is the method to discover all other files that should be linked/referenced within a project.

<details>
    <summary> (*.vbp) VB6 Project file parser feature support: </summary>

- [x] **Project Types**
    - [x] Exe
    - [x] Control
    - [x] OleExe
    - [x] OleDll
- [x] **References**
- [x] **Objects**
- [x] **Modules**
- [x] **Designers**
- [x] **Classes**
- [x] **Forms**
- [x] **UserControls**
- [x] **UserDocuments**
- [x] **ResFile32** - Partial support. Default value not correctly handled.
- [x] **IconForm** - Partial support. Default value not correctly handled.
- [x] **Startup** - Partial support. Default value not correctly handled.
- [x] **HelpFile** - Partial support. Default value not correctly handled.
- [x] **Title** - Partial support. Default value not correctly handled. 
- [x] **ExeName32** - Partial support. Default value not correctly handled. 
- [x] **Command32** - Partial support. Default value not correctly handled. 
- [x] **Name** - Partial support. Default value not correctly handled. 
- [x] **HelpContextID** - Partial support. Default value not correctly handled. 
- [x] **CompatibleMode** - Partial support. Default value not correctly handled. 
- [x] **NoControlUpgrade** - Full support for the 'ActiveX Control Upgrade' option, including the default or empty reverting to true.
- [x] **MajorVer** - Partial support. Default value not correctly handled.
- [x] **MinorVer** - Partial support. Default value not correctly handled.
- [x] **RevisionVer** - Partial support. Default value not correctly handled.
- [x] **AutoIncrementVer** - Partial support. Default value not correctly handled.
- [x] **ServerSupportFiles**
- [x] **VersionCompanyName**
- [x] **VersionFileDescription**
- [x] **VersionLegalCopyright**
- [x] **VersionLegalTrademarks**
- [x] **VersionProductName**
- [x] **CondComp**
- [x] **CompilationType**
- [x] **OptimizationType**
- [x] **NoAliasing**
- [x] **CodeViewDebugInfo**
- [x] **FavorPentiumPro(tm)** - Yes, this is exactly what this looks like inside the project file, '(tm)' and all.
- [x] **BoundsCheck**
- [x] **OverflowCheck**
- [x] **FlPointCheck**
- [x] **FDIVCheck**
- [x] **UnroundedFP**
- [x] **StartMode**
- [x] **Unattended**
- [x] **Retained**
- [x] **ThreadPerObject**
- [x] **MaxNumberOfThreads**
- [x] **DebugStartOption**
- [x] **AutoRefresh**

</details>

<details>
    <summary> (*.cls) VB6 Class file parser feature support: </summary>

- [x] **Header**
- [x] **VB6 Token stream lexed**

</details>

<details>
    <summary> (*.bas) VB6 module file parser feature support: </summary>

- [x] **Header**
- [x] **VB6 Token stream lexed**

</details>


#### VB6Project API:
- [x] Unit Testing (partial).
- [x] Integration/End-to-End Testing.
- [x] Benchmarking.
- [ ] Top level API finalization.
- [x] Documentation.
- [x] Examples.


### Test:

Be sure to use ```git submodule update --init --recursive``` to get all integration test submodule data.