
# Goals:

VB6Parse aims to be a complete, end-to-end parser library for VB6. Including, but not limited to:

* (*.vbp) VB6 project files.
* (*.vbw) VB6 windows files for determining editor windows when they are opened. 
* (*.frm) VB6 Forms.
* (*.frx) VB6 Form Resource files.
* (*.dsx) VB Data Environment files.
* (*.dsr) VB Data Environment Resource files.
* (*.cls) Class files.
* (*.ttx) Crystal Report files.
* (*.ctl) User Control files.
* (*.dob) User Document files.

## Current support:

First work has focused on the (vbp) project files since is the method to discover all other files that should be linked/referenced within a project.

#### (*.vbp) VB6 Project file parser feature support:
- [X] **Project Types**
    - [X] Exe
    - [X] Control
    - [X] OleExe
    - [X] OleDll
- [x] **References**
- [x] **Objects**
- [x] **Modules**
- [x] **Designers**
- [x] **Classes**
- [x] **Forms**
- [x] **UserControls**
- [X] **UserDocuments**
- [X] **ResFile32** - Partial support. Default value not correctly handled.
- [X] **IconForm** - Partial support. Default value not correctly handled.
- [X] **Startup** - Partial support. Default value not correctly handled.
- [X] **HelpFile** - Partial support. Default value not correctly handled.
- [X] **Title** - Partial support. Default value not correctly handled. 
- [X] **ExeName32** - Partial support. Default value not correctly handled. 
- [X] **Command32** - Partial support. Default value not correctly handled. 
- [X] **Name** - Partial support. Default value not correctly handled. 
- [X] **HelpContextID** - Partial support. Default value not correctly handled. 
- [X] **CompatibleMode** - Partial support. Default value not correctly handled. 
- [X] **NoControlUpgrade** - Full support for the 'ActiveX Control Upgrade' option, including the default or empty reverting to true.
- [ ] **MajorVer**
- [ ] **MinorVer**
- [ ] **RevisionVer**
- [ ] **AutoIncrementVer**
- [ ] **ServerSupportFiles**
- [ ] **VersionCompanyName**
- [ ] **VersionFileDescription**
- [ ] **VersionLegalCopyright**
- [ ] **VersionLegalTrademarks**
- [ ] **VersionProductName**
- [ ] **CondComp**
- [ ] **CompilationType**
- [ ] **OptimizationType**
- [ ] **NoAliasing**
- [ ] **CodeViewDebugInfo**
- [ ] **FavorPentiumPro(tm)** - Yes, this is exactly what this looks like inside the project file, '(tm)' and all.
- [ ] **BoundsCheck**
- [ ] **OverflowCheck**
- [ ] **FlPointCheck**
- [ ] **FDIVCheck**
- [ ] **UnroundedFP**
- [ ] **StartMode**
- [ ] **Unattended**
- [ ] **Retained**
- [ ] **ThreadPerObject**
- [ ] **MaxNumberOfThreads**
- [ ] **DebugStartOption**
- [ ] **AutoRefresh**

#### VB6Project API:
- [x] Unit Testing.
- [x] Integration/End-to-End Testing.
- [ ] Benchmarking.
- [ ] Top level API finalization.
- [ ] Documentation.
- [ ] Examples.