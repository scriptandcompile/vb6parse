# VB6Parse WebAssembly Demo

## Overview

This document outlines the requirements and implementation strategy for compiling VB6Parse to WebAssembly (WASM) and creating an interactive web demo similar to the Rust Playground.

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      Browser Frontend                        │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │ Code Editor  │  │ Output Panel │  │ Options Menu │      │
│  │ (Monaco/CM)  │  │  (Results)   │  │ (File Type)  │      │
│  └──────┬───────┘  └──────▲───────┘  └──────────────┘      │
│         │                  │                                 │
│         └──────────┬───────┘                                 │
│                    │                                         │
│         ┌──────────▼───────────┐                            │
│         │   JavaScript Glue    │                            │
│         │   (WASM Bindings)    │                            │
│         └──────────┬───────────┘                            │
│                    │                                         │
│         ┌──────────▼───────────┐                            │
│         │  VB6Parse WASM       │                            │
│         │  (Compiled Rust)     │                            │
│         └──────────────────────┘                            │
└─────────────────────────────────────────────────────────────┘
```

## Technical Requirements

### 1. Build Tools

- **wasm-pack**: Tool for building Rust-generated WebAssembly packages
  ```bash
  cargo install wasm-pack
  ```

- **wasm-bindgen**: Generate JavaScript bindings for Rust
  ```bash
  # Already included as dependency
  ```

### 2. Cargo Configuration

Add to `Cargo.toml`:

```toml
[lib]
crate-type = ["cdylib", "rlib"]

[dependencies]
# Existing dependencies remain as-is
# No WASM dependencies in the main section

[target.'cfg(target_arch = "wasm32")'.dependencies]
wasm-bindgen = "0.2"
serde = { version = "1.0", features = ["derive"] }
serde-wasm-bindgen = "0.6"
console_error_panic_hook = "0.1"
wee_alloc = "0.4"
getrandom = { version = "0.2", features = ["js"] }

[profile.release]
opt-level = "z"     # Optimize for size
lto = true          # Enable link-time optimization
codegen-units = 1   # Reduce parallel codegen for smaller binary
```

### 3. WASM Module Interface

Create `src/wasm.rs`:

```rust
use wasm_bindgen::prelude::*;
use serde::{Deserialize, Serialize};

#[wasm_bindgen]
pub fn init_panic_hook() {
    console_error_panic_hook::set_once();
}

#[derive(Serialize, Deserialize)]
pub struct ParseResult {
    success: bool,
    errors: Vec<ParseError>,
    cst_json: Option<String>,
    output: String,
}

#[derive(Serialize, Deserialize)]
pub struct ParseError {
    line: usize,
    column: usize,
    message: String,
}

#[wasm_bindgen]
pub fn parse_vb6_code(code: &str, file_type: &str) -> JsValue {
    use crate::io::SourceFile;
    
    let result = match file_type {
        "class" => {
            let source = SourceFile::decode_with_replacement(code.as_bytes(), "input.cls");
            let (parsed, failures) = crate::files::ClassFile::parse(&source).unpack();
            
            ParseResult {
                success: parsed.is_some(),
                errors: failures.iter().map(|e| ParseError {
                    line: e.span.start_position().line,
                    column: e.span.start_position().column,
                    message: format!("{:?}", e.error),
                }).collect(),
                cst_json: parsed.as_ref().map(|p| format!("{:#?}", p.code())),
                output: format!("{:#?}", parsed),
            }
        },
        "module" => {
            let source = SourceFile::decode_with_replacement(code.as_bytes(), "input.bas");
            let (parsed, failures) = crate::files::ModuleFile::parse(&source).unpack();
            
            ParseResult {
                success: parsed.is_some(),
                errors: failures.iter().map(|e| ParseError {
                    line: e.span.start_position().line,
                    column: e.span.start_position().column,
                    message: format!("{:?}", e.error),
                }).collect(),
                cst_json: parsed.as_ref().map(|p| format!("{:#?}", p.code())),
                output: format!("{:#?}", parsed),
            }
        },
        "form" => {
            let source = SourceFile::decode_with_replacement(code.as_bytes(), "input.frm");
            let (parsed, failures) = crate::files::FormFile::parse(&source).unpack();
            
            ParseResult {
                success: parsed.is_some(),
                errors: failures.iter().map(|e| ParseError {
                    line: e.span.start_position().line,
                    column: e.span.start_position().column,
                    message: format!("{:?}", e.error),
                }).collect(),
                cst_json: parsed.as_ref().map(|p| format!("{:#?}", p.code())),
                output: format!("{:#?}", parsed),
            }
        },
        "project" => {
            let source = SourceFile::decode_with_replacement(code.as_bytes(), "input.vbp");
            let (parsed, failures) = crate::files::ProjectFile::parse(&source).unpack();
            
            ParseResult {
                success: parsed.is_some(),
                errors: failures.iter().map(|e| ParseError {
                    line: e.span.start_position().line,
                    column: e.span.start_position().column,
                    message: format!("{:?}", e.error),
                }).collect(),
                cst_json: None,
                output: format!("{:#?}", parsed),
            }
        },
        _ => ParseResult {
            success: false,
            errors: vec![ParseError {
                line: 0,
                column: 0,
                message: format!("Unknown file type: {}", file_type),
            }],
            cst_json: None,
            output: String::new(),
        },
    };
    
    serde_wasm_bindgen::to_value(&result).unwrap()
}

#[wasm_bindgen]
pub fn tokenize_vb6_code(code: &str) -> JsValue {
    use crate::io::{SourceFile, SourceStream};
    use crate::lexer::tokenize;
    
    let source = SourceFile::decode_with_replacement(code.as_bytes(), "input.vb");
    let stream = SourceStream::new(&source);
    let (tokens, failures) = tokenize(&stream).unpack();
    
    let result = serde_json::json!({
        "tokens": tokens.map(|t| {
            t.tokens().map(|tok| {
                serde_json::json!({
                    "kind": format!("{:?}", tok.kind()),
                    "text": tok.text().to_string(),
                    "line": tok.span().start_position().line,
                    "column": tok.span().start_position().column,
                })
            }).collect::<Vec<_>>()
        }),
        "errors": failures.iter().map(|e| {
            serde_json::json!({
                "line": e.span.start_position().line,
                "column": e.span.start_position().column,
                "message": format!("{:?}", e.error),
            })
        }).collect::<Vec<_>>(),
    });
    
    serde_wasm_bindgen::to_value(&result).unwrap()
}
```

Add to `src/lib.rs`:

```rust
#[cfg(target_arch = "wasm32")]
pub mod wasm;
```

## Build Process

### 1. Build WASM Package

```bash
# Build for web (generates JS + WASM)
wasm-pack build --target web --out-dir docs/wasm

# Or build for bundler (webpack, rollup)
wasm-pack build --target bundler --out-dir docs/wasm

# For production (optimized)
wasm-pack build --release --target web --out-dir docs/wasm
```

This generates:
- `vb6parse_bg.wasm` - The WebAssembly binary
- `vb6parse.js` - JavaScript bindings
- `vb6parse.d.ts` - TypeScript definitions
- `package.json` - NPM package metadata

### 2. Optimize WASM Size

```bash
# Install wasm-opt
cargo install wasm-opt

# Optimize the WASM binary
wasm-opt -Oz -o docs/wasm/vb6parse_bg_opt.wasm docs/wasm/vb6parse_bg.wasm

# Check size reduction
ls -lh docs/wasm/*.wasm
```

## Web Frontend Implementation

### HTML Structure (`playground.html`)

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>VB6Parse Playground</title>
    <link rel="stylesheet" href="playground.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/codemirror.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.2/theme/monokai.min.css">
</head>
<body>
    <div class="playground-container">
        <header>
            <h1>🎮 VB6Parse Playground</h1>
            <div class="controls">
                <select id="file-type">
                    <option value="module">Module (.bas)</option>
                    <option value="class">Class (.cls)</option>
                    <option value="form">Form (.frm)</option>
                    <option value="project">Project (.vbp)</option>
                </select>
                <select id="view-mode">
                    <option value="parsed">Parsed Output</option>
                    <option value="tokens">Tokens</option>
                    <option value="cst">CST Tree</option>
                </select>
                <button id="parse-btn">Parse</button>
                <button id="share-btn">Share</button>
                <button id="examples-btn">Examples</button>
            </div>
        </header>
        
        <div class="editor-container">
            <div class="editor-pane">
                <div class="pane-header">
                    <span>Input Code</span>
                    <button id="clear-btn">Clear</button>
                </div>
                <textarea id="code-editor"></textarea>
            </div>
            
            <div class="output-pane">
                <div class="pane-header">
                    <span>Output</span>
                    <button id="copy-btn">Copy</button>
                </div>
                <div id="output-container">
                    <div id="errors-section" class="hidden"></div>
                    <pre id="output-content"></pre>
                </div>
            </div>
        </div>
        
        <div id="status-bar">
            <span id="status-text">Ready</span>
            <span id="parse-time"></span>
        </div>
    </div>
    
    <script type="module" src="playground.js"></script>
</body>
</html>
```

### JavaScript Integration (`playground.js`)

```javascript
import init, { init_panic_hook, parse_vb6_code, tokenize_vb6_code } from './wasm/vb6parse.js';

let editor;
let wasmLoaded = false;

// Initialize WASM module
async function initWasm() {
    try {
        await init();
        init_panic_hook();
        wasmLoaded = true;
        updateStatus('WASM module loaded successfully', 'success');
    } catch (err) {
        console.error('Failed to load WASM:', err);
        updateStatus('Failed to load WASM module', 'error');
    }
}

// Initialize CodeMirror editor
function initEditor() {
    editor = CodeMirror.fromTextArea(document.getElementById('code-editor'), {
        mode: 'vb',
        theme: 'monokai',
        lineNumbers: true,
        indentUnit: 4,
        tabSize: 4,
        indentWithTabs: true,
        lineWrapping: true,
        matchBrackets: true,
        autoCloseBrackets: true,
    });
    
    // Set default example
    editor.setValue(getDefaultExample('module'));
}

// Parse button handler
document.getElementById('parse-btn').addEventListener('click', async () => {
    if (!wasmLoaded) {
        updateStatus('WASM module not loaded yet', 'error');
        return;
    }
    
    const code = editor.getValue();
    const fileType = document.getElementById('file-type').value;
    const viewMode = document.getElementById('view-mode').value;
    
    if (!code.trim()) {
        updateStatus('Please enter some code', 'warning');
        return;
    }
    
    updateStatus('Parsing...', 'info');
    const startTime = performance.now();
    
    try {
        let result;
        
        if (viewMode === 'tokens') {
            result = tokenize_vb6_code(code);
            displayTokens(result);
        } else {
            result = parse_vb6_code(code, fileType);
            
            if (viewMode === 'cst' && result.cst_json) {
                displayOutput(result.cst_json);
            } else {
                displayOutput(result.output);
            }
            
            if (result.errors && result.errors.length > 0) {
                displayErrors(result.errors);
            } else {
                hideErrors();
            }
        }
        
        const duration = (performance.now() - startTime).toFixed(2);
        updateStatus(`Parsed successfully in ${duration}ms`, 'success');
        document.getElementById('parse-time').textContent = `${duration}ms`;
        
    } catch (err) {
        console.error('Parse error:', err);
        updateStatus(`Error: ${err.message}`, 'error');
    }
});

// Display functions
function displayOutput(text) {
    const output = document.getElementById('output-content');
    output.textContent = text;
    hljs.highlightElement(output);
}

function displayTokens(result) {
    const output = document.getElementById('output-content');
    
    if (result.tokens && result.tokens.length > 0) {
        const tokenList = result.tokens.map(tok => 
            `<div class="token">
                <span class="token-kind">${tok.kind}</span>
                <span class="token-text">"${escapeHtml(tok.text)}"</span>
                <span class="token-pos">(${tok.line}:${tok.column})</span>
            </div>`
        ).join('');
        
        output.innerHTML = tokenList;
    } else {
        output.textContent = 'No tokens found';
    }
    
    if (result.errors && result.errors.length > 0) {
        displayErrors(result.errors);
    }
}

function displayErrors(errors) {
    const errorsSection = document.getElementById('errors-section');
    errorsSection.classList.remove('hidden');
    
    errorsSection.innerHTML = '<h3>Errors:</h3>' + errors.map(err => 
        `<div class="error-item">
            <span class="error-location">[${err.line}:${err.column}]</span>
            <span class="error-message">${escapeHtml(err.message)}</span>
        </div>`
    ).join('');
}

function hideErrors() {
    document.getElementById('errors-section').classList.add('hidden');
}

function updateStatus(message, type) {
    const statusText = document.getElementById('status-text');
    statusText.textContent = message;
    statusText.className = type;
}

// Example code snippets
function getDefaultExample(fileType) {
    const examples = {
        module: `Attribute VB_Name = "Module1"
Option Explicit

Public Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function

Private Sub Main()
    Dim result As Integer
    result = AddNumbers(5, 10)
    MsgBox "Result: " & result
End Sub`,
        
        class: `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "Calculator"

Private m_value As Double

Public Property Get Value() As Double
    Value = m_value
End Property

Public Sub Add(num As Double)
    m_value = m_value + num
End Sub

Public Sub Clear()
    m_value = 0
End Sub`,
        
        form: `VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "My Application"
   ClientHeight    =   3090
   ClientWidth     =   4560
   Begin VB.CommandButton btnSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "MainForm"

Private Sub btnSubmit_Click()
    MsgBox "Button clicked!"
End Sub`,
        
        project: `Type=Exe
Form=Form1.frm
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation
Module=Module1; Module1.bas
Class=Calculator; Calculator.cls
IconForm="Form1"
Startup="Form1"
ExeName32="MyApp.exe"
Command32=""
Name="Project1"
HelpContextID="0"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionCompanyName="Company"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck=0
OverflowCheck=0
FlPointCheck=0
FDIVCheck=0
UnroundedFP=0
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1`
    };
    
    return examples[fileType] || examples.module;
}

// File type change handler
document.getElementById('file-type').addEventListener('change', (e) => {
    editor.setValue(getDefaultExample(e.target.value));
});

// Share functionality (URL encoding)
document.getElementById('share-btn').addEventListener('click', () => {
    const code = editor.getValue();
    const fileType = document.getElementById('file-type').value;
    const encoded = btoa(encodeURIComponent(code));
    const url = `${window.location.origin}${window.location.pathname}?code=${encoded}&type=${fileType}`;
    
    navigator.clipboard.writeText(url).then(() => {
        updateStatus('Link copied to clipboard!', 'success');
    });
});

// Load shared code from URL
function loadFromUrl() {
    const params = new URLSearchParams(window.location.search);
    const encoded = params.get('code');
    const fileType = params.get('type');
    
    if (encoded) {
        try {
            const decoded = decodeURIComponent(atob(encoded));
            editor.setValue(decoded);
            if (fileType) {
                document.getElementById('file-type').value = fileType;
            }
        } catch (err) {
            console.error('Failed to load shared code:', err);
        }
    }
}

// Utility functions
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Initialize on load
window.addEventListener('DOMContentLoaded', async () => {
    initEditor();
    await initWasm();
    loadFromUrl();
});
```

### CSS Styling (`playground.css`)

```css
:root {
    --bg-primary: #1e1e1e;
    --bg-secondary: #252526;
    --bg-tertiary: #2d2d30;
    --text-primary: #d4d4d4;
    --text-secondary: #9d9d9d;
    --border-color: #3e3e42;
    --accent-color: #007acc;
    --success-color: #4caf50;
    --error-color: #f44336;
    --warning-color: #ff9800;
}

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    height: 100vh;
    overflow: hidden;
}

.playground-container {
    display: flex;
    flex-direction: column;
    height: 100vh;
}

header {
    background: var(--bg-secondary);
    padding: 1rem;
    border-bottom: 1px solid var(--border-color);
    display: flex;
    justify-content: space-between;
    align-items: center;
}

header h1 {
    font-size: 1.5rem;
    font-weight: 600;
}

.controls {
    display: flex;
    gap: 0.5rem;
}

select, button {
    padding: 0.5rem 1rem;
    background: var(--bg-tertiary);
    color: var(--text-primary);
    border: 1px solid var(--border-color);
    border-radius: 4px;
    cursor: pointer;
    font-size: 0.9rem;
}

button:hover {
    background: var(--accent-color);
}

.editor-container {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1px;
    flex: 1;
    overflow: hidden;
    background: var(--border-color);
}

.editor-pane, .output-pane {
    display: flex;
    flex-direction: column;
    background: var(--bg-secondary);
    overflow: hidden;
}

.pane-header {
    padding: 0.75rem 1rem;
    background: var(--bg-tertiary);
    border-bottom: 1px solid var(--border-color);
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-weight: 600;
}

.CodeMirror {
    height: 100% !important;
    font-size: 14px;
    font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
}

#output-container {
    flex: 1;
    overflow: auto;
    padding: 1rem;
}

#output-content {
    white-space: pre-wrap;
    font-family: 'Consolas', 'Monaco', monospace;
    font-size: 13px;
    line-height: 1.5;
}

#errors-section {
    background: rgba(244, 67, 54, 0.1);
    border-left: 3px solid var(--error-color);
    padding: 1rem;
    margin-bottom: 1rem;
}

#errors-section.hidden {
    display: none;
}

.error-item {
    margin: 0.5rem 0;
    font-family: monospace;
}

.error-location {
    color: var(--error-color);
    font-weight: bold;
    margin-right: 0.5rem;
}

.token {
    padding: 0.25rem 0;
    font-family: monospace;
    font-size: 13px;
}

.token-kind {
    color: var(--accent-color);
    font-weight: bold;
    width: 150px;
    display: inline-block;
}

.token-text {
    color: var(--success-color);
}

.token-pos {
    color: var(--text-secondary);
    margin-left: 1rem;
}

#status-bar {
    background: var(--bg-tertiary);
    padding: 0.5rem 1rem;
    border-top: 1px solid var(--border-color);
    display: flex;
    justify-content: space-between;
    font-size: 0.9rem;
}

#status-text.success { color: var(--success-color); }
#status-text.error { color: var(--error-color); }
#status-text.warning { color: var(--warning-color); }
#status-text.info { color: var(--accent-color); }

@media (max-width: 768px) {
    .editor-container {
        grid-template-columns: 1fr;
        grid-template-rows: 1fr 1fr;
    }
}
```

## Deployment

### GitHub Pages

Add to repository root `.github/workflows/deploy-playground.yml`:

```yaml
name: Deploy Playground

on:
  push:
    branches: [ main ]
  workflow_dispatch:

jobs:
  build-and-deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Install Rust
        uses: actions-rs/toolchain@v1
        with:
          toolchain: stable
          target: wasm32-unknown-unknown
      
      - name: Install wasm-pack
        run: curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh
      
      - name: Build WASM
        run: wasm-pack build --release --target web --out-dir docs/wasm
      
      - name: Deploy to GitHub Pages
        uses: peaceiris/actions-gh-pages@v3
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./docs
```

## Performance Considerations

1. **WASM Binary Size**:
   - Current estimate: 200-400KB (optimized)
   - Use `wasm-opt -Oz` for maximum compression
   - Enable gzip compression on server (reduces to ~50-100KB)

2. **Initial Load Time**:
   - WASM compilation: ~50-200ms
   - Show loading indicator during init
   - Cache WASM module in browser

3. **Parse Performance**:
   - Expected: 1-10ms for typical files
   - 100-500ms for very large files (10,000+ lines)
   - Consider web worker for large files

4. **Memory Usage**:
   - Typical: 2-10MB for parser
   - Grows with input size
   - Monitor with `performance.memory`

## Limitations & Challenges

### Current Limitations

1. **No File System Access**: Form resource files (`.frx`) cannot be loaded
2. **Windows-1252 Encoding**: Limited support in browser
3. **No Multi-file Projects**: Can't parse multiple dependent files
4. **Memory Constraints**: Large projects may hit browser limits

### Technical Challenges

1. **Error Reporting**: Need to serialize error types to JavaScript
2. **CST Visualization**: Complex tree structure requires custom rendering
3. **Syntax Highlighting**: VB6 mode for CodeMirror needs custom definition
4. **Browser Compatibility**: Older browsers may not support WASM

### Future Enhancements

1. **Multi-file Support**: Upload multiple files, parse project dependencies
2. **Visual Form Designer**: Render form layouts from `.frm` files
3. **Export Options**: Download parsed AST as JSON, convert to other formats
4. **Collaborative Editing**: Share and edit code in real-time
5. **Analysis Tools**: Show complexity metrics, dead code detection
6. **Diff View**: Compare before/after parsing results

## Testing

### Unit Tests for WASM

```rust
#[cfg(all(test, target_arch = "wasm32"))]
mod wasm_tests {
    use super::*;
    use wasm_bindgen_test::*;
    
    #[wasm_bindgen_test]
    fn test_parse_simple_module() {
        let code = r#"
            Attribute VB_Name = "Module1"
            
            Sub Main()
                MsgBox "Hello"
            End Sub
        "#;
        
        let result = parse_vb6_code(code, "module");
        // Assert on result
    }
}
```

Run with:
```bash
wasm-pack test --headless --firefox
wasm-pack test --headless --chrome
```

## Example Repository Structure

```
vb6parse/
├── src/
│   ├── lib.rs
│   ├── wasm.rs          # New WASM bindings
│   └── ...
├── docs/
│   ├── playground.html  # New playground page
│   ├── playground.js    # New JS integration
│   ├── playground.css   # New styles
│   ├── wasm/            # Generated by wasm-pack
│   │   ├── vb6parse.js
│   │   ├── vb6parse_bg.wasm
│   │   └── ...
│   └── ...
├── Cargo.toml           # Updated with WASM deps
└── .github/
    └── workflows/
        └── deploy-playground.yml
```

## Getting Started Checklist

- [ ] Add wasm-bindgen dependencies to Cargo.toml
- [ ] Create src/wasm.rs with WASM bindings
- [ ] Install wasm-pack: `cargo install wasm-pack`
- [ ] Build WASM: `wasm-pack build --target web`
- [ ] Create playground.html frontend
- [ ] Create playground.js integration
- [ ] Create playground.css styles
- [ ] Test locally with simple HTTP server
- [ ] Optimize WASM binary size
- [ ] Set up GitHub Pages deployment
- [ ] Add examples and documentation
- [ ] Test cross-browser compatibility

## Resources

- [wasm-bindgen Book](https://rustwasm.github.io/wasm-bindgen/)
- [Rust and WebAssembly](https://rustwasm.github.io/book/)
- [wasm-pack Documentation](https://rustwasm.github.io/wasm-pack/)
- [CodeMirror Documentation](https://codemirror.net/doc/manual.html)
- [Monaco Editor](https://microsoft.github.io/monaco-editor/) (alternative to CodeMirror)
