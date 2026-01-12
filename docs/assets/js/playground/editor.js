/**
 * VB6Parse Playground - Editor Module
 * 
 * Handles Monaco Editor initialization, configuration, and VB6 syntax highlighting.
 * 
 * TODO: Implement VB6 language definition for Monaco
 * TODO: Add editor configuration options (font size, theme, etc.)
 * TODO: Implement localStorage for editor preferences
 */

let editor = null;
let currentFileType = 'module';

/**
 * Initialize Monaco Editor
 * NOTE: This requires Monaco Editor to be loaded via CDN in playground.html
 * 
 * Alternative: Use Ace Editor for lighter weight
 * @param {string} containerId - The DOM element ID for the editor
 * @returns {Promise<object>} The editor instance
 */
export async function initEditor(containerId) {
    return new Promise((resolve, reject) => {
        // Check if Monaco loader is available
        if (typeof require === 'undefined') {
            console.error('Monaco Editor loader not found!');
            reject(new Error('Monaco Editor loader not found'));
            return;
        }

        // Configure Monaco loader to use CDN
        require.config({ 
            paths: { 
                'vs': 'https://cdn.jsdelivr.net/npm/monaco-editor@0.45.0/min/vs'
            } 
        });

        // Load Monaco Editor
        require(['vs/editor/editor.main'], function() {
            // Register VB6 language
            registerVB6Language();

            // Create editor instance
            editor = monaco.editor.create(document.getElementById(containerId), {
                value: getDefaultCode(),
                language: 'vb6',
                theme: getCurrentTheme(),
                automaticLayout: true,
                fontSize: 14,
                lineNumbers: 'on',
                minimap: {
                    enabled: true
                },
                scrollBeyondLastLine: false,
                wordWrap: 'on',
                formatOnPaste: true,
                formatOnType: true,
                tabSize: 4,
                insertSpaces: true,
                // VB6 is case-insensitive
                matchBrackets: 'always'
            });

            // Listen for content changes
            editor.onDidChangeModelContent(() => {
                updateEditorStats();
                handleCodeChange();
            });

            // Listen for cursor position changes
            editor.onDidChangeCursorPosition((e) => {
                updateCursorPosition(e.position);
            });

            console.log('✅ Monaco Editor initialized');
            resolve(editor);
        });
    });
}

/**
 * Register VB6 as a custom language in Monaco
 * TODO: Complete the VB6 language definition with all keywords
 * TODO: Add support for VB6-specific syntax features
 */
function registerVB6Language() {
    // Register VB6 language
    monaco.languages.register({ id: 'vb6' });

    // Define VB6 tokens
    monaco.languages.setMonarchTokensProvider('vb6', {
        // TODO: Complete this with full VB6 syntax
        keywords: [
            'As', 'Boolean', 'ByRef', 'Byte', 'ByVal', 'Call', 'Case', 'Class',
            'Const', 'Currency', 'Date', 'Declare', 'Dim', 'Do', 'Double', 'Each',
            'Else', 'ElseIf', 'End', 'Enum', 'Event', 'Exit', 'False', 'For',
            'Friend', 'Function', 'Get', 'Global', 'GoSub', 'GoTo', 'If', 'Implements',
            'In', 'Integer', 'Is', 'Let', 'Like', 'Long', 'Loop', 'Me', 'Mod',
            'New', 'Next', 'Not', 'Nothing', 'Object', 'On', 'Option', 'Optional',
            'Or', 'ParamArray', 'Preserve', 'Private', 'Property', 'Public', 'RaiseEvent',
            'ReDim', 'Resume', 'Return', 'Select', 'Set', 'Single', 'Static', 'Step',
            'String', 'Sub', 'Then', 'To', 'True', 'Type', 'Until', 'Variant',
            'Version', 'While', 'With', 'WithEvents', 'Xor'
        ],

        operators: [
            '=', '>', '<', '<=', '>=', '<>', '+', '-', '*', '/', '\\', 
            '^', '&', 'And', 'Or', 'Not', 'Xor', 'Eqv', 'Imp', 'Mod'
        ],

        // Case insensitive
        ignoreCase: true,

        // Token rules
        tokenizer: {
            root: [
                // Comments (single quote or REM)
                [/'.*$/, 'comment'],
                [/^REM\s+.*$/, 'comment'],
                
                // Strings
                [/"([^"\\]|\\.)*$/, 'string.invalid'],  // non-terminated string
                [/"/, 'string', '@string'],
                
                // Numbers
                [/\b\d+\.?\d*[#!@%&]?\b/, 'number'],
                [/&H[0-9A-Fa-f]+/, 'number.hex'],
                [/&O[0-7]+/, 'number.octal'],
                
                // Keywords
                [/\b(?:Sub|Function|Property|End)\b/, 'keyword.control'],
                [/\b(?:If|Then|Else|ElseIf|Select|Case|For|Do|While|Loop|Next|Exit|GoTo|GoSub|On|Resume)\b/, 'keyword.control'],
                [/@?[a-zA-Z_]\w*/, {
                    cases: {
                        '@keywords': 'keyword',
                        '@default': 'identifier'
                    }
                }],
                
                // Operators
                [/[=<>!+\-*\/\\^&]/, 'operator'],
                
                // Line continuation
                [/_$/, 'operator'],
                
                // Delimiters
                [/[()[\]]/, 'delimiter.bracket'],
                [/[,.:;]/, 'delimiter'],
            ],

            string: [
                [/[^\\"]+/, 'string'],
                [/\\"/, 'string.escape'],
                [/"/, 'string', '@pop']
            ],
        },
    });

    // Configure editor theme for VB6
    monaco.editor.defineTheme('vb6-dark', {
        base: 'vs-dark',
        inherit: true,
        rules: [
            { token: 'comment', foreground: '6A9955' },
            { token: 'keyword', foreground: '569CD6' },
            { token: 'keyword.control', foreground: 'C586C0' },
            { token: 'string', foreground: 'CE9178' },
            { token: 'number', foreground: 'B5CEA8' },
            { token: 'operator', foreground: 'D4D4D4' },
        ],
        colors: {}
    });

    monaco.editor.defineTheme('vb6-light', {
        base: 'vs',
        inherit: true,
        rules: [
            { token: 'comment', foreground: '008000' },
            { token: 'keyword', foreground: '0000FF' },
            { token: 'keyword.control', foreground: 'AF00DB' },
            { token: 'string', foreground: 'A31515' },
            { token: 'number', foreground: '098658' },
            { token: 'operator', foreground: '000000' },
        ],
        colors: {}
    });
}

/**
 * Get default code to show in editor
 */
function getDefaultCode() {
    return `' VB6Parse Playground
' Enter your VB6 code here and click Parse

Option Explicit

Public Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
`;
}

/**
 * Get current theme based on site theme
 */
function getCurrentTheme() {
    const theme = document.documentElement.getAttribute('data-theme');
    return theme === 'dark' ? 'vb6-dark' : 'vb6-light';
}

/**
 * Update editor theme when site theme changes
 */
export function updateEditorTheme() {
    if (editor) {
        monaco.editor.setTheme(getCurrentTheme());
    }
}

/**
 * Get editor content
 * @returns {string} Current editor content
 */
export function getEditorContent() {
    return editor ? editor.getValue() : '';
}

/**
 * Set editor content
 * @param {string} code - Code to set in editor
 */
export function setEditorContent(code) {
    if (editor) {
        editor.setValue(code);
    }
}

/**
 * Clear editor content
 */
export function clearEditor() {
    if (editor) {
        editor.setValue('');
    }
}

/**
 * Update editor statistics (line count, character count)
 */
function updateEditorStats() {
    if (!editor) return;

    const model = editor.getModel();
    const lineCount = model.getLineCount();
    const charCount = model.getValueLength();
    
    const statsElement = document.getElementById('code-stats');
    if (statsElement) {
        statsElement.textContent = `Lines: ${lineCount} | Chars: ${charCount}`;
    }
}

/**
 * Update cursor position display
 * @param {object} position - Monaco position object {lineNumber, column}
 */
function updateCursorPosition(position) {
    const statusElement = document.getElementById('editor-line-col');
    if (statusElement) {
        statusElement.textContent = `Ln ${position.lineNumber}, Col ${position.column}`;
    }
}

/**
 * Handle code change (for auto-parse feature)
 * TODO: Implement debounced auto-parse
 */
function handleCodeChange() {
    // This will be called by main.js to trigger auto-parse
    const event = new CustomEvent('editorContentChanged');
    document.dispatchEvent(event);
}

/**
 * Highlight a range in the editor
 * @param {number} startLine - Start line (1-based)
 * @param {number} startCol - Start column (1-based)
 * @param {number} endLine - End line (1-based)
 * @param {number} endCol - End column (1-based)
 */
export function highlightRange(startLine, startCol, endLine, endCol) {
    if (!editor) return;

    // TODO: Implement highlighting
    // Use editor.deltaDecorations() to add/remove decorations
    editor.revealLineInCenter(startLine);
    editor.setSelection({
        startLineNumber: startLine,
        startColumn: startCol,
        endLineNumber: endLine,
        endColumn: endCol
    });
}

/**
 * Set the current file type
 * @param {string} fileType - 'module', 'class', 'form', or 'project'
 */
export function setFileType(fileType) {
    currentFileType = fileType;
    // Could update editor configuration based on file type
}

/**
 * Get editor instance for external access
 * @returns {object} Monaco editor instance
 */
export function getEditor() {
    return editor;
}

/**
 * Save editor content to localStorage
 * TODO: Implement localStorage persistence
 */
export function saveToLocalStorage() {
    if (editor) {
        const content = editor.getValue();
        localStorage.setItem('vb6parse-playground-code', content);
        localStorage.setItem('vb6parse-playground-filetype', currentFileType);
    }
}

/**
 * Load editor content from localStorage
 * TODO: Implement localStorage loading
 */
export function loadFromLocalStorage() {
    const content = localStorage.getItem('vb6parse-playground-code');
    const fileType = localStorage.getItem('vb6parse-playground-filetype');
    
    if (content) {
        setEditorContent(content);
    }
    
    if (fileType) {
        setFileType(fileType);
    }
}

/**
 * TODO: Future enhancements
 * - Add keyboard shortcuts (Ctrl+S to parse, etc.)
 * - Implement find/replace functionality
 * - Add code folding support
 * - Support multiple tabs/files
 * - Add snippet support for common VB6 patterns
 * - Implement IntelliSense/autocomplete
 * - Add error squiggles from parser output
 */
