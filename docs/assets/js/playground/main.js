/**
 * VB6Parse Playground - Main Application Module
 * 
 * Entry point that coordinates all other modules:
 * - Initializes WASM module
 * - Sets up editor
 * - Handles UI events
 * - Coordinates parsing and rendering
 * 
 * TODO: Wire up all event handlers
 * TODO: Implement auto-parse with debouncing
 * TODO: Add URL sharing functionality
 */

import { getExample } from './examples.js';
import * as Parser from './parser.js';
import * as Editor from './editor.js';
import * as Renderer from './renderer.js';
import * as TreeViz from './tree-viz.js';

// Application state
const state = {
    currentFileType: 'module',
    autoParse: true,
    parseTimeout: null,
    lastParseResult: null,
    isInitialized: false
};

/**
 * Main initialization function
 * Called when DOM is ready
 */
async function init() {
    console.log('🚀 Initializing VB6Parse Playground...');

    try {
        // Show loading overlay
        showLoading('Initializing WASM module...');

        // Initialize WASM module
        const wasmOk = await Parser.initWasm();
        if (!wasmOk) {
            throw new Error('Failed to initialize WASM module');
        }

        // Initialize editor
        await Editor.initEditor('editor-container');

        // Initialize tree visualization
        TreeViz.initTreeViz('tree-viz-container');

        // Set up event listeners
        setupEventListeners();

        // Load from localStorage if available
        loadFromLocalStorage();

        // Hide loading overlay
        hideLoading();

        state.isInitialized = true;
        console.log('✅ Playground initialized successfully');

    } catch (error) {
        console.error('❌ Initialization failed:', error);
        showError(`Failed to initialize playground: ${error.message}`);
        hideLoading();
    }
}

/**
 * Set up all event listeners
 */
function setupEventListeners() {
    // File type selector
    document.getElementById('file-type')?.addEventListener('change', handleFileTypeChange);

    // Examples selector
    document.getElementById('examples')?.addEventListener('change', handleExampleChange);

    // Parse button
    document.getElementById('parse-btn')?.addEventListener('click', handleParse);
    document.getElementById('parse-footer-btn')?.addEventListener('click', handleParse);

    // Share button
    document.getElementById('share-btn')?.addEventListener('click', handleShare);

    // Clear button
    document.getElementById('clear-btn')?.addEventListener('click', handleClear);

    // Auto-parse toggle
    document.getElementById('auto-parse')?.addEventListener('change', handleAutoParseToggle);

    // Tab navigation
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => handleTabChange(btn.dataset.tab));
    });

    // Token filter
    document.getElementById('show-whitespace')?.addEventListener('change', () => {
        if (state.lastParseResult) {
            Renderer.renderTokensTab(state.lastParseResult.tokens);
        }
    });

    document.getElementById('token-filter')?.addEventListener('change', handleTokenFilter);
    document.getElementById('token-search')?.addEventListener('input', handleTokenSearch);

    // CST controls
    document.getElementById('expand-all')?.addEventListener('click', handleExpandAll);
    document.getElementById('collapse-all')?.addEventListener('click', handleCollapseAll);
    document.getElementById('show-byte-ranges')?.addEventListener('change', () => {
        if (state.lastParseResult) {
            Renderer.renderCstTab(state.lastParseResult.cst);
        }
    });

    // Tree visualization controls
    document.getElementById('tree-layout-toggle')?.addEventListener('click', TreeViz.toggleLayout);
    document.getElementById('tree-fit')?.addEventListener('click', TreeViz.fitToScreen);
    document.getElementById('tree-reset-zoom')?.addEventListener('click', TreeViz.resetZoom);

    // Editor content change (for auto-parse)
    document.addEventListener('editorContentChanged', handleEditorChange);

    // Highlight in editor (from renderer)
    document.addEventListener('highlightInEditor', handleHighlightRequest);

    // Theme toggle (inherited from main site)
    document.getElementById('theme-toggle')?.addEventListener('click', handleThemeToggle);

    // Resizer for split panel
    setupResizer();

    // Window resize
    window.addEventListener('resize', handleWindowResize);

    console.log('✅ Event listeners set up');
}

/**
 * Handle file type change
 */
function handleFileTypeChange(e) {
    state.currentFileType = e.target.value;
    Editor.setFileType(state.currentFileType);
    console.log(`📄 File type changed to: ${state.currentFileType}`);

    // Auto-parse if enabled
    if (state.autoParse) {
        debouncedParse();
    }
}

/**
 * Handle example selection
 */
function handleExampleChange(e) {
    const exampleId = e.target.value;
    if (!exampleId) return;

    const example = getExample(exampleId);
    if (!example) {
        console.error(`Example ${exampleId} not found`);
        return;
    }

    // Set file type
    document.getElementById('file-type').value = example.fileType;
    state.currentFileType = example.fileType;

    // Load code into editor
    Editor.setEditorContent(example.code);

    // Auto-parse if enabled
    if (state.autoParse) {
        handleParse();
    }

    // Reset selector
    e.target.value = '';

    console.log(`📝 Loaded example: ${example.name}`);
}

/**
 * Handle parse button click
 */
async function handleParse() {
    if (!state.isInitialized) {
        showError('Playground not initialized yet');
        return;
    }

    const code = Editor.getEditorContent();
    if (!code || code.trim().length === 0) {
        return;
    }

    try {
        console.log(`🔍 Parsing ${state.currentFileType}...`);

        // Parse code
        const result = await Parser.parseCode(code, state.currentFileType);
        state.lastParseResult = result;

        // Render results
        Renderer.renderOutput(result);
        TreeViz.renderTree(result.cst);

        console.log(`✅ Parse complete in ${result.parseTimeMs.toFixed(2)}ms`);

    } catch (error) {
        console.error('❌ Parse failed:', error);
        showError(`Parse failed: ${error.message}`);
    }
}

/**
 * Handle editor content change (for auto-parse)
 */
function handleEditorChange() {
    if (state.autoParse) {
        debouncedParse();
    }

    // Save to localStorage
    saveToLocalStorage();
}

/**
 * Debounced parse (500ms delay)
 */
function debouncedParse() {
    if (state.parseTimeout) {
        clearTimeout(state.parseTimeout);
    }

    state.parseTimeout = setTimeout(() => {
        handleParse();
    }, 500);
}

/**
 * Handle auto-parse toggle
 */
function handleAutoParseToggle(e) {
    state.autoParse = e.target.checked;
    console.log(`🔄 Auto-parse ${state.autoParse ? 'enabled' : 'disabled'}`);
}

/**
 * Handle share button click
 * TODO: Implement URL encoding and sharing
 */
function handleShare() {
    console.log('🔧 TODO: Implement share functionality');
    
    // TODO: Encode code and file type in URL
    // const code = Editor.getEditorContent();
    // const encoded = btoa(encodeURIComponent(code));
    // const url = `${window.location.origin}${window.location.pathname}?code=${encoded}&type=${state.currentFileType}`;
    
    // TODO: Copy to clipboard or show share dialog
    showError('Share functionality coming soon!');
}

/**
 * Handle clear button click
 */
function handleClear() {
    if (confirm('Clear editor and output?')) {
        Editor.clearEditor();
        Renderer.clearOutput();
        TreeViz.clearTree();
        state.lastParseResult = null;
        console.log('🗑️ Cleared editor and output');
    }
}

/**
 * Handle tab change
 */
function handleTabChange(tabId) {
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });

    // Update tab panes
    document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.classList.toggle('active', pane.id === `${tabId}-tab`);
    });

    console.log(`📑 Switched to ${tabId} tab`);

    // Initialize tree viz if switching to tree tab for the first time
    if (tabId === 'tree' && state.lastParseResult) {
        TreeViz.renderTree(state.lastParseResult.cst);
    }
}

/**
 * Handle token filter
 * TODO: Implement token filtering
 */
function handleTokenFilter(e) {
    console.log('🔧 TODO: Implement token filter:', e.target.value);
}

/**
 * Handle token search
 * TODO: Implement token search
 */
function handleTokenSearch(e) {
    console.log('🔧 TODO: Implement token search:', e.target.value);
}

/**
 * Handle expand all (CST)
 */
function handleExpandAll() {
    document.querySelectorAll('.cst-node.collapsed').forEach(node => {
        node.classList.remove('collapsed');
    });
}

/**
 * Handle collapse all (CST)
 */
function handleCollapseAll() {
    document.querySelectorAll('.cst-node').forEach(node => {
        if (node.querySelector('.cst-node-children')) {
            node.classList.add('collapsed');
        }
    });
}

/**
 * Handle highlight request from renderer
 */
function handleHighlightRequest(e) {
    const { line, column, length } = e.detail;
    Editor.highlightRange(line, column, line, column + length);
}

/**
 * Handle theme toggle
 */
function handleThemeToggle() {
    // Theme switcher is handled by theme-switcher.js from main site
    // Just update editor theme
    Editor.updateEditorTheme();
}

/**
 * Set up split panel resizer
 */
function setupResizer() {
    const resizer = document.getElementById('resizer');
    const leftPanel = document.querySelector('.editor-panel');
    const rightPanel = document.querySelector('.output-panel');

    if (!resizer || !leftPanel || !rightPanel) return;

    let isResizing = false;
    let startX = 0;
    let startLeftWidth = 0;

    resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        startX = e.clientX;
        startLeftWidth = leftPanel.offsetWidth;
        document.body.style.cursor = 'col-resize';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;

        const deltaX = e.clientX - startX;
        const newLeftWidth = startLeftWidth + deltaX;
        const minWidth = 300;
        const maxWidth = window.innerWidth - 300 - 8; // 8px for resizer

        if (newLeftWidth >= minWidth && newLeftWidth <= maxWidth) {
            leftPanel.style.width = `${newLeftWidth}px`;
            leftPanel.style.flex = 'none';
        }
    });

    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            document.body.style.cursor = '';
        }
    });
}

/**
 * Handle window resize
 */
function handleWindowResize() {
    // Tree viz will handle its own resize due to automaticLayout
    // Just log for now
    console.log('↔️ Window resized');
}

/**
 * Show loading overlay
 */
function showLoading(message = 'Loading...') {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.querySelector('p').textContent = message;
        overlay.classList.remove('hidden');
    }
}

/**
 * Hide loading overlay
 */
function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.classList.add('hidden');
    }
}

/**
 * Show error modal
 */
function showError(message) {
    const modal = document.getElementById('error-modal');
    const messageEl = document.getElementById('error-message');
    
    if (modal && messageEl) {
        messageEl.textContent = message;
        modal.classList.remove('hidden');
    }

    console.error('❌ Error:', message);
}

/**
 * Hide error modal
 */
function hideError() {
    const modal = document.getElementById('error-modal');
    if (modal) {
        modal.classList.add('hidden');
    }
}

// Error modal close button
document.querySelector('.modal-close')?.addEventListener('click', hideError);
document.getElementById('error-modal')?.addEventListener('click', (e) => {
    if (e.target.id === 'error-modal') {
        hideError();
    }
});

/**
 * Save state to localStorage
 */
function saveToLocalStorage() {
    try {
        const code = Editor.getEditorContent();
        localStorage.setItem('vb6parse-playground-code', code);
        localStorage.setItem('vb6parse-playground-filetype', state.currentFileType);
        localStorage.setItem('vb6parse-playground-autoparse', state.autoParse);
    } catch (error) {
        console.warn('Failed to save to localStorage:', error);
    }
}

/**
 * Load state from localStorage
 */
function loadFromLocalStorage() {
    try {
        const code = localStorage.getItem('vb6parse-playground-code');
        const fileType = localStorage.getItem('vb6parse-playground-filetype');
        const autoParse = localStorage.getItem('vb6parse-playground-autoparse');

        if (code) {
            Editor.setEditorContent(code);
        }

        if (fileType) {
            state.currentFileType = fileType;
            document.getElementById('file-type').value = fileType;
        }

        if (autoParse !== null) {
            state.autoParse = autoParse === 'true';
            document.getElementById('auto-parse').checked = state.autoParse;
        }

        console.log('📂 Loaded state from localStorage');
    } catch (error) {
        console.warn('Failed to load from localStorage:', error);
    }
}

/**
 * Load code from URL parameter
 * TODO: Implement URL parameter loading
 */
function loadFromUrl() {
    const params = new URLSearchParams(window.location.search);
    const encodedCode = params.get('code');
    const fileType = params.get('type');

    if (encodedCode) {
        try {
            const code = decodeURIComponent(atob(encodedCode));
            Editor.setEditorContent(code);
            console.log('🔗 Loaded code from URL');
        } catch (error) {
            console.error('Failed to decode URL code:', error);
        }
    }

    if (fileType) {
        state.currentFileType = fileType;
        document.getElementById('file-type').value = fileType;
    }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}

/**
 * TODO: Future Enhancements
 * 
 * 1. URL Sharing:
 *    - Encode code in URL with LZ-string compression
 *    - Generate shareable links
 *    - QR code generation
 * 
 * 2. Keyboard Shortcuts:
 *    - Ctrl+Enter: Parse
 *    - Ctrl+K: Clear
 *    - Ctrl+S: Save (download)
 *    - Ctrl+O: Load file
 * 
 * 3. File Operations:
 *    - Load .vb6 files from disk
 *    - Save editor content to file
 *    - Drag & drop file support
 * 
 * 4. Session Management:
 *    - Multiple tabs/files
 *    - History of parsed code
 *    - Favorites/bookmarks
 * 
 * 5. Collaboration:
 *    - Share sessions with others
 *    - Real-time collaboration
 *    - Comments and annotations
 * 
 * 6. Analytics:
 *    - Track usage statistics
 *    - Popular examples
 *    - Performance metrics
 */

export default {
    init,
    state
};
