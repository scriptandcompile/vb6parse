/**
 * VB6Parse Playground - Renderer Module
 * 
 * Handles rendering parse results to the output tabs (Tokens, CST, Info).
 * Each tab has its own rendering logic.
 * 
 * TODO: Implement filtering and search for tokens
 * TODO: Add syntax highlighting for CST view
 * TODO: Implement click-to-highlight interaction
 */

import { getEditor } from './editor.js';

/**
 * Render all output tabs from parse result
 * @param {ParseResult} result - Parse result from parser.js
 */
export function renderOutput(result) {
    renderTokensTab(result.tokens);
    renderCstTab(result.cst);
    renderInfoTab(result);
    
    // Update parse time in header (parseTimeMs is set by JavaScript in parser.js)
    const parseTime = result.parseTimeMs ?? 0;
    updateParseTime(parseTime);
}

/**
 * Update parse time display
 * @param {number} timeMs - Parse time in milliseconds
 */
function updateParseTime(timeMs) {
    const element = document.getElementById('parse-time');
    if (element) {
        element.textContent = `Parse time: ${timeMs.toFixed(2)}ms`;
        element.style.color = timeMs < 100 ? 'var(--status-success)' : 
                             timeMs < 500 ? 'var(--status-warning)' : 
                             'var(--status-error)';
    }
}

/**
 * Render tokens to the Tokens tab
 * @param {TokenInfo[]} tokens - Array of tokens
 */
export function renderTokensTab(tokens) {
    const container = document.getElementById('tokens-content');
    if (!container) return;

    // Remove placeholder
    container.innerHTML = '';

    // Check if we should show whitespace
    const showWhitespace = document.getElementById('show-whitespace')?.checked ?? false;

    // Filter tokens
    const filteredTokens = showWhitespace ? tokens : tokens.filter(t => t.kind !== 'whitespace' & t.kind !== 'Newline');

    // Create table
    const table = document.createElement('table');
    table.className = 'tokens-table';

    // Table header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Type</th>
            <th>Value</th>
            <th>Position</th>
            <th>Length</th>
        </tr>
    `;
    table.appendChild(thead);

    // Table body
    const tbody = document.createElement('tbody');
    filteredTokens.forEach((token, index) => {
        const row = document.createElement('tr');
        row.dataset.tokenIndex = index;
        row.dataset.line = token.line;
        row.dataset.column = token.column;
        
        // Add click handler to highlight in editor
        row.addEventListener('click', () => {
            highlightTokenInEditor(token);
        });

        // Type badge
        const typeCell = document.createElement('td');
        const typeBadge = document.createElement('span');
        typeBadge.className = `token-type ${token.kind}`;
        typeBadge.textContent = token.kind;
        typeCell.appendChild(typeBadge);

        // Value (escape HTML)
        const valueCell = document.createElement('td');
        valueCell.textContent = token.content;
        valueCell.style = "white-space: pre;";
        valueCell.style.fontFamily = "'Courier New', monospace";

        // Position
        const posCell = document.createElement('td');
        posCell.textContent = `${token.line}:${token.column}`;
        posCell.style.fontFamily = "'Courier New', monospace";

        // Length
        const lengthCell = document.createElement('td');
        lengthCell.textContent = token.length;

        row.appendChild(typeCell);
        row.appendChild(valueCell);
        row.appendChild(posCell);
        row.appendChild(lengthCell);

        tbody.appendChild(row);
    });

    table.appendChild(tbody);
    container.appendChild(table);

    console.log(`✅ Rendered ${filteredTokens.length} tokens`);
}

// Counter for generating unique node IDs
let nodeIdCounter = 0;

/**
 * Render CST to the CST tab
 * @param {CstNode} cst - Root CST node
 */
export function renderCstTab(cst) {
    const container = document.getElementById('cst-content');
    if (!container) return;

    // Remove placeholder
    container.innerHTML = '';

    // Reset node ID counter
    nodeIdCounter = 0;

    // Check if we should show byte ranges
    const showRanges = document.getElementById('show-byte-ranges')?.checked ?? true;

    // Add node IDs to CST tree
    addNodeIds(cst);

    // Render tree recursively
    const treeElement = renderCstNode(cst, 0, showRanges);
    container.appendChild(treeElement);

    console.log(`✅ Rendered CST tree`);
}

/**
 * Add unique IDs to CST nodes recursively
 * @param {CstNode} node - CST node
 */
function addNodeIds(node) {
    if (!node) return;
    node._nodeId = nodeIdCounter++;
    if (node.children) {
        node.children.forEach(child => addNodeIds(child));
    }
}

/**
 * Render a single CST node recursively
 * @param {CstNode} node - CST node
 * @param {number} depth - Current depth
 * @param {boolean} showRanges - Show byte ranges
 * @returns {HTMLElement} Rendered node element
 */
function renderCstNode(node, depth, showRanges) {
    const nodeDiv = document.createElement('div');
    nodeDiv.className = 'cst-node';
    nodeDiv.dataset.nodeType = node.kind;
    nodeDiv.dataset.depth = depth;

    // Node header
    const header = document.createElement('div');
    header.className = 'cst-node-header';

    // Toggle (if has children)
    if (node.children && node.children.length > 0) {
        const toggle = document.createElement('span');
        toggle.className = 'cst-node-toggle';
        toggle.addEventListener('click', (e) => {
            e.stopPropagation();
            nodeDiv.classList.toggle('collapsed');
        });
        header.appendChild(toggle);
    } else {
        const spacer = document.createElement('span');
        spacer.style.marginRight = '16px';
        header.appendChild(spacer);
    }

    // Node name
    const name = document.createElement('span');
    name.className = 'cst-node-name';
    name.textContent = node.kind;
    header.appendChild(name);

    // Node value (for leaf nodes)
    if (node.value !== undefined) {
        const value = document.createElement('span');
        value.textContent = ` = "${escapeHtml(node.value)}"`;
        value.style.color = 'var(--placeholder-color)';
        header.appendChild(value);
    }

    // Byte range
    if (showRanges && node.range) {
        const range = document.createElement('span');
        range.className = 'cst-node-range';
        range.textContent = ` [${node.range[0]}..${node.range[1]}]`;
        header.appendChild(range);
    }

    // Store range data and node ID for event handlers
    if (node.range) {
        nodeDiv.dataset.rangeStart = node.range[0];
        nodeDiv.dataset.rangeEnd = node.range[1];
    }
    if (node._nodeId !== undefined) {
        nodeDiv.dataset.nodeId = node._nodeId;
    }

    // Add click handler to highlight and position cursor
    header.addEventListener('click', (e) => {
        e.stopPropagation();
        highlightAndPositionCursor(node);
    });

    // Add hover handler to highlight
    header.addEventListener('mouseenter', () => {
        highlightNodeInEditor(node, true);
    });

    header.addEventListener('mouseleave', () => {
        clearEditorHighlight();
    });

    nodeDiv.appendChild(header);

    // Children
    if (node.children && node.children.length > 0) {
        const childrenDiv = document.createElement('div');
        childrenDiv.className = 'cst-node-children';

        node.children.forEach(child => {
            const childElement = renderCstNode(child, depth + 1, showRanges);
            childrenDiv.appendChild(childElement);
        });

        nodeDiv.appendChild(childrenDiv);
    }

    return nodeDiv;
}

/**
 * Render Info/Stats tab
 * @param {ParseResult} result - Parse result
 */
export function renderInfoTab(result) {
    const container = document.getElementById('info-content');
    if (!container) return;

    // Hide placeholder
    const placeholder = container.querySelector('.placeholder');
    if (placeholder) {
        placeholder.classList.add('hidden');
    }

    // Show and populate statistics section
    const statsSection = document.getElementById('info-stats');
    if (statsSection) {
        statsSection.classList.remove('hidden');
        
        // Update stat values (handle both camelCase and snake_case for compatibility)
        const tokenCount = result.stats?.token_count ?? result.stats?.tokenCount ?? 0;
        // parseTimeMs is set by JavaScript in parser.js, parse_time_ms is from WASM (always 0)
        const parseTime = result.parseTimeMs ?? 0;
        const treeDepth = result.stats?.tree_depth ?? result.stats?.treeDepth ?? 0;
        const nodeCount = result.stats?.node_count ?? result.stats?.nodeCount ?? 0;
        const fileType = result.cst?.kind ?? 'Unknown';
        
        // Get lines and chars from editor
        const editor = getEditor();
        let lineCount = 0;
        let charCount = 0;
        if (editor) {
            const model = editor.getModel();
            lineCount = model.getLineCount();
            charCount = model.getValueLength();
        }
        
        const statTokensEl = document.getElementById('stat-tokens');
        const statParseTimeEl = document.getElementById('stat-parse-time');
        const statTreeDepthEl = document.getElementById('stat-tree-depth');
        const statNodeCountEl = document.getElementById('stat-node-count');
        const statFileTypeEl = document.getElementById('stat-file-type');
        const statLinesEl = document.getElementById('stat-lines');
        const statCharsEl = document.getElementById('stat-chars');
        
        if (statTokensEl) statTokensEl.textContent = tokenCount.toLocaleString();
        if (statParseTimeEl) statParseTimeEl.textContent = `${parseTime.toFixed(2)}ms`;
        if (statTreeDepthEl) statTreeDepthEl.textContent = treeDepth.toLocaleString();
        if (statNodeCountEl) statNodeCountEl.textContent = nodeCount.toLocaleString();
        if (statFileTypeEl) statFileTypeEl.textContent = fileType;
        if (statLinesEl) statLinesEl.textContent = lineCount.toLocaleString();
        if (statCharsEl) statCharsEl.textContent = charCount.toLocaleString();
    }

    // Render errors
    const errorsSection = document.getElementById('info-errors');
    const errorsList = document.getElementById('errors-list');
    if (result.errors && result.errors.length > 0) {
        errorsSection?.classList.remove('hidden');
        if (errorsList) {
            errorsList.innerHTML = '';
            result.errors.forEach(error => {
                const errorElement = createErrorElement(error);
                errorsList.appendChild(errorElement);
            });
        }
    } else {
        errorsSection?.classList.add('hidden');
    }

    // Render warnings
    const warningsSection = document.getElementById('info-warnings');
    const warningsList = document.getElementById('warnings-list');
    if (result.warnings && result.warnings.length > 0) {
        warningsSection?.classList.remove('hidden');
        if (warningsList) {
            warningsList.innerHTML = '';
            result.warnings.forEach(warning => {
                const warningElement = createWarningElement(warning);
                warningsList.appendChild(warningElement);
            });
        }
    } else {
        warningsSection?.classList.add('hidden');
    }

    console.log(`✅ Rendered info tab`);
}

/**
 * Create error element
 * @param {ErrorInfo} error - Error info
 * @returns {HTMLElement} Error element
 */
function createErrorElement(error) {
    const div = document.createElement('div');
    div.className = 'error-item';
    
    const header = document.createElement('div');
    header.className = 'error-header';
    
    const type = document.createElement('span');
    type.className = 'error-type';
    type.textContent = error.type || 'Parse Error';
    
    const location = document.createElement('span');
    location.className = 'error-location';
    location.textContent = `Line ${error.line}, Col ${error.column}`;
    
    header.appendChild(type);
    header.appendChild(location);
    
    const message = document.createElement('div');
    message.className = 'error-message';
    message.textContent = error.message;
    
    div.appendChild(header);
    div.appendChild(message);
    
    // Click to highlight in editor
    div.addEventListener('click', () => {
        if (error.range) {
            // TODO: Calculate line/col from byte offset
            highlightRange(error.line, error.column, error.line, error.column + 1);
        }
    });
    
    return div;
}

/**
 * Create warning element
 * @param {ErrorInfo} warning - Warning info
 * @returns {HTMLElement} Warning element
 */
function createWarningElement(warning) {
    const div = document.createElement('div');
    div.className = 'warning-item';
    
    const header = document.createElement('div');
    header.className = 'warning-header';
    
    const type = document.createElement('span');
    type.className = 'warning-type';
    type.textContent = warning.type || 'Warning';
    
    const location = document.createElement('span');
    location.className = 'warning-location';
    location.textContent = `Line ${warning.line}, Col ${warning.column}`;
    
    header.appendChild(type);
    header.appendChild(location);
    
    const message = document.createElement('div');
    message.className = 'warning-message';
    message.textContent = warning.message;
    
    div.appendChild(header);
    div.appendChild(message);
    
    return div;
}

/**
 * Highlight token in editor
 * TODO: Integrate with editor.js
 * @param {TokenInfo} token - Token to highlight
 */
function highlightTokenInEditor(token) {
    console.log('TODO: Highlight token in editor:', token);
    // Import and call editor.highlightRange()
    const event = new CustomEvent('highlightInEditor', {
        detail: {
            line: token.line,
            column: token.column,
            length: token.length
        }
    });
    document.dispatchEvent(event);
}

/**
 * Highlight node in editor (hover)
 * @param {CstNode} node - Node to highlight
 * @param {boolean} isHover - Whether this is a hover event
 */
function highlightNodeInEditor(node, isHover = false) {
    if (!node.range) return;

    const event = new CustomEvent('highlightNodeInEditor', {
        detail: {
            startOffset: node.range[0],
            endOffset: node.range[1],
            isHover: isHover
        }
    });
    document.dispatchEvent(event);
}

/**
 * Highlight node and position cursor at start (click)
 * @param {CstNode} node - Node to highlight
 */
function highlightAndPositionCursor(node) {
    if (!node.range) return;

    const event = new CustomEvent('highlightAndPositionCursor', {
        detail: {
            startOffset: node.range[0],
            endOffset: node.range[1]
        }
    });
    document.dispatchEvent(event);
}

/**
 * Clear editor highlight
 */
function clearEditorHighlight() {
    const event = new CustomEvent('clearEditorHighlight');
    document.dispatchEvent(event);
}

/**
 * Escape HTML special characters
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Clear all output tabs
 */
export function clearOutput() {
    // Reset tokens tab
    const tokensContent = document.getElementById('tokens-content');
    if (tokensContent) {
        tokensContent.innerHTML = '<div class="placeholder"><p>👈 Enter VB6 code and click Parse to see tokens</p></div>';
    }

    // Reset CST tab
    const cstContent = document.getElementById('cst-content');
    if (cstContent) {
        cstContent.innerHTML = '<div class="placeholder"><p>👈 Parse code to see the Concrete Syntax Tree</p></div>';
    }

    // Reset info tab
    const infoStats = document.getElementById('info-stats');
    if (infoStats) {
        infoStats.classList.add('hidden');
    }
    
    const infoErrors = document.getElementById('info-errors');
    if (infoErrors) {
        infoErrors.classList.add('hidden');
    }
    
    const infoWarnings = document.getElementById('info-warnings');
    if (infoWarnings) {
        infoWarnings.classList.add('hidden');
    }

    // Show placeholder
    const infoContent = document.getElementById('info-content');
    const placeholder = infoContent?.querySelector('.placeholder');
    if (placeholder) {
        placeholder.classList.remove('hidden');
    }

    // Reset parse time
    updateParseTime(0);
}

/**
 * TODO: Future enhancements
 * - Add search/filter functionality for tokens
 * - Implement token type filter dropdown
 * - Add copy button for token/node text
 * - Syntax highlighting for CST node values
 * - Export tokens/CST to JSON
 * - Add statistics charts (token type distribution, etc.)
 * - Implement virtual scrolling for large token lists
 */

export default {
    renderOutput,
    renderTokensTab,
    renderCstTab,
    renderInfoTab,
    clearOutput
};
