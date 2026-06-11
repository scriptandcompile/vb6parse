/**
 * VB6Parse Playground - Tree Visualization Module
 *
 * Renders CST data with Treant.js and wires tree interactions back to the editor.
 */

import { getEditorContent } from './editor.js';

let currentContainerId = null;
let currentCst = null;
let currentLayout = 'vertical'; // 'vertical' or 'horizontal'
let renderVersion = 0;
let nodeCounter = 0;
let removeWhitespaceNodes = true;
let selectedNodeDomId = null;

const panState = {
    active: false,
    startX: 0,
    startY: 0,
    startScrollLeft: 0,
    startScrollTop: 0
};

const nodeMetaByDomId = new Map();

/**
 * Initialize the tree visualization
 * @param {string} containerId - Container element ID
 */
export function initTreeViz(containerId) {
    currentContainerId = containerId;

    const container = getContainer();
    if (!container) {
        console.error(`Container ${containerId} not found`);
        return;
    }

    const removeWhitespaceToggle = document.getElementById('tree-remove-whitespace');
    if (removeWhitespaceToggle) {
        removeWhitespaceNodes = removeWhitespaceToggle.checked;
    }

    updateLayoutIcon();
    setupPanDrag(container);
    renderPlaceholder('Parse code to see the visual tree');
    console.log('✅ Tree visualization initialized (Treant.js)');
}

/**
 * Render CST as a Treant.js tree
 * @param {CstNode} cst - Root CST node
 */
export function renderTree(cst) {
    const container = getContainer();
    if (!container) {
        console.error('Tree visualization not initialized');
        return;
    }

    currentCst = cst;

    if (!cst) {
        renderPlaceholder('No CST available');
        return;
    }

    if (typeof window.Treant !== 'function') {
        console.error('Treant.js is not loaded');
        renderPlaceholder('Treant.js failed to load');
        return;
    }

    nodeCounter = 0;
    nodeMetaByDomId.clear();

    const nodeStructure = convertCstToTreantNode(cst, 0, true);
    const config = {
        chart: {
            container: `#${currentContainerId}`,
            rootOrientation: currentLayout === 'vertical' ? 'NORTH' : 'WEST',
            nodeAlign: 'TOP',
            hideRootNode: false,
            animateOnInit: false,
            animateOnInitDelay: 0,
            connectors: {
                type: 'step',
                style: {
                    stroke: '#94a3b8',
                    'stroke-width': 2,
                    'arrow-end': 'none'
                }
            },
            node: {
                HTMLclass: 'vb6-tree-node'
            }
        },
        nodeStructure
    };

    container.innerHTML = '';
    renderVersion += 1;
    const thisRender = renderVersion;
    selectedNodeDomId = null;

    try {
        new window.Treant(config);
        bindNodeInteractions(thisRender);
        console.log('✅ Rendered CST tree with Treant.js');
    } catch (error) {
        console.error('Failed to render Treant tree:', error);
        renderPlaceholder('Failed to render tree');
    }
}

/**
 * Toggle tree layout between vertical and horizontal
 */
export function toggleLayout() {
    currentLayout = currentLayout === 'vertical' ? 'horizontal' : 'vertical';
    updateLayoutIcon();

    if (currentCst) {
        renderTree(currentCst);
    }

    console.log(`🔄 Switched to ${currentLayout} layout`);
}

/**
 * Fit tree to screen (best-effort centering for scrollable Treant container)
 */
export function fitToScreen() {
    centerTreeView();
    console.log('📐 Centered tree view');
}

/**
 * Reset tree position
 */
export function resetZoom() {
    const container = getContainer();
    if (!container) return;

    container.scrollTo({ left: 0, top: 0, behavior: 'smooth' });
    console.log('🔄 Reset tree view');
}

/**
 * Clear tree visualization
 */
export function clearTree() {
    currentCst = null;
    nodeMetaByDomId.clear();
    selectedNodeDomId = null;
    renderPlaceholder('Parse code to see the visual tree');
}

/**
 * Toggle removal of whitespace/newline CST nodes from the tree
 * @param {boolean} enabled
 */
export function setRemoveWhitespace(enabled) {
    removeWhitespaceNodes = Boolean(enabled);
}

/**
 * Focus the most specific rendered tree node that contains a byte offset
 * @param {number} byteOffset
 * @returns {boolean} True if a node was found and focused
 */
export function focusNodeByOffset(byteOffset) {
    const target = findBestNodeByOffset(byteOffset);
    if (!target) {
        return false;
    }

    focusTreeNode(target.domId, target.meta);
    return true;
}

/**
 * Get tree statistics
 * @param {CstNode} cst - Root CST node
 * @returns {object} Tree statistics
 */
export function getTreeStats(cst) {
    let nodeCount = 0;
    let maxDepth = 0;

    function traverse(node, depth) {
        nodeCount += 1;
        maxDepth = Math.max(maxDepth, depth);

        if (node.children && node.children.length > 0) {
            node.children.forEach(child => traverse(child, depth + 1));
        }
    }

    traverse(cst, 0);

    return {
        nodeCount,
        maxDepth
    };
}

/**
 * Export tree as SVG
 * TODO: Implement SVG export
 */
export function exportAsSvg() {
    console.log('🔧 TODO: Export tree as SVG');
}

/**
 * Export tree as PNG
 * TODO: Implement PNG export
 */
export function exportAsPng() {
    console.log('🔧 TODO: Export tree as PNG');
}

function getContainer() {
    if (!currentContainerId) return null;
    return document.getElementById(currentContainerId);
}

function renderPlaceholder(message) {
    const container = getContainer();
    if (!container) return;

    container.innerHTML = `
        <div class="placeholder">
            <p>🌳 ${escapeHtml(message)}</p>
        </div>
    `;
}

function updateLayoutIcon() {
    const icon = document.getElementById('layout-icon');
    if (!icon) return;

    icon.textContent = currentLayout === 'vertical' ? '↕️' : '↔️';
}

function convertCstToTreantNode(node, depth, isRoot = false) {
    if (!isRoot && shouldFilterOutNode(node)) {
        return null;
    }

    const domId = `tree-node-${nodeCounter++}`;
    const children = Array.isArray(node.children)
        ? buildTreantChildren(node.children, depth + 1)
        : [];
    const childCount = children.length;

    nodeMetaByDomId.set(domId, {
        kind: node.kind,
        depth,
        childCount,
        value: node.value,
        range: Array.isArray(node.range) ? node.range : null
    });

    const displayName = truncate(node.kind || 'Unknown', 40);
    const rangeText = Array.isArray(node.range)
        ? `[${node.range[0]}..${node.range[1]}]`
        : 'No range';

    const treantNode = {
        HTMLid: domId,
        text: {
            name: displayName,
            desc: rangeText
        },
        HTMLclass: getNodeClass(node.kind)
    };

    if (childCount > 0) {
        treantNode.children = children;
    }

    return treantNode;
}

function buildTreantChildren(children, depth) {
    const result = [];

    children.forEach((child) => {
        const converted = convertCstToTreantNode(child, depth, false);
        if (converted) {
            result.push(converted);
            return;
        }

        if (Array.isArray(child.children) && child.children.length > 0) {
            result.push(...buildTreantChildren(child.children, depth + 1));
        }
    });

    return result;
}

function shouldFilterOutNode(node) {
    if (!removeWhitespaceNodes || !node || !node.kind) {
        return false;
    }

    const kind = String(node.kind).toLowerCase();
    return kind.includes('whitespace') || kind.includes('newline');
}

function getNodeClass(kind) {
    const value = (kind || '').toLowerCase();

    if (value.includes('statement') || value.includes('declaration')) {
        return 'vb6-tree-node node-statement';
    }
    if (value.includes('expression')) {
        return 'vb6-tree-node node-expression';
    }
    if (value.includes('literal') || value.includes('number') || value.includes('string')) {
        return 'vb6-tree-node node-literal';
    }
    if (value.includes('keyword')) {
        return 'vb6-tree-node node-keyword';
    }

    return 'vb6-tree-node node-default';
}

function bindNodeInteractions(renderToken) {
    const container = getContainer();
    if (!container) return;

    requestAnimationFrame(() => {
        if (renderToken !== renderVersion) {
            return;
        }

        const nodeElements = container.querySelectorAll('.node');
        nodeElements.forEach((nodeEl) => {
            const meta = nodeMetaByDomId.get(nodeEl.id);
            if (!meta) {
                return;
            }

            nodeEl.addEventListener('mouseenter', () => {
                if (!meta.range) return;
                nodeEl.classList.add('hovered');
                dispatchHighlightNode(meta.range[0], meta.range[1], true);
            });

            nodeEl.addEventListener('mouseleave', () => {
                nodeEl.classList.remove('hovered');
                dispatchClearHighlight();
            });

            nodeEl.addEventListener('click', () => {
                if (!meta.range) return;
                focusTreeNode(nodeEl.id, meta);
                dispatchHighlightAndPosition(meta.range[0], meta.range[1]);
            });
        });

        centerTreeView();
    });
}

function focusTreeNode(domId, meta) {
    const container = getContainer();
    if (!container) {
        return;
    }

    if (selectedNodeDomId) {
        const prevNode = document.getElementById(selectedNodeDomId);
        prevNode?.classList.remove('selected');
    }

    const nodeEl = document.getElementById(domId);
    if (!nodeEl) {
        return;
    }

    nodeEl.classList.add('selected');
    selectedNodeDomId = domId;
    showNodeDetails(meta);
    nodeEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
}

function findBestNodeByOffset(byteOffset) {
    let best = null;
    let bestRangeSize = Number.POSITIVE_INFINITY;

    nodeMetaByDomId.forEach((meta, domId) => {
        if (!meta.range) {
            return;
        }

        const [start, end] = meta.range;
        if (byteOffset < start || byteOffset >= end) {
            return;
        }

        const size = end - start;
        if (!best || size < bestRangeSize) {
            best = { domId, meta };
            bestRangeSize = size;
        }
    });

    return best;
}

function setupPanDrag(container) {
    container.addEventListener('mousedown', (event) => {
        if (event.button !== 0) {
            return;
        }

        if (event.target.closest('.node')) {
            return;
        }

        panState.active = true;
        panState.startX = event.clientX;
        panState.startY = event.clientY;
        panState.startScrollLeft = container.scrollLeft;
        panState.startScrollTop = container.scrollTop;
        container.classList.add('panning');
        event.preventDefault();
    });

    container.addEventListener('mousemove', (event) => {
        if (!panState.active) {
            return;
        }

        const dx = event.clientX - panState.startX;
        const dy = event.clientY - panState.startY;

        container.scrollLeft = panState.startScrollLeft - dx;
        container.scrollTop = panState.startScrollTop - dy;
    });

    const stopPanning = () => {
        panState.active = false;
        container.classList.remove('panning');
    };

    container.addEventListener('mouseup', stopPanning);
    container.addEventListener('mouseleave', stopPanning);
}

function centerTreeView() {
    const container = getContainer();
    if (!container) return;

    const firstNode = container.querySelector('.node');
    if (!firstNode) return;

    const containerRect = container.getBoundingClientRect();
    const nodeRect = firstNode.getBoundingClientRect();

    const left = container.scrollLeft + (nodeRect.left - containerRect.left) - (container.clientWidth / 2) + (nodeRect.width / 2);
    const top = container.scrollTop + (nodeRect.top - containerRect.top) - 24;

    container.scrollTo({
        left: Math.max(0, left),
        top: Math.max(0, top),
        behavior: 'smooth'
    });
}

function showNodeDetails(meta) {
    const detailsPanel = document.getElementById('tree-node-details');
    const detailsContent = document.getElementById('node-details-content');
    if (!detailsPanel || !detailsContent) {
        return;
    }

    const range = meta.range
        ? `[${meta.range[0]}..${meta.range[1]}]`
        : 'N/A';
    const nodeText = getNodeText(meta);

    detailsContent.innerHTML = `
        <p><strong>Kind:</strong> ${escapeHtml(meta.kind || 'Unknown')}</p>
        <p><strong>Range:</strong> ${range}</p>
        <p><strong>Text:</strong><span class="tree-node-text">${escapeHtml(nodeText)}</span></p>
        <p><strong>Depth:</strong> ${meta.depth}</p>
        <p><strong>Children:</strong> ${meta.childCount}</p>
    `;

    detailsPanel.classList.remove('hidden');
}

function getNodeText(meta) {
    if (meta.value !== undefined && meta.value !== null) {
        const valueText = String(meta.value);
        if (valueText.length > 0) {
            return valueText;
        }
    }

    if (meta.range) {
        const content = getEditorContent();
        if (content) {
            const [start, end] = meta.range;
            if (Number.isInteger(start) && Number.isInteger(end) && end >= start) {
                const slice = content.slice(start, end);
                if (slice.length > 0) {
                    return slice;
                }
            }
        }
    }

    return '(no text)';
}

function dispatchHighlightNode(startOffset, endOffset, isHover) {
    const event = new CustomEvent('highlightNodeInEditor', {
        detail: { startOffset, endOffset, isHover }
    });
    document.dispatchEvent(event);
}

function dispatchHighlightAndPosition(startOffset, endOffset) {
    const event = new CustomEvent('highlightAndPositionCursor', {
        detail: { startOffset, endOffset }
    });
    document.dispatchEvent(event);
}

function dispatchClearHighlight() {
    document.dispatchEvent(new CustomEvent('clearEditorHighlight'));
}

function truncate(text, max) {
    const value = String(text ?? '');
    if (value.length <= max) return value;
    return `${value.slice(0, max - 1)}…`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

export default {
    initTreeViz,
    renderTree,
    toggleLayout,
    fitToScreen,
    resetZoom,
    clearTree,
    setRemoveWhitespace,
    focusNodeByOffset,
    getTreeStats,
    exportAsSvg,
    exportAsPng
};
