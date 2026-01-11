// VB6 Library Reference Search
// Simple client-side search functionality

let searchIndex = [];
let searchInitialized = false;

// Initialize search when DOM is ready
document.addEventListener('DOMContentLoaded', function() {
    initializeSearch();
});

async function initializeSearch() {
    try {
        const response = await fetch('search-index.json');
        searchIndex = await response.json();
        searchInitialized = true;
        console.log(`Loaded ${searchIndex.length} library items for search`);
    } catch (error) {
        console.error('Failed to load search index:', error);
    }
}

function searchLibrary(query) {
    if (!searchInitialized || !query || query.length < 2) {
        return [];
    }
    
    const lowerQuery = query.toLowerCase();
    const results = [];
    
    for (const item of searchIndex) {
        let score = 0;
        const lowerName = item.name.toLowerCase();
        const lowerDesc = item.description.toLowerCase();
        
        // Exact name match (highest priority)
        if (lowerName === lowerQuery) {
            score = 100;
        }
        // Name starts with query
        else if (lowerName.startsWith(lowerQuery)) {
            score = 80;
        }
        // Name contains query
        else if (lowerName.includes(lowerQuery)) {
            score = 60;
        }
        // Description contains query
        else if (lowerDesc.includes(lowerQuery)) {
            score = 30;
        }
        // Category contains query
        else if (item.category.toLowerCase().includes(lowerQuery)) {
            score = 20;
        }
        
        if (score > 0) {
            results.push({ ...item, score });
        }
    }
    
    // Sort by score (descending)
    results.sort((a, b) => b.score - a.score);
    
    return results.slice(0, 20); // Return top 20 results
}

function displaySearchResults(results, query) {
    const container = document.getElementById('search-results');
    if (!container) return;
    
    if (results.length === 0) {
        container.innerHTML = `
            <div class="no-results">
                <p>No results found for "${escapeHtml(query)}"</p>
            </div>
        `;
        return;
    }
    
    let html = `<h3>Search Results (${results.length})</h3><div class="search-results-list">`;
    
    for (const result of results) {
        const typeIcon = result.type === 'function' ? '📊' : '⚡';
        html += `
            <div class="search-result-item">
                <h4>
                    <a href="${result.url}">
                        ${typeIcon} <code>${escapeHtml(result.name)}</code>
                    </a>
                </h4>
                <p class="result-meta">${escapeHtml(result.type)} • ${escapeHtml(result.category)}</p>
                <p class="result-description">${escapeHtml(result.description.substring(0, 150))}...</p>
            </div>
        `;
    }
    
    html += '</div>';
    container.innerHTML = html;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Export for use in search page
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { searchLibrary, displaySearchResults };
}
