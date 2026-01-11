// Theme Switcher Script
(function() {
    const THEME_KEY = 'vb6parse-theme';
    const LIGHT_HIGHLIGHT = 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github.min.css';
    const DARK_HIGHLIGHT = 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github-dark.min.css';
    
    // Detect system preference
    function getSystemPreference() {
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            return 'dark';
        } else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) {
            return 'light';
        }
        // Fallback to dark if no preference detected
        return 'dark';
    }
    
    // Get stored theme or detect system preference with fallback to dark
    function getTheme() {
        const savedTheme = localStorage.getItem(THEME_KEY);
        if (savedTheme) {
            return savedTheme;
        }
        // No saved preference, use system preference (with dark fallback)
        return getSystemPreference();
    }
    
    // Save theme preference
    function saveTheme(theme) {
        localStorage.setItem(THEME_KEY, theme);
    }
    
    // Update highlight.js theme
    function updateHighlightTheme(theme) {
        const highlightLink = document.querySelector('link[href*="highlight.js"]');
        if (highlightLink) {
            highlightLink.href = theme === 'dark' ? DARK_HIGHLIGHT : LIGHT_HIGHLIGHT;
        }
    }
    
    // Apply theme to page
    function applyTheme(theme) {
        if (theme === 'dark') {
            document.documentElement.setAttribute('data-theme', 'dark');
            const themeIcon = document.querySelector('.theme-icon');
            if (themeIcon) {
                themeIcon.textContent = '☀️';
            }
        } else {
            document.documentElement.setAttribute('data-theme', 'light');
            const themeIcon = document.querySelector('.theme-icon');
            if (themeIcon) {
                themeIcon.textContent = '🌙';
            }
        }
        updateHighlightTheme(theme);
    }
    
    // Toggle theme
    function toggleTheme() {
        const currentTheme = getTheme();
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        saveTheme(newTheme);
        applyTheme(newTheme);
    }
    
    // Apply theme immediately to prevent flash (before DOM is ready)
    const initialTheme = getTheme();
    document.documentElement.setAttribute('data-theme', initialTheme);
    
    // Listen for system preference changes
    if (window.matchMedia) {
        const darkModeQuery = window.matchMedia('(prefers-color-scheme: dark)');
        darkModeQuery.addEventListener('change', function(e) {
            // Only auto-update if user hasn't manually set a preference
            if (!localStorage.getItem(THEME_KEY)) {
                const newTheme = e.matches ? 'dark' : 'light';
                applyTheme(newTheme);
            }
        });
    }
    
    // Also update the Highlight.js link immediately if it exists
    if (document.readyState === 'loading') {
        // Document still loading, use DOMContentLoaded
        document.addEventListener('DOMContentLoaded', function() {
            applyTheme(initialTheme);
            
            // Add click handler to toggle button
            const toggleButton = document.getElementById('theme-toggle');
            if (toggleButton) {
                toggleButton.addEventListener('click', toggleTheme);
            }
        });
    } else {
        // DOM already loaded (script loaded late)
        applyTheme(initialTheme);
        
        const toggleButton = document.getElementById('theme-toggle');
        if (toggleButton) {
            toggleButton.addEventListener('click', toggleTheme);
        }
    }
    
    // Try to update highlight stylesheet immediately
    updateHighlightTheme(initialTheme);
})();
