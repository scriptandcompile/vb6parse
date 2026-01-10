// Load test statistics
(function() {
    async function loadStats() {
        try {
            const response = await fetch('stats.json');
            if (!response.ok) {
                console.log('Stats not available - stats.json not found');
                return;
            }
            
            const stats = await response.json();
            
            // Update test count displays
            const testCountElements = document.querySelectorAll('#test-count-header, .test-count');
            testCountElements.forEach(el => {
                if (el) {
                    el.textContent = stats.test_count.toLocaleString();
                }
            });
            
            // Update rounded test count displays (rounded down to nearest hundred)
            const roundedTestCountElements = document.querySelectorAll('.test-count-rounded');
            roundedTestCountElements.forEach(el => {
                if (el) {
                    const roundedCount = Math.floor(stats.test_count / 100) * 100;
                    el.textContent = roundedCount.toLocaleString();
                }
            });
            
            // Update stat cards with data-stat attributes
            let statsLoaded = false;
            document.querySelectorAll('[data-stat]').forEach(el => {
                const statKey = el.getAttribute('data-stat');
                if (stats[statKey] !== undefined) {
                    el.textContent = stats[statKey].toLocaleString();
                    statsLoaded = true;
                }
            });
            
            // Show stats grid and hide loading message if stats were loaded
            if (statsLoaded) {
                const statsGrid = document.getElementById('stats-grid');
                const statsLoading = document.getElementById('stats-loading');
                if (statsGrid) statsGrid.style.display = 'grid';
                if (statsLoading) statsLoading.style.display = 'none';
            }
        } catch (error) {
            console.log('Stats not available:', error);
        }
    }
    
    // Load stats when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', loadStats);
    } else {
        loadStats();
    }
})();
