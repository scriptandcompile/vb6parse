// Benchmark data loading and display for VB6Parse documentation

// Format nanoseconds to readable time
function formatTime(ns) {
    if (ns < 1000) {
        return `${ns.toFixed(2)} ns`;
    } else if (ns < 1000000) {
        return `${(ns / 1000).toFixed(2)} μs`;
    } else if (ns < 1000000000) {
        return `${(ns / 1000000).toFixed(2)} ms`;
    } else {
        return `${(ns / 1000000000).toFixed(2)} s`;
    }
}

// Format benchmark name for display (e.g., "load multiple projects_27" -> "Load Multiple Projects #27")
function formatBenchmarkName(name) {
    const match = name.match(/^(.+?)(?:_(\d+))?$/);
    if (!match) return name;
    
    const baseName = match[1];
    const runNumber = match[2];
    
    // Capitalize each word in base name
    const formatted = baseName.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1)
    ).join(' ');
    
    if (runNumber) {
        return `${formatted} #${runNumber}`;
    }
    return formatted;
}

// Generate anchor ID for a benchmark
function getBenchmarkId(name) {
    return 'benchmark-' + name.replace(/[^a-z0-9]/gi, '-').toLowerCase();
}

// Load and display benchmark data
async function loadBenchmarks() {
    try {
        const response = await fetch('benchmarks.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        displayBenchmarks(data);
    } catch (error) {
        showError(error.message);
    }
}

function displayBenchmarks(data) {
    const benchmarks = data.benchmarks || [];
    
    if (benchmarks.length === 0) {
        showError('No benchmark data available');
        return;
    }

    // Group benchmarks by base name
    const grouped = {};
    benchmarks.forEach(benchmark => {
        // Extract base name (everything before _number)
        const match = benchmark.name.match(/^(.+?)(?:_\d+)?$/);
        const baseName = match ? match[1] : benchmark.name;
        
        if (!grouped[baseName]) {
            grouped[baseName] = [];
        }
        grouped[baseName].push(benchmark);
    });

    // Sort grouped benchmarks numerically by run number
    Object.keys(grouped).forEach(groupName => {
        grouped[groupName].sort((a, b) => {
            const aMatch = a.name.match(/_(\d+)$/);
            const bMatch = b.name.match(/_(\d+)$/);
            const aNum = aMatch ? parseInt(aMatch[1]) : 0;
            const bNum = bMatch ? parseInt(bMatch[1]) : 0;
            return aNum - bNum;
        });
    });

    // Calculate statistics
    const times = benchmarks.map(b => b.mean);
    const avgTime = times.reduce((a, b) => a + b, 0) / times.length;
    const fastest = benchmarks.reduce((min, b) => b.mean < min.mean ? b : min);
    const slowest = benchmarks.reduce((max, b) => b.mean > max.mean ? b : max);

    // Update summary cards
    document.getElementById('benchmark-count').textContent = benchmarks.length;
    document.getElementById('avg-time').textContent = formatTime(avgTime);
    document.getElementById('fastest-time').textContent = formatTime(fastest.mean);
    document.getElementById('fastest-name').textContent = formatBenchmarkName(fastest.name);
    document.getElementById('fastest-card').href = `#${getBenchmarkId(fastest.name)}`;
    
    document.getElementById('slowest-time').textContent = formatTime(slowest.mean);
    document.getElementById('slowest-name').textContent = formatBenchmarkName(slowest.name);
    document.getElementById('slowest-card').href = `#${getBenchmarkId(slowest.name)}`;

    // Display grouped benchmark cards
    const container = document.getElementById('benchmark-cards');
    container.innerHTML = Object.keys(grouped).sort().map(groupName => {
        const group = grouped[groupName];
        const hasMultiple = group.length > 1;
        
        // Always create sections for specific benchmark groups, even if only one item
        const alwaysSectionNames = ['load multiple forms', 'load multiple projects', 'load multiple cls files'];
        const shouldBeSection = hasMultiple || alwaysSectionNames.includes(groupName.toLowerCase());
        
        const cardsHtml = group.map((benchmark, idx) => {
            const stdDevPercent = (benchmark.std_dev / benchmark.mean * 100).toFixed(2);
            const runNumber = (hasMultiple || shouldBeSection) ? (benchmark.name.match(/_(\d+)$/) ? `Run #${benchmark.name.match(/_(\d+)$/)[1]}` : 'Run #1') : null;
            
            return `
                <div class="benchmark-card" id="${getBenchmarkId(benchmark.name)}">
                    <div class="benchmark-header">
                        <h3 class="benchmark-name">${(hasMultiple || shouldBeSection) ? runNumber : benchmark.name}</h3>
                    </div>
                    <div class="benchmark-metrics">
                        <div class="metric">
                            <span class="metric-label">Mean</span>
                            <span class="metric-value">${formatTime(benchmark.mean)}</span>
                        </div>
                        <div class="metric">
                            <span class="metric-label">Median</span>
                            <span class="metric-value">${formatTime(benchmark.median)}</span>
                        </div>
                        <div class="metric">
                            <span class="metric-label">Std Dev</span>
                            <span class="metric-value">${formatTime(benchmark.std_dev)} <span class="metric-percent">(±${stdDevPercent}%)</span></span>
                        </div>
                    </div>
                    <div class="benchmark-bar">
                        <div class="benchmark-bar-fill" style="width: ${(benchmark.mean / slowest.mean * 100).toFixed(2)}%"></div>
                    </div>
                </div>
            `;
        }).join('');
        
        if (shouldBeSection) {
            return `
                <div class="benchmark-section">
                    <h2 class="section-header">${groupName.charAt(0).toUpperCase() + groupName.slice(1)}</h2>
                    <div class="benchmark-section-cards">
                        ${cardsHtml}
                    </div>
                </div>
            `;
        } else {
            return cardsHtml;
        }
    }).join('');

    // Setup search filter
    const searchInput = document.getElementById('benchmark-search');
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const sections = document.querySelectorAll('.benchmark-section');
        const standaloneCards = document.querySelectorAll('#benchmark-cards > .benchmark-card');
        
        sections.forEach(section => {
            const header = section.querySelector('.section-header').textContent.toLowerCase();
            section.style.display = header.includes(query) ? 'block' : 'none';
        });
        
        standaloneCards.forEach(card => {
            const name = card.querySelector('.benchmark-name').textContent.toLowerCase();
            card.style.display = name.includes(query) ? 'block' : 'none';
        });
    });

    // Show content, hide loading
    document.getElementById('loading').style.display = 'none';
    document.getElementById('benchmark-content').style.display = 'block';
}

function showError(message) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('error').style.display = 'block';
}

// Load benchmarks on page load
loadBenchmarks();
