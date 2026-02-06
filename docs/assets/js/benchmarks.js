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

// Parse benchmark name into components
// Format: {size}_{filename} (e.g., "large_Curves.frm", "medium_FastDrawing.cls")
function parseBenchmarkName(name) {
    const match = name.match(/^(small|medium|large)_(.+)$/);
    if (!match) {
        return {
            size: 'unknown',
            filename: name,
            baseName: name,
            extension: '',
            displayName: name
        };
    }
    
    const size = match[1];
    const filename = match[2];
    const extMatch = filename.match(/\.([a-z]+)$/i);
    const extension = extMatch ? extMatch[1].toLowerCase() : '';
    
    // Remove extension for base name
    const baseName = extension ? filename.slice(0, -(extension.length + 1)) : filename;
    
    return {
        size: size,
        filename: filename,
        baseName: baseName,
        extension: extension,
        displayName: filename
    };
}

// Get file type category from extension
function getFileTypeCategory(extension) {
    const categories = {
        'vbp': 'Project Files',
        'cls': 'Class Files',
        'bas': 'Module Files',
        'frm': 'Form Files',
        'frx': 'Form Resources'
    };
    return categories[extension] || 'Other';
}

// Format benchmark name for display
function formatBenchmarkName(name) {
    const parsed = parseBenchmarkName(name);
    return parsed.displayName;
}

// Generate anchor ID for a benchmark
function getBenchmarkId(name) {
    return 'benchmark-' + name.replace(/[^a-z0-9]/gi, '-').toLowerCase();
}

// Get trend information for a benchmark
function getTrendInfo(benchmarkName, history) {
    if (!history || !history.benchmarks_summary) {
        return null;
    }
    
    const summary = history.benchmarks_summary[benchmarkName];
    if (!summary || !summary.trend) {
        return null;
    }
    
    return summary.trend;
}

// Format trend badge HTML
function formatTrendBadge(trend) {
    if (!trend) {
        return '';
    }
    
    const icons = {
        'improving': '📈 ↓',
        'degrading': '📉 ↑',
        'stable': '→'
    };
    
    const colors = {
        'improving': '#10b981',
        'degrading': '#ef4444',
        'stable': '#6b7280'
    };
    
    const icon = icons[trend.direction] || '→';
    const color = colors[trend.direction] || '#6b7280';
    const changeText = Math.abs(trend.change_percent).toFixed(2);
    
    return `
        <span class="trend-badge" style="background-color: ${color}20; color: ${color}; border: 1px solid ${color}40;">
            ${icon} ${changeText}%
        </span>
    `;
}

// Load and display benchmark data
async function loadBenchmarks() {
    try {
        // Load current snapshot
        const response = await fetch('assets/data/benchmarks.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        
        // Try to load historical data
        let history = null;
        try {
            const historyResponse = await fetch('assets/data/benchmarks-history.json');
            if (historyResponse.ok) {
                history = await historyResponse.json();
            }
        } catch (e) {
            console.log('No historical data available yet');
        }
        
        displayBenchmarks(data, history);
        
        // Render historical trends chart if history is available
        if (history && history.snapshots && history.snapshots.length >= 2) {
            renderTrendsChart(history, 30); // Default to 30 days
            setupTimeRangeSelector(history);
        }
    } catch (error) {
        showError(error.message);
    }
}

function displayBenchmarks(data, history) {
    const benchmarks = data.benchmarks || [];
    
    if (benchmarks.length === 0) {
        showError('No benchmark data available');
        return;
    }
    
    // Add history info to summary if available
    if (history && history.snapshots) {
        const snapshotCount = history.snapshots.length;
        const lastUpdated = history.last_updated;
        const summaryContainer = document.getElementById('overall-summary');
        
        // Add history info card
        if (snapshotCount > 1) {
            const historyCard = document.createElement('div');
            historyCard.className = 'summary-card history-card';
            historyCard.innerHTML = `
                <h3>📊 Historical Data</h3>
                <div class="summary-value">${snapshotCount}</div>
                <div class="summary-label">snapshots</div>
                <div class="summary-sublabel" style="font-size: 0.8em; opacity: 0.7; margin-top: 4px;">
                    Last: ${new Date(lastUpdated).toLocaleDateString()}
                </div>
            `;
            summaryContainer.insertBefore(historyCard, summaryContainer.firstChild);
        }
    }

    // Group benchmarks by file type and size, then by base name
    const grouped = {};
    benchmarks.forEach(benchmark => {
        const parsed = parseBenchmarkName(benchmark.name);
        const fileType = getFileTypeCategory(parsed.extension);
        const size = parsed.size.charAt(0).toUpperCase() + parsed.size.slice(1); // Capitalize
        const groupKey = `${fileType} - ${size}`;
        
        if (!grouped[groupKey]) {
            grouped[groupKey] = {};
        }
        
        // Group by base filename within the category
        if (!grouped[groupKey][parsed.baseName]) {
            grouped[groupKey][parsed.baseName] = [];
        }
        grouped[groupKey][parsed.baseName].push({
            ...benchmark,
            parsed: parsed
        });
    });

    // Sort benchmarks within each group by mean time
    Object.keys(grouped).forEach(groupKey => {
        Object.keys(grouped[groupKey]).forEach(baseName => {
            grouped[groupKey][baseName].sort((a, b) => a.mean - b.mean);
        });
    });

    // Calculate statistics
    const times = benchmarks.map(b => b.mean);
    const avgTime = times.reduce((a, b) => a + b, 0) / times.length;
    const fastest = benchmarks.reduce((min, b) => b.mean < min.mean ? b : min);
    const slowest = benchmarks.reduce((max, b) => b.mean > max.mean ? b : max);

    // Calculate statistics per file type
    const fileTypeStats = {};
    Object.keys(grouped).forEach(groupKey => {
        const group = grouped[groupKey];
        const groupBenchmarks = [];
        Object.keys(group).forEach(baseName => {
            groupBenchmarks.push(...group[baseName]);
        });
        
        const [fileType] = groupKey.split(' - ');
        if (!fileTypeStats[fileType]) {
            fileTypeStats[fileType] = {
                count: 0,
                benchmarks: []
            };
        }
        fileTypeStats[fileType].count += groupBenchmarks.length;
        fileTypeStats[fileType].benchmarks.push(...groupBenchmarks);
    });

    // Update overall summary cards with file type breakdown
    const summaryContainer = document.getElementById('overall-summary');
    summaryContainer.innerHTML = `
        <div class="summary-card">
            <h3>Total Benchmarks</h3>
            <div class="summary-value">${benchmarks.length}</div>
        </div>
        ${Object.keys(fileTypeStats).sort((a, b) => {
            const order = {
                'Project Files': 1,
                'Class Files': 2,
                'Module Files': 3,
                'Form Files': 4,
                'Form Resources': 5,
                'Other': 6
            };
            return (order[a] || 999) - (order[b] || 999);
        }).map(fileType => {
            const stats = fileTypeStats[fileType];
            const icon = {
                'Project Files': '📦',
                'Class Files': '📋',
                'Module Files': '📄',
                'Form Files': '🖼️',
                'Form Resources': '🗂️',
                'Other': '📁'
            }[fileType] || '📁';
            
            return `
                <div class="summary-card type-card">
                    <h3>${icon} ${fileType}</h3>
                    <div class="summary-value">${stats.count}</div>
                    <div class="summary-label">benchmarks</div>
                </div>
            `;
        }).join('')}
    `;
    const container = document.getElementById('benchmark-cards');
    
    // Sort group keys (e.g., "Class Files - Large", "Form Files - Medium")
    const sortedGroupKeys = Object.keys(grouped).sort((a, b) => {
        // Extract file type and size
        const [typeA, sizeA] = a.split(' - ');
        const [typeB, sizeB] = b.split(' - ');
        
        // Define sort order for file types
        const typeOrder = {
            'Project Files': 1,
            'Class Files': 2,
            'Module Files': 3,
            'Form Files': 4,
            'Form Resources': 5,
            'Other': 6
        };
        
        // Define sort order for sizes
        const sizeOrder = { 'Small': 1, 'Medium': 2, 'Large': 3, 'Unknown': 4 };
        
        // First sort by type, then by size
        const typeCompare = (typeOrder[typeA] || 999) - (typeOrder[typeB] || 999);
        if (typeCompare !== 0) return typeCompare;
        
        return (sizeOrder[sizeA] || 999) - (sizeOrder[sizeB] || 999);
    });
    
    container.innerHTML = sortedGroupKeys.map(groupKey => {
        const group = grouped[groupKey];
        const baseNames = Object.keys(group).sort();
        
        // Calculate section statistics
        const sectionBenchmarks = [];
        baseNames.forEach(baseName => {
            sectionBenchmarks.push(...group[baseName]);
        });
        
        const sectionTimes = sectionBenchmarks.map(b => b.mean);
        const sectionAvg = sectionTimes.reduce((a, b) => a + b, 0) / sectionTimes.length;
        const sectionFastest = sectionBenchmarks.reduce((min, b) => b.mean < min.mean ? b : min);
        const sectionSlowest = sectionBenchmarks.reduce((max, b) => b.mean > max.mean ? b : max);
        
        // Generate section summary cards
        const sectionSummary = `
            <div class="section-summary">
                <div class="section-stat">
                    <span class="stat-label">Count</span>
                    <span class="stat-value">${sectionBenchmarks.length}</span>
                </div>
                <div class="section-stat">
                    <span class="stat-label">Average</span>
                    <span class="stat-value">${formatTime(sectionAvg)}</span>
                </div>
                <a href="#${getBenchmarkId(sectionFastest.name)}" class="section-stat section-stat-link">
                    <span class="stat-label">Fastest</span>
                    <span class="stat-value">${formatTime(sectionFastest.mean)}</span>
                    <span class="stat-sublabel">${sectionFastest.parsed.displayName}</span>
                </a>
                <a href="#${getBenchmarkId(sectionSlowest.name)}" class="section-stat section-stat-link">
                    <span class="stat-label">Slowest</span>
                    <span class="stat-value">${formatTime(sectionSlowest.mean)}</span>
                    <span class="stat-sublabel">${sectionSlowest.parsed.displayName}</span>
                </a>
            </div>
        `;
        
        const cardsHtml = baseNames.map(baseName => {
            const items = group[baseName];
            const hasMultiple = items.length > 1;
            
            // If multiple benchmarks for same file, show them as a sub-section
            if (hasMultiple) {
                const itemsHtml = items.map((benchmark, idx) => {
                    const stdDevPercent = (benchmark.std_dev / benchmark.mean * 100).toFixed(2);
                    
                    return `
                        <div class="benchmark-card" id="${getBenchmarkId(benchmark.name)}">
                            <div class="benchmark-header">
                                <h4 class="benchmark-name">Run #${idx + 1}</h4>
                                ${formatTrendBadge(getTrendInfo(benchmark.name, history))}
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
                
                return `
                    <div class="benchmark-subsection">
                        <h3 class="subsection-header">${items[0].parsed.displayName}</h3>
                        <div class="benchmark-subsection-cards">
                            ${itemsHtml}
                        </div>
                    </div>
                `;
            } else {
                // Single benchmark - display directly
                const benchmark = items[0];
                const stdDevPercent = (benchmark.std_dev / benchmark.mean * 100).toFixed(2);
                
                return `
                    <div class="benchmark-card" id="${getBenchmarkId(benchmark.name)}">
                        <div class="benchmark-header">
                            <h3 class="benchmark-name">${benchmark.parsed.displayName}</h3>
                            ${formatTrendBadge(getTrendInfo(benchmark.name, history))}
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
            }
        }).join('');
        
        return `
            <div class="benchmark-section">
                <h2 class="section-header">${groupKey}</h2>
                ${sectionSummary}
                <div class="benchmark-section-cards">
                    ${cardsHtml}
                </div>
            </div>
        `;
    }).join('');

    // Setup search filter
    const searchInput = document.getElementById('benchmark-search');
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const sections = document.querySelectorAll('.benchmark-section');
        
        sections.forEach(section => {
            const header = section.querySelector('.section-header').textContent.toLowerCase();
            const subsections = section.querySelectorAll('.benchmark-subsection');
            const cards = section.querySelectorAll('.benchmark-card');
            
            let hasVisibleContent = false;
            
            // Check subsections
            subsections.forEach(subsection => {
                const subsectionHeader = subsection.querySelector('.subsection-header').textContent.toLowerCase();
                const matchesSubsection = subsectionHeader.includes(query);
                subsection.style.display = matchesSubsection ? 'block' : 'none';
                if (matchesSubsection) hasVisibleContent = true;
            });
            
            // Check standalone cards (not in subsections)
            cards.forEach(card => {
                // Skip if card is inside a subsection (already handled above)
                if (card.closest('.benchmark-subsection')) return;
                
                const name = card.querySelector('.benchmark-name').textContent.toLowerCase();
                const matchesCard = name.includes(query);
                card.style.display = matchesCard ? 'block' : 'none';
                if (matchesCard) hasVisibleContent = true;
            });
            
            // Show section if header matches or has visible content
            section.style.display = (header.includes(query) || hasVisibleContent) ? 'block' : 'none';
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

// Render historical trends chart
function renderTrendsChart(history, daysRange) {
    const trendsSection = document.getElementById('historical-trends');
    const canvas = document.getElementById('trends-chart');
    
    if (!canvas || !history || !history.snapshots || history.snapshots.length < 2) {
        if (trendsSection) trendsSection.style.display = 'none';
        return;
    }
    
    trendsSection.style.display = 'block';
    
    // Filter snapshots by date range
    let snapshots = history.snapshots;
    if (daysRange !== 'all') {
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - parseInt(daysRange));
        snapshots = snapshots.filter(s => new Date(s.timestamp) >= cutoffDate);
    }
    
    if (snapshots.length < 2) {
        trendsSection.style.display = 'none';
        return;
    }
    
    // Group benchmarks by file type and calculate averages
    const fileTypes = ['Project Files', 'Class Files', 'Module Files', 'Form Files', 'Form Resources'];
    const datasets = [];
    
    const colors = {
        'Project Files': '#3b82f6',
        'Class Files': '#8b5cf6',
        'Module Files': '#ec4899',
        'Form Files': '#f59e0b',
        'Form Resources': '#10b981'
    };
    
    fileTypes.forEach(fileType => {
        const data = snapshots.map(snapshot => {
            const benchmarks = snapshot.benchmarks.filter(b => {
                const parsed = parseBenchmarkName(b.name);
                return getFileTypeCategory(parsed.extension) === fileType;
            });
            
            if (benchmarks.length === 0) return null;
            
            const avgTime = benchmarks.reduce((sum, b) => sum + b.mean, 0) / benchmarks.length;
            return {
                x: new Date(snapshot.timestamp),
                y: avgTime / 1000000  // Convert to milliseconds
            };
        }).filter(d => d !== null);
        
        if (data.length > 0) {
            datasets.push({
                label: fileType,
                data: data,
                borderColor: colors[fileType],
                backgroundColor: colors[fileType] + '20',
                tension: 0.4,
                borderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 5
            });
        }
    });
    
    // Destroy existing chart if any
    if (window.trendsChartInstance) {
        window.trendsChartInstance.destroy();
    }
    
    // Create new chart
    const ctx = canvas.getContext('2d');
    window.trendsChartInstance = new Chart(ctx, {
        type: 'line',
        data: { datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 15
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + ' ms';
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'time',
                    time: {
                        unit: daysRange <= 30 ? 'day' : daysRange <= 90 ? 'week' : 'month',
                        displayFormats: {
                            day: 'MMM d',
                            week: 'MMM d',
                            month: 'MMM yyyy'
                        }
                    },
                    title: {
                        display: true,
                        text: 'Date'
                    }
                },
                y: {
                    beginAtZero: false,
                    title: {
                        display: true,
                        text: 'Average Time (ms)'
                    },
                    ticks: {
                        callback: function(value) {
                            return value.toFixed(2) + ' ms';
                        }
                    }
                }
            }
        }
    });
}

// Setup time range selector buttons
function setupTimeRangeSelector(history) {
    const buttons = document.querySelectorAll('.time-range-btn');
    buttons.forEach(btn => {
        btn.addEventListener('click', () => {
            buttons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const range = btn.getAttribute('data-range');
            renderTrendsChart(history, range);
        });
    });
}

// Load benchmarks on page load
loadBenchmarks();
