// Coverage History Loading and Display for VB6Parse documentation

// GitHub repository URL
const GITHUB_REPO = 'https://github.com/scriptandcompile/vb6parse';

// Format trend badge HTML
function formatTrendBadge(trend, changePercent) {
    if (!trend || trend === 'no_data') {
        return '<span class="trend-badge stable">→ No Data</span>';
    }
    
    const icons = {
        'improving': '↑',
        'degrading': '↓',
        'stable': '→'
    };
    
    const icon = icons[trend] || '→';
    const changeText = Math.abs(changePercent).toFixed(2);
    const sign = changePercent > 0 ? '+' : '';
    
    return `<span class="trend-badge ${trend}">${icon} ${sign}${changeText}%</span>`;
}

// Load and display coverage with history
async function loadCoverageWithHistory() {
    try {
        // Load current coverage
        const response = await fetch('assets/data/coverage.json');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        
        // Try to load historical data
        let history = null;
        try {
            const historyResponse = await fetch('assets/data/coverage-history.json');
            if (historyResponse.ok) {
                history = await historyResponse.json();
            }
        } catch (e) {
            console.log('No historical data available yet');
        }
        
        displayCoverage(data, history);
        
        // Render historical trends if available
        if (history && history.snapshots && history.snapshots.length >= 2) {
            renderTrendsChart(history, 7); // Default to 7 days
            renderTestChart(history);
            renderSnapshotsTable(history);
            setupTimeRangeSelector(history);
            setupCSVExport(history);
        }
    } catch (error) {
        console.error('Error in loadCoverageWithHistory:', error);
        showError(error.message);
    }
}

function displayCoverage(data, history) {
    const totals = data.data[0].totals;
    const files = data.data[0].files;

    // Display summary with trends
    displaySummaryWithTrend('line', totals.lines, history);
    displaySummaryWithTrend('function', totals.functions, history);
    displaySummaryWithTrend('region', totals.regions, history);

    // Check for regression and show alert if needed
    if (history && history.coverage_summary) {
        checkForRegression(history.coverage_summary);
    }

    // Display file table
    displayFiles(files);

    // Show content, hide loading
    document.getElementById('loading').style.display = 'none';
    document.getElementById('coverage-content').style.display = 'block';
}

function displaySummaryWithTrend(type, data, history) {
    const percent = data.percent.toFixed(2);
    document.getElementById(`${type}-percent`).textContent = `${percent}%`;
    document.getElementById(`${type}-details`).textContent = 
        `${data.covered.toLocaleString()} / ${data.count.toLocaleString()} ${type}s`;
    
    // Add trend indicator if history available
    if (history && history.coverage_summary) {
        const coverageKey = `${type}_coverage`;
        const summary = history.coverage_summary[coverageKey];
        const trendElement = document.getElementById(`${type}-trend`);
        if (trendElement && summary && summary.trend && summary.change_percent !== undefined) {
            trendElement.innerHTML = formatTrendBadge(summary.trend, summary.change_percent);
        }
    }
    
    // Animate bar
    setTimeout(() => {
        document.getElementById(`${type}-bar`).style.width = `${percent}%`;
    }, 100);
}

function checkForRegression(summary) {
    const alertElement = document.getElementById('regression-alert');
    const messageElement = document.getElementById('regression-message');
    
    // Define regression threshold (0.5% drop)
    const REGRESSION_THRESHOLD = -0.5;
    
    const regressions = [];
    ['line', 'function', 'region'].forEach(type => {
        const coverageKey = `${type}_coverage`;
        if (summary[coverageKey] && summary[coverageKey].trend === 'degrading' && summary[coverageKey].change_percent < REGRESSION_THRESHOLD) {
            regressions.push(`${type}: ${summary[coverageKey].change_percent.toFixed(2)}%`);
        }
    });
    
    if (regressions.length > 0) {
        messageElement.textContent = `Coverage has decreased: ${regressions.join(', ')}`;
        alertElement.style.display = 'flex';
    }
}

function getCoverageBadge(percent) {
    const p = parseFloat(percent);
    let className = 'coverage-poor';
    if (p >= 80) className = 'coverage-excellent';
    else if (p >= 60) className = 'coverage-good';
    return `<span class="coverage-badge ${className}">${p.toFixed(1)}%</span>`;
}

function displayFiles(files) {
    const tbody = document.getElementById('file-table-body');
    const completeTbody = document.getElementById('complete-file-table-body');
    const allFiles = [];

    files.forEach(file => {
        const fileName = file.filename.replace(/^.*\/src\//, 'src/');
        const lines = file.summary.lines;
        const functions = file.summary.functions;
        const regions = file.summary.regions;

        allFiles.push({
            name: fileName,
            lines: lines,
            functions: functions,
            regions: regions,
            isComplete: lines.percent === 100 && functions.percent === 100 && regions.percent === 100
        });
    });

    // Group files by directory and calculate directory totals
    const groupedFiles = {};
    const completeGroupedFiles = {};
    
    allFiles.forEach(file => {
        let dir = 'src/';
        if (file.name.includes('/')) {
            const parts = file.name.split('/');
            if (parts.length > 2) {
                dir = parts.slice(0, 2).join('/');
            }
        }
        
        const targetGroup = file.isComplete ? completeGroupedFiles : groupedFiles;
        
        if (!targetGroup[dir]) {
            targetGroup[dir] = {
                files: [],
                totalLines: { covered: 0, count: 0 },
                totalFunctions: { covered: 0, count: 0 },
                totalRegions: { covered: 0, count: 0 }
            };
        }
        targetGroup[dir].files.push(file);
        
        targetGroup[dir].totalLines.covered += file.lines.covered;
        targetGroup[dir].totalLines.count += file.lines.count;
        targetGroup[dir].totalFunctions.covered += file.functions.covered;
        targetGroup[dir].totalFunctions.count += file.functions.count;
        targetGroup[dir].totalRegions.covered += file.regions.covered;
        targetGroup[dir].totalRegions.count += file.regions.count;
    });

    // Calculate percentages for each directory
    [groupedFiles, completeGroupedFiles].forEach(grouped => {
        Object.keys(grouped).forEach(dir => {
            const group = grouped[dir];
            group.totalLines.percent = group.totalLines.count > 0 
                ? (group.totalLines.covered / group.totalLines.count * 100) 
                : 0;
            group.totalFunctions.percent = group.totalFunctions.count > 0 
                ? (group.totalFunctions.covered / group.totalFunctions.count * 100) 
                : 0;
            group.totalRegions.percent = group.totalRegions.count > 0 
                ? (group.totalRegions.covered / group.totalRegions.count * 100) 
                : 0;
            group.avgCoverage = (group.totalLines.percent + group.totalFunctions.percent + group.totalRegions.percent) / 3;
        });
    });

    // Function to render a table
    function renderTable(targetTbody, grouped, sortAscending = true) {
        const sortedDirs = Object.keys(grouped).sort((a, b) => {
            return sortAscending 
                ? grouped[a].avgCoverage - grouped[b].avgCoverage
                : grouped[b].avgCoverage - grouped[a].avgCoverage;
        });

        sortedDirs.forEach(dir => {
            const group = grouped[dir];
            
            const dirRow = document.createElement('tr');
            dirRow.className = 'directory-header';
            dirRow.innerHTML = `
                <td class="directory-name">
                    <span class="directory-link">📁 ${dir}</span>
                </td>
                <td>${getCoverageBadge(group.totalLines.percent)}</td>
                <td>${getCoverageBadge(group.totalFunctions.percent)}</td>
                <td>${getCoverageBadge(group.totalRegions.percent)}</td>
            `;
            targetTbody.appendChild(dirRow);

            group.files.sort((a, b) => sortAscending 
                ? a.lines.percent - b.lines.percent
                : b.lines.percent - a.lines.percent);

            group.files.forEach(file => {
                const row = document.createElement('tr');
                row.className = 'file-row';
                const displayName = file.name.replace(dir + '/', '');
                const filePathWithoutExt = file.name.replace(/\.(rs|toml|md|txt|json|yml|yaml)$/, '');
                const coverageFileUrl = `assets/coverage/${filePathWithoutExt}.html`;
                
                row.innerHTML = `
                    <td class="file-name">
                        &nbsp;&nbsp;<a href="${coverageFileUrl}" class="file-link">${displayName}</a>
                    </td>
                    <td>${getCoverageBadge(file.lines.percent)}</td>
                    <td>${getCoverageBadge(file.functions.percent)}</td>
                    <td>${getCoverageBadge(file.regions.percent)}</td>
                `;
                row.dataset.fullName = file.name;
                targetTbody.appendChild(row);
            });
        });
    }

    renderTable(tbody, groupedFiles, true);
    renderTable(completeTbody, completeGroupedFiles, false);

    if (Object.keys(completeGroupedFiles).length > 0) {
        document.getElementById('complete-coverage-section').style.display = 'block';
    }

    function setupSearch(searchId, tableBodyId) {
        document.getElementById(searchId).addEventListener('input', (e) => {
            const search = e.target.value.toLowerCase();
            const targetTbody = document.getElementById(tableBodyId);
            const dirHeaders = targetTbody.querySelectorAll('.directory-header');
            
            const fileRows = targetTbody.querySelectorAll('.file-row');
            fileRows.forEach(row => {
                const fileName = row.dataset.fullName.toLowerCase();
                row.style.display = fileName.includes(search) ? '' : 'none';
            });
            
            dirHeaders.forEach(header => {
                let nextSibling = header.nextElementSibling;
                let hasVisibleFiles = false;
                
                while (nextSibling && nextSibling.classList.contains('file-row')) {
                    if (nextSibling.style.display !== 'none') {
                        hasVisibleFiles = true;
                        break;
                    }
                    nextSibling = nextSibling.nextElementSibling;
                }
                
                header.style.display = hasVisibleFiles ? '' : 'none';
            });
        });
    }
    
    setupSearch('file-search', 'file-table-body');
    setupSearch('complete-file-search', 'complete-file-table-body');
}

// Render coverage trends chart
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
    
    // Prepare datasets for line, function, and region coverage
    const datasets = [
        {
            label: 'Line Coverage',
            data: snapshots.map(s => ({
                x: new Date(s.timestamp),
                y: s.coverage.line_coverage
            })),
            borderColor: '#3b82f6',
            backgroundColor: '#3b82f620',
            tension: 0.4,
            borderWidth: 3,
            pointRadius: 4,
            pointHoverRadius: 6
        },
        {
            label: 'Function Coverage',
            data: snapshots.map(s => ({
                x: new Date(s.timestamp),
                y: s.coverage.function_coverage
            })),
            borderColor: '#8b5cf6',
            backgroundColor: '#8b5cf620',
            tension: 0.4,
            borderWidth: 3,
            pointRadius: 4,
            pointHoverRadius: 6
        },
        {
            label: 'Region Coverage',
            data: snapshots.map(s => ({
                x: new Date(s.timestamp),
                y: s.coverage.region_coverage
            })),
            borderColor: '#10b981',
            backgroundColor: '#10b98120',
            tension: 0.4,
            borderWidth: 3,
            pointRadius: 4,
            pointHoverRadius: 6
        }
    ];
    
    // Destroy existing chart if any
    if (window.coverageTrendsChart) {
        window.coverageTrendsChart.destroy();
    }
    
    // Store snapshots for click handling
    window.coverageChartSnapshots = snapshots;
    
    // Create new chart
    const ctx = canvas.getContext('2d');
    window.coverageTrendsChart = new Chart(ctx, {
        type: 'line',
        data: { datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            onClick: (event, elements) => {
                if (elements.length > 0) {
                    const index = elements[0].index;
                    const snapshot = window.coverageChartSnapshots[index];
                    if (snapshot && snapshot.commit_sha && snapshot.commit_sha !== 'unknown') {
                        window.open(`${GITHUB_REPO}/commit/${snapshot.commit_sha}`, '_blank');
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 15,
                        font: {
                            size: 12,
                            weight: 'bold'
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            if (context.length > 0) {
                                const snapshot = window.coverageChartSnapshots[context[0].dataIndex];
                                if (snapshot) {
                                    const date = new Date(snapshot.timestamp).toLocaleDateString('en-US', {
                                        year: 'numeric',
                                        month: 'short',
                                        day: 'numeric',
                                        hour: '2-digit',
                                        minute: '2-digit'
                                    });
                                    return date;
                                }
                            }
                            return '';
                        },
                        beforeBody: function(context) {
                            if (context.length > 0) {
                                const snapshot = window.coverageChartSnapshots[context[0].dataIndex];
                                if (snapshot && snapshot.commit_sha && snapshot.commit_sha !== 'unknown') {
                                    const commitShort = snapshot.commit_sha.substring(0, 8);
                                    const commitMsg = snapshot.commit_message.substring(0, 60);
                                    return [
                                        `Commit: ${commitShort}`,
                                        `${commitMsg}${snapshot.commit_message.length > 60 ? '...' : ''}`,
                                        ''
                                    ];
                                }
                            }
                            return [];
                        },
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + '%';
                        },
                        footer: function(context) {
                            if (context.length > 0) {
                                const snapshot = window.coverageChartSnapshots[context[0].dataIndex];
                                if (snapshot) {
                                    const totalTests = Object.values(snapshot.tests || {}).reduce((a, b) => a + b, 0);
                                    const footer = [`Total Tests: ${totalTests.toLocaleString()}`];
                                    if (snapshot.commit_sha && snapshot.commit_sha !== 'unknown') {
                                        footer.push('(click to view on GitHub)');
                                    }
                                    return footer;
                                }
                            }
                            return '';
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'time',
                    time: {
                        unit: daysRange <= 7 ? 'day' : daysRange <= 30 ? 'day' : daysRange <= 90 ? 'week' : 'month',
                        displayFormats: {
                            day: 'MMM d',
                            week: 'MMM d',
                            month: 'MMM yyyy'
                        }
                    },
                    title: {
                        display: true,
                        text: 'Date',
                        font: {
                            weight: 'bold'
                        }
                    }
                },
                y: {
                    min: Math.max(0, Math.min(...datasets.flatMap(d => d.data.map(p => p.y))) - 5),
                    max: 100,
                    title: {
                        display: true,
                        text: 'Coverage (%)',
                        font: {
                            weight: 'bold'
                        }
                    },
                    ticks: {
                        callback: function(value) {
                            return value.toFixed(0) + '%';
                        }
                    }
                }
            }
        }
    });
}

// Render test count chart
function renderTestChart(history) {
    const testSection = document.getElementById('test-trends');
    const canvas = document.getElementById('test-chart');
    
    if (!canvas || !history || !history.snapshots || history.snapshots.length < 2) {
        if (testSection) testSection.style.display = 'none';
        return;
    }
    
    testSection.style.display = 'block';
    
    const snapshots = history.snapshots;
    
    // Prepare stacked bar chart data
    const testCategories = ['lib_tests', 'doc_tests', 'integration_tests', 'fuzz_targets'];
    const labels = ['Library', 'Doc', 'Integration', 'Fuzz'];
    const colors = ['#3b82f6', '#8b5cf6', '#ec4899', '#f59e0b'];
    
    const datasets = testCategories.map((category, idx) => ({
        label: labels[idx],
        data: snapshots.map(s => s.tests[category] || 0),
        backgroundColor: colors[idx],
        borderColor: colors[idx],
        borderWidth: 1
    }));
    
    // Destroy existing chart if any
    if (window.testCountChart) {
        window.testCountChart.destroy();
    }
    
    // Store snapshots for click handling
    window.testChartSnapshots = snapshots;
    
    const ctx = canvas.getContext('2d');
    window.testCountChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: snapshots.map(s => new Date(s.timestamp)),
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            onClick: (event, elements) => {
                if (elements.length > 0) {
                    const index = elements[0].index;
                    const snapshot = window.testChartSnapshots[index];
                    if (snapshot && snapshot.commit_sha && snapshot.commit_sha !== 'unknown') {
                        window.open(`${GITHUB_REPO}/commit/${snapshot.commit_sha}`, '_blank');
                    }
                }
            },
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 15,
                        font: {
                            size: 12,
                            weight: 'bold'
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            if (context.length > 0) {
                                const snapshot = window.testChartSnapshots[context[0].dataIndex];
                                if (snapshot) {
                                    const date = new Date(snapshot.timestamp).toLocaleDateString('en-US', {
                                        year: 'numeric',
                                        month: 'short',
                                        day: 'numeric'
                                    });
                                    return date;
                                }
                            }
                            return '';
                        },
                        footer: function(context) {
                            if (context.length > 0) {
                                const snapshot = window.testChartSnapshots[context[0].dataIndex];
                                if (snapshot) {
                                    const total = Object.values(snapshot.tests || {}).reduce((a, b) => a + b, 0);
                                    const footer = [`Total: ${total.toLocaleString()}`];
                                    if (snapshot.commit_sha && snapshot.commit_sha !== 'unknown') {
                                        const commitShort = snapshot.commit_sha.substring(0, 8);
                                        footer.push(`Commit: ${commitShort}`, '(click to view on GitHub)');
                                    }
                                    return footer;
                                }
                            }
                            return '';
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'time',
                    time: {
                        unit: 'day',
                        displayFormats: {
                            day: 'MMM d'
                        }
                    },
                    stacked: true,
                    title: {
                        display: true,
                        text: 'Date',
                        font: {
                            weight: 'bold'
                        }
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Test Count',
                        font: {
                            weight: 'bold'
                        }
                    },
                    ticks: {
                        callback: function(value) {
                            return value.toLocaleString();
                        }
                    }
                }
            }
        }
    });
}

// Render snapshots table
function renderSnapshotsTable(history) {
    const section = document.getElementById('recent-snapshots');
    const tbody = document.getElementById('snapshots-table-body');
    
    if (!tbody || !history || !history.snapshots || history.snapshots.length === 0) {
        if (section) section.style.display = 'none';
        return;
    }
    
    section.style.display = 'block';
    
    // Show last 10 snapshots
    const recentSnapshots = history.snapshots.slice(-10).reverse();
    
    tbody.innerHTML = recentSnapshots.map(snapshot => {
        const date = new Date(snapshot.timestamp).toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        const commitShort = snapshot.commit_sha.substring(0, 8);
        const commitUrl = snapshot.commit_sha !== 'unknown' 
            ? `${GITHUB_REPO}/commit/${snapshot.commit_sha}`
            : '#';
        const totalTests = snapshot.tests.total || Object.values(snapshot.tests || {}).reduce((a, b) => a + b, 0);
        
        return `
            <tr>
                <td>${date}</td>
                <td>${commitShort !== 'unknown' ? `<a href="${commitUrl}" target="_blank" class="commit-link">${commitShort}</a>` : 'unknown'}</td>
                <td>${snapshot.coverage.line_coverage.toFixed(2)}%</td>
                <td>${snapshot.coverage.function_coverage.toFixed(2)}%</td>
                <td>${snapshot.coverage.region_coverage.toFixed(2)}%</td>
                <td>${totalTests.toLocaleString()}</td>
            </tr>
        `;
    }).join('');
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

// Setup CSV export
function setupCSVExport(history) {
    const exportBtn = document.getElementById('export-csv');
    if (!exportBtn || !history) return;
    
    exportBtn.addEventListener('click', () => {
        const snapshots = history.snapshots;
        
        // CSV header
        let csv = 'Date,Commit,Line %,Function %,Region %,Lib Tests,Doc Tests,Integration Tests,Fuzz Targets,Total Tests\n';
        
        // CSV rows
        snapshots.forEach(snapshot => {
            const date = new Date(snapshot.timestamp).toISOString();
            const commit = snapshot.commit_sha;
            const line = snapshot.coverage.line_coverage.toFixed(2);
            const func = snapshot.coverage.function_coverage.toFixed(2);
            const region = snapshot.coverage.region_coverage.toFixed(2);
            const libTests = snapshot.tests.lib_tests || 0;
            const docTests = snapshot.tests.doc_tests || 0;
            const integTests = snapshot.tests.integration_tests || 0;
            const fuzzTargets = snapshot.tests.fuzz_targets || 0;
            const totalTests = snapshot.tests.total || (libTests + docTests + integTests + fuzzTargets);
            
            csv += `${date},${commit},${line},${func},${region},${libTests},${docTests},${integTests},${fuzzTargets},${totalTests}\n`;
        });
        
        // Download CSV
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'vb6parse-coverage-history.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
}

function showError(message) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('error').style.display = 'block';
    document.getElementById('error-message').textContent = message;
}

// Load coverage with history on page load
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', loadCoverageWithHistory);
} else {
    loadCoverageWithHistory();
}
