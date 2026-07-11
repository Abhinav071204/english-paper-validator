// Application State
let state = {
    activeDoc: null,
    uploadedFile: null,
    validationResult: null,
    activeTab: 'preview',
    activeFilter: 'all'
};

// API Endpoints
const API_FILES = '/api/files';
const API_VALIDATE_STORED = '/api/validate_stored';
const API_VALIDATE_UPLOAD = '/api/validate';
const API_CORRECT_STORED = '/api/correct_stored';
const API_CORRECT_UPLOAD = '/api/correct';

// DOM Elements
const elements = {
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),
    storedFilesList: document.getElementById('storedFilesList'),
    readabilitySection: document.getElementById('readabilitySection'),
    readabilityScores: document.getElementById('readabilityScores'),
    activeDocName: document.getElementById('activeDocName'),
    activeDocType: document.getElementById('activeDocType'),
    scoreContainer: document.getElementById('scoreContainer'),
    overallScore: document.getElementById('overallScore'),
    scoreCircle: document.getElementById('scoreCircle'),
    statsGrid: document.getElementById('statsGrid'),
    fontViolationsCount: document.getElementById('fontViolationsCount'),
    placeholdersCount: document.getElementById('placeholdersCount'),
    wordingViolationsCount: document.getElementById('wordingViolationsCount'),
    grammarViolationsCount: document.getElementById('grammarViolationsCount'),
    contentBody: document.getElementById('contentBody'),
    docViewer: document.getElementById('docViewer'),
    findingsCount: document.getElementById('findingsCount'),
    findingsList: document.getElementById('findingsList'),
    welcomeScreen: document.getElementById('welcomeScreen'),
    tabButtons: document.querySelectorAll('.tab-btn'),
    tabPanes: document.querySelectorAll('.tab-pane'),
    filterButtons: document.querySelectorAll('.filter-btn'),
    headerActions: document.getElementById('headerActions'),
    correctBtn: document.getElementById('correctBtn')
};

// Initialize Application
function initializeApp() {
    fetchPreloadedFiles();
    setupEventListeners();
}

if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}

// Setup Event Listeners
function setupEventListeners() {
    // 1. Drag & Drop Event Listeners
    elements.dropzone.addEventListener('click', () => elements.fileInput.click());
    
    elements.fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileUpload(e.target.files[0]);
        }
    });

    elements.dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        elements.dropzone.style.borderColor = '#6366f1';
        elements.dropzone.style.backgroundColor = 'rgba(99, 102, 241, 0.08)';
    });

    elements.dropzone.addEventListener('dragleave', () => {
        elements.dropzone.style.borderColor = 'rgba(99, 102, 241, 0.3)';
        elements.dropzone.style.backgroundColor = 'rgba(99, 102, 241, 0.03)';
    });

    elements.dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        elements.dropzone.style.borderColor = 'rgba(99, 102, 241, 0.3)';
        elements.dropzone.style.backgroundColor = 'rgba(99, 102, 241, 0.03)';
        if (e.dataTransfer.files.length > 0) {
            handleFileUpload(e.dataTransfer.files[0]);
        }
    });

    // 2. Tab Switching
    elements.tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            elements.tabButtons.forEach(b => b.classList.remove('active'));
            elements.tabPanes.forEach(p => p.classList.remove('active'));
            
            btn.classList.add('active');
            const tabName = btn.dataset.tab;
            document.getElementById(`pane-${tabName}`).classList.add('active');
            state.activeTab = tabName;
        });
    });

    // 3. Findings Filter
    elements.filterButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            elements.filterButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.activeFilter = btn.dataset.filter;
            renderFindingsList();
        });
    });

    // 4. Auto-Correct & Download Button
    if (elements.correctBtn) {
        elements.correctBtn.addEventListener('click', () => {
            if (!state.activeDoc) return;
            
            elements.correctBtn.textContent = 'Correcting... ⏳';
            elements.correctBtn.disabled = true;
            
            if (state.uploadedFile) {
                const formData = new FormData();
                formData.append('file', state.uploadedFile);
                
                fetch(API_CORRECT_UPLOAD, {
                    method: 'POST',
                    body: formData
                })
                .then(res => {
                    if (!res.ok) throw new Error('Correction API failed');
                    return res.blob();
                })
                .then(blob => {
                    triggerDownload(blob, state.activeDoc.replace('.docx', '_Corrected.docx'));
                    resetCorrectBtn();
                })
                .catch(err => {
                    alert(`Error correcting document: ${err.message}`);
                    resetCorrectBtn();
                });
            } else {
                const url = `${API_CORRECT_STORED}?file=${encodeURIComponent(state.activeDoc)}`;
                fetch(url)
                .then(res => {
                    if (!res.ok) throw new Error('Correction API failed');
                    return res.blob();
                })
                .then(blob => {
                    triggerDownload(blob, state.activeDoc.replace('.docx', '_Corrected.docx'));
                    resetCorrectBtn();
                })
                .catch(err => {
                    alert(`Error correcting document: ${err.message}`);
                    resetCorrectBtn();
                });
            }
        });
    }
}

function triggerDownload(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
}

function resetCorrectBtn() {
    elements.correctBtn.textContent = '🛠️ Auto-Correct & Download';
    elements.correctBtn.disabled = false;
}

// Fetch lists of preloaded files from local backend
function fetchPreloadedFiles() {
    fetch(API_FILES)
        .then(res => res.json())
        .then(data => {
            if (data.files && data.files.length > 0) {
                renderPreloadedFilesList(data.files);
            } else {
                elements.storedFilesList.innerHTML = '<p class="file-limits">No stored docx files found.</p>';
            }
        })
        .catch(err => {
            console.error('Error fetching preloaded files:', err);
            elements.storedFilesList.innerHTML = '<p class="file-limits" style="color:var(--color-placeholder)">Server offline or unreachable</p>';
        });
}

// Render preloaded files list
function renderPreloadedFilesList(files) {
    elements.storedFilesList.innerHTML = '';
    
    // Sort files to order template and key files cleanly
    files.sort().forEach(filename => {
        const btn = document.createElement('button');
        btn.className = 'file-item';
        
        let icon = '📄';
        if (filename.includes('Answer_Key') || filename.includes('AnswerKey')) {
            icon = '🔑';
        } else if (filename.includes('Annotated') || filename.includes('Example')) {
            icon = '📐';
        }
        
        btn.innerHTML = `<span class="file-item-icon">${icon}</span> ${filename}`;
        btn.addEventListener('click', () => loadStoredFile(filename, btn));
        elements.storedFilesList.appendChild(btn);
    });
}

// Handle preloaded file selection
function loadStoredFile(filename, element) {
    // Update active states
    document.querySelectorAll('.file-item').forEach(el => el.classList.remove('active'));
    element.classList.add('active');
    
    state.uploadedFile = null; // Clear uploaded file
    showLoadingState();
    
    fetch(`${API_VALIDATE_STORED}?file=${encodeURIComponent(filename)}`)
        .then(res => res.json())
        .then(data => {
            state.activeDoc = filename;
            state.validationResult = data;
            displayValidationResult(filename);
        })
        .catch(err => {
            console.error(err);
            alert(`Validation failed: ${err.message}`);
            hideLoadingState();
        });
}

// Handle drag-and-drop file upload
function handleFileUpload(file) {
    if (!file.name.endsWith('.docx')) {
        alert('Please upload a Microsoft Word (.docx) document only.');
        return;
    }
    
    // Clear active states in stored sidebar list
    document.querySelectorAll('.file-item').forEach(el => el.classList.remove('active'));
    
    state.uploadedFile = file; // Save uploaded file in state
    showLoadingState();
    
    const formData = new FormData();
    formData.append('file', file);
    
    fetch(API_VALIDATE_UPLOAD, {
        method: 'POST',
        body: formData
    })
        .then(res => res.json())
        .then(data => {
            state.activeDoc = file.name;
            state.validationResult = data;
            displayValidationResult(file.name);
        })
        .catch(err => {
            console.error(err);
            alert(`Upload validation failed: ${err.message}`);
            hideLoadingState();
        });
}

// Display loading indicator
function showLoadingState() {
    elements.welcomeScreen.style.display = 'none';
    elements.contentBody.style.display = 'none';
    elements.statsGrid.style.display = 'none';
    elements.scoreContainer.style.display = 'none';
    elements.headerActions.style.display = 'none';
    
    elements.activeDocName.textContent = 'Analyzing document...';
    elements.activeDocType.textContent = 'Running wording & font checks';
    
    // Inject loading spinner temporarily
    let spinner = document.getElementById('mainLoader');
    if (!spinner) {
        spinner = document.createElement('div');
        spinner.id = 'mainLoader';
        spinner.className = 'loading-spinner';
        elements.activeDocName.parentNode.appendChild(spinner);
    }
    spinner.style.display = 'block';
}

function hideLoadingState() {
    const spinner = document.getElementById('mainLoader');
    if (spinner) spinner.style.display = 'none';
}

// Process and display report details
function displayValidationResult(filename) {
    hideLoadingState();
    
    const result = state.validationResult;
    if (result.error) {
        alert(`Validation Error: ${result.error}`);
        return;
    }
    
    // Show sections
    elements.contentBody.style.display = 'flex';
    elements.statsGrid.style.display = 'grid';
    elements.scoreContainer.style.display = 'block';
    elements.headerActions.style.display = 'flex';
    
    // File details
    elements.activeDocName.textContent = filename;
    let typeText = 'Part 1 Question Paper';
    if (result.doc_type === 'part2_test') typeText = 'Part 2 Graded Assignment';
    if (result.doc_type === 'part2_answer_key') typeText = 'Part 2 Graded Assignment (Answer Key)';
    elements.activeDocType.textContent = typeText;
    
    // Radial Score gauge mapping
    const score = result.stats.overall_compliance_score;
    elements.overallScore.textContent = score;
    elements.scoreCircle.setAttribute('stroke-dasharray', `${score}, 100`);
    
    // Radial color classes
    elements.scoreCircle.className.baseVal = 'circle';
    if (score >= 80) {
        elements.scoreCircle.classList.add('score-good');
    } else if (score >= 50) {
        elements.scoreCircle.classList.add('score-warn');
    } else {
        elements.scoreCircle.classList.add('score-bad');
    }
    
    // Stats cards numbers
    elements.fontViolationsCount.textContent = result.stats.font_violations_count;
    elements.placeholdersCount.textContent = result.stats.placeholders_count;
    elements.wordingViolationsCount.textContent = result.stats.wording_violations_count;
    elements.grammarViolationsCount.textContent = result.stats.grammar_violations_count;
    elements.findingsCount.textContent = result.findings.length;
    
    // Render passages readability list
    renderReadability(result.passages_readability);
    
    // Render tabs contents
    renderDocumentPreview();
    renderFindingsList();
}

// Render readability sidebar card
function renderReadability(readability) {
    const keys = Object.keys(readability);
    if (keys.length === 0) {
        elements.readabilitySection.style.display = 'none';
        return;
    }
    
    elements.readabilitySection.style.display = 'block';
    elements.readabilityScores.innerHTML = '';
    
    keys.forEach(name => {
        const data = readability[name];
        const item = document.createElement('div');
        item.className = 'readability-item';
        
        // Color border based on score range
        if (data.grade < 4.0 || data.grade > 7.0) {
            item.style.borderLeftColor = 'var(--color-font)';
        } else {
            item.style.borderLeftColor = 'var(--color-success)';
        }
        
        item.innerHTML = `
            <div class="readability-header">
                <strong>${name}</strong>
                <span>Grade ${data.grade}</span>
            </div>
            <div class="readability-stats">
                <span>Words: <strong>${data.words}</strong></span>
                <span>Flesch Ease: <strong>${data.ease}</strong></span>
            </div>
        `;
        elements.readabilityScores.appendChild(item);
    });
}

// Escape HTML utility
function escapeHtml(text) {
    if (!text) return '';
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Render document in-context editor view with inline highlights
function renderDocumentPreview() {
    elements.docViewer.innerHTML = '';
    const elementsData = state.validationResult.html_preview;
    
    elementsData.forEach((el, elIndex) => {
        if (el.type === 'paragraph') {
            const p = document.createElement('p');
            
            // Map style names
            const style = el.style.toLowerCase();
            if (style.startsWith('heading 1')) p.className = 'heading-1';
            else if (style.startsWith('heading 2')) p.className = 'heading-2';
            else if (style.startsWith('heading 3')) p.className = 'heading-3';
            
            p.id = `doc-p-${elIndex}`;
            
            // Build paragraph HTML with highlights
            p.innerHTML = applyTextHighlights(el.text, el.findings);
            
            // Block validation styling
            const blockWarnings = el.findings.filter(f => !f.text);
            if (blockWarnings.length > 0) {
                const hasError = blockWarnings.some(f => f.level === 'error');
                p.classList.add(hasError ? 'highlight-block-error' : 'highlight-block-warning');
                p.setAttribute('data-tooltip', blockWarnings.map(f => f.message).join('\n'));
            }
            
            elements.docViewer.appendChild(p);
        } 
        else if (el.type === 'table') {
            const table = document.createElement('table');
            table.id = `doc-t-${elIndex}`;
            
            // Block validation styling
            const blockWarnings = el.findings.filter(f => !f.text);
            if (blockWarnings.length > 0) {
                const hasError = blockWarnings.some(f => f.level === 'error');
                table.classList.add(hasError ? 'highlight-block-error' : 'highlight-block-warning');
                table.setAttribute('data-tooltip', blockWarnings.map(f => f.message).join('\n'));
            }

            const tbody = document.createElement('tbody');
            
            el.cells.forEach((row, rIdx) => {
                const tr = document.createElement('tr');
                row.forEach((cell, cIdx) => {
                    const td = document.createElement('td');
                    td.innerHTML = applyTextHighlights(cell.text, cell.findings);
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
            
            table.appendChild(tbody);
            elements.docViewer.appendChild(table);
        }
    });
}

// Injects highlights spans inside strings matching text runs
function applyTextHighlights(text, findings) {
    if (!text.trim()) return '';
    let escaped = escapeHtml(text);
    
    // Sort findings by length of target text descending (so we replace larger substrings first to prevent corrupting indices)
    const runFindings = findings.filter(f => f.text);
    runFindings.sort((a, b) => b.text.length - a.text.length);
    
    // Keep track of what we have replaced to avoid double-replacing
    const replaced = new Set();
    
    runFindings.forEach(f => {
        const target = escapeHtml(f.text);
        if (!target || replaced.has(target)) return;
        
        // Escape special regex characters in target string
        const escapedTarget = target.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        const regex = new RegExp(escapedTarget, 'g');
        
        const highlightSpan = `<span class="highlight-${f.category}" data-tooltip="${escapeHtml(f.message)}">${target}</span>`;
        
        // Only replace if matching substring is found in the current output
        if (escaped.includes(target)) {
            escaped = escaped.replace(regex, highlightSpan);
            replaced.add(target);
        }
    });
    
    // Replace newline breaks with <br> tags
    return escaped.replace(/\n/g, '<br>');
}

// Render list of detailed findings
function renderFindingsList() {
    elements.findingsList.innerHTML = '';
    const findings = state.validationResult.findings;
    
    const filtered = findings.filter(f => {
        if (state.activeFilter === 'all') return true;
        return f.level === state.activeFilter;
    });
    
    if (filtered.length === 0) {
        elements.findingsList.innerHTML = '<p class="file-limits" style="text-align:center; padding: 2rem;">No findings match the selected filter.</p>';
        return;
    }
    
    filtered.forEach(f => {
        const card = document.createElement('div');
        card.className = `finding-card level-${f.level}`;
        
        let typeLbl = f.level.toUpperCase();
        let classLbl = `lbl-${f.level}`;
        
        card.innerHTML = `
            <div class="finding-meta">
                <span class="finding-lbl ${classLbl}">${typeLbl}</span>
                <span class="finding-elem">${f.element_type ? f.element_type.toUpperCase() : 'DOCUMENT'} #${f.index !== null && f.index !== undefined ? f.index : ''} (${f.category})</span>
            </div>
            <div class="finding-msg">${escapeHtml(f.message)}</div>
            ${f.text ? `<div class="finding-context">Context: "${escapeHtml(f.text)}"</div>` : ''}
        `;
        
        // Click navigation to element in Interactive Preview Tab
        card.addEventListener('click', () => {
            // Switch tab to preview
            elements.tabButtons[0].click();
            
            // Find target element
            const prefix = f.element_type === 'paragraph' ? 'doc-p-' : 'doc-t-';
            const element = document.getElementById(`${prefix}${f.index}`);
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                // Add a brief glow highlight
                element.style.transition = 'all 0.3s ease';
                element.style.boxShadow = '0 0 15px rgba(99, 102, 241, 0.6)';
                setTimeout(() => {
                    element.style.boxShadow = 'none';
                }, 2000);
            }
        });
        
        elements.findingsList.appendChild(card);
    });
}
