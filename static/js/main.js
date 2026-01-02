/**
 * CO-PO Attainment Generator - Frontend JavaScript
 * Handles form interactions, file uploads, and API calls
 */

// API Base URL
const API_BASE = '';

// State
let state = {
    regulation: '',
    category: '',
    deptType: 'default',
    files: {},
    requiredInputs: [],
    isLoading: false
};

// DOM Elements
const elements = {
    regulationSelect: document.getElementById('regulation'),
    categorySelect: document.getElementById('category'),
    deptTypeSelect: document.getElementById('dept_type'),
    deptTypeGroup: document.getElementById('deptTypeGroup'),
    filesSection: document.getElementById('filesSection'),
    fileInputsContainer: document.getElementById('fileInputs'),
    submitSection: document.getElementById('submitSection'),
    mainForm: document.getElementById('mainForm'),
    resetBtn: document.getElementById('resetBtn'),
    submitBtn: document.getElementById('submitBtn'),
    progressSteps: document.querySelectorAll('.progress-step'),
    progressBar: document.getElementById('progressBar'),
    progressCount: document.getElementById('progressCount'),
    alertContainer: document.getElementById('alertContainer')
};

// Initialize
document.addEventListener('DOMContentLoaded', init);

function init() {
    setupEventListeners();
    updateProgress();
}

// Event Listeners
function setupEventListeners() {
    elements.regulationSelect.addEventListener('change', handleRegulationChange);
    elements.categorySelect.addEventListener('change', handleCategoryChange);
    elements.deptTypeSelect.addEventListener('change', handleDeptTypeChange);
    elements.mainForm.addEventListener('submit', handleSubmit);
    elements.resetBtn.addEventListener('click', handleReset);
}

// Handlers
async function handleRegulationChange(e) {
    state.regulation = e.target.value;
    resetSubsequentFields('regulation');
    
    if (!state.regulation) return;
    
    try {
        const response = await fetch(`${API_BASE}/api/categories/${state.regulation}`);
        const data = await response.json();
        
        populateSelect(elements.categorySelect, data.categories, 'Select Category');
        elements.categorySelect.disabled = false;
        updateProgress();
    } catch (error) {
        showAlert('Error fetching categories', 'error');
        console.error(error);
    }
}

async function handleCategoryChange(e) {
    state.category = e.target.value;
    resetSubsequentFields('category');
    
    if (!state.category) return;
    
    try {
        // Fetch department types
        const deptResponse = await fetch(`${API_BASE}/api/dept_types/${state.regulation}/${state.category}`);
        const deptData = await deptResponse.json();
        
        if (deptData.dept_types.length > 0 && deptData.dept_types[0] !== 'default') {
            populateSelect(elements.deptTypeSelect, deptData.dept_types, 'Select Type', formatDeptType);
            elements.deptTypeGroup.classList.remove('hidden');
        } else {
            state.deptType = 'default';
            elements.deptTypeGroup.classList.add('hidden');
        }
        
        // Fetch required inputs
        await loadRequiredInputs();
        updateProgress();
    } catch (error) {
        showAlert('Error fetching options', 'error');
        console.error(error);
    }
}

function handleDeptTypeChange(e) {
    state.deptType = e.target.value;
    updateProgress();
}

async function loadRequiredInputs() {
    try {
        const response = await fetch(`${API_BASE}/api/required_inputs/${state.regulation}/${state.category}`);
        const data = await response.json();
        
        if (data.error) {
            showAlert(data.error, 'error');
            return;
        }
        
        state.requiredInputs = data.inputs;
        state.files = {};
        
        renderFileInputs(data.inputs);
        elements.filesSection.classList.remove('hidden');
    } catch (error) {
        showAlert('Error fetching required inputs', 'error');
        console.error(error);
    }
}

function renderFileInputs(inputs) {
    elements.fileInputsContainer.innerHTML = inputs.map((input, index) => `
        <div class="form-group" id="fileGroup_${input.toLowerCase()}">
            <label>
                <span class="step-badge">${index + 1}</span>
                ${input} Evaluation Sheet
                <span class="required">*</span>
            </label>
            <div class="upload-zone" 
                 id="dropzone_${input.toLowerCase()}"
                 ondragover="handleDragOver(event, '${input.toLowerCase()}')"
                 ondragleave="handleDragLeave(event, '${input.toLowerCase()}')"
                 ondrop="handleDrop(event, '${input.toLowerCase()}')"
                 onclick="document.getElementById('file_${input.toLowerCase()}').click()">
                <div class="upload-zone-icon">📁</div>
                <h4>Drag & drop ${input} file here</h4>
                <p>or <span class="browse-link">browse</span> to upload</p>
                <p style="margin-top: 0.5rem; font-size: 0.75rem; color: #94a3b8;">
                    Supports: .xlsx, .xls
                </p>
            </div>
            <input type="file" 
                   id="file_${input.toLowerCase()}" 
                   name="file_${input.toLowerCase()}"
                   accept=".xlsx,.xls"
                   class="hidden"
                   onchange="handleFileSelect(event, '${input.toLowerCase()}', '${input}')"
                   required>
            <div class="file-list" id="fileList_${input.toLowerCase()}"></div>
        </div>
    `).join('');
}

// File Handling
function handleDragOver(e, inputType) {
    e.preventDefault();
    document.getElementById(`dropzone_${inputType}`).classList.add('dragover');
}

function handleDragLeave(e, inputType) {
    e.preventDefault();
    document.getElementById(`dropzone_${inputType}`).classList.remove('dragover');
}

function handleDrop(e, inputType) {
    e.preventDefault();
    document.getElementById(`dropzone_${inputType}`).classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        if (isValidFile(file)) {
            setFile(inputType, file);
        } else {
            showAlert('Please upload a valid Excel file (.xlsx or .xls)', 'error');
        }
    }
}

function handleFileSelect(e, inputType, displayName) {
    const file = e.target.files[0];
    if (file && isValidFile(file)) {
        setFile(inputType, file);
    }
}

function isValidFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const validExtensions = ['.xlsx', '.xls'];
    const extension = '.' + file.name.split('.').pop().toLowerCase();
    
    return validTypes.includes(file.type) || validExtensions.includes(extension);
}

function setFile(inputType, file) {
    state.files[inputType] = file;
    
    // Update UI
    const fileList = document.getElementById(`fileList_${inputType}`);
    fileList.innerHTML = `
        <div class="file-item">
            <div class="file-item-info">
                <span class="file-item-icon">✅</span>
                <div>
                    <div class="file-item-name">${file.name}</div>
                    <div class="file-item-size">${formatFileSize(file.size)}</div>
                </div>
            </div>
            <button type="button" class="file-item-remove" onclick="removeFile('${inputType}')">
                ✕
            </button>
        </div>
    `;
    
    // Hide dropzone
    document.getElementById(`dropzone_${inputType}`).style.display = 'none';
    
    updateProgress();
    checkAllFilesUploaded();
}

function removeFile(inputType) {
    delete state.files[inputType];
    
    // Reset file input
    const fileInput = document.getElementById(`file_${inputType}`);
    fileInput.value = '';
    
    // Update UI
    document.getElementById(`fileList_${inputType}`).innerHTML = '';
    document.getElementById(`dropzone_${inputType}`).style.display = 'block';
    
    elements.submitSection.classList.add('hidden');
    updateProgress();
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function checkAllFilesUploaded() {
    const allUploaded = state.requiredInputs.every(input => 
        state.files[input.toLowerCase()]
    );
    
    if (allUploaded) {
        elements.submitSection.classList.remove('hidden');
        elements.submitBtn.disabled = false;
    } else {
        elements.submitSection.classList.add('hidden');
    }
}

// Form Submission
async function handleSubmit(e) {
    e.preventDefault();
    
    if (state.isLoading) return;
    
    state.isLoading = true;
    elements.submitBtn.disabled = true;
    elements.submitBtn.innerHTML = '<span class="spinner"></span> Generating...';
    
    try {
        const formData = new FormData();
        formData.append('regulation', state.regulation);
        formData.append('category', state.category);
        formData.append('dept_type', state.deptType);
        
        // Append files
        for (const [inputType, file] of Object.entries(state.files)) {
            formData.append(`file_${inputType}`, file);
        }
        
        const response = await fetch(`${API_BASE}/generate`, {
            method: 'POST',
            body: formData
        });
        
        // Check if we got redirected to result page or back to index
        if (response.redirected) {
            window.location.href = response.url;
        } else {
            const html = await response.text();
            document.open();
            document.write(html);
            document.close();
        }
    } catch (error) {
        showAlert('Error generating attainment sheet: ' + error.message, 'error');
        console.error(error);
        resetSubmitButton();
    }
}

function resetSubmitButton() {
    state.isLoading = false;
    elements.submitBtn.disabled = false;
    elements.submitBtn.innerHTML = '📥 Generate Attainment Sheet';
}

// Reset Form
function handleReset() {
    if (confirm('Are you sure you want to reset the form? All selections and uploaded files will be cleared.')) {
        state = {
            regulation: '',
            category: '',
            deptType: 'default',
            files: {},
            requiredInputs: [],
            isLoading: false
        };
        
        elements.mainForm.reset();
        elements.regulationSelect.value = '';
        elements.categorySelect.innerHTML = '<option value="">-- Select Category --</option>';
        elements.categorySelect.disabled = true;
        elements.deptTypeGroup.classList.add('hidden');
        elements.filesSection.classList.add('hidden');
        elements.submitSection.classList.add('hidden');
        elements.fileInputsContainer.innerHTML = '';
        
        hideAlert();
        updateProgress();
    }
}

// Progress
function updateProgress() {
    let completed = 0;
    const total = 3;
    
    if (state.regulation) completed++;
    if (state.category) completed++;
    if (Object.keys(state.files).length === state.requiredInputs.length && state.requiredInputs.length > 0) completed++;
    
    // Update steps
    elements.progressSteps.forEach((step, index) => {
        step.classList.remove('pending', 'active', 'completed');
        if (index < completed) {
            step.classList.add('completed');
            step.innerHTML = '✓';
        } else if (index === completed) {
            step.classList.add('active');
            step.innerHTML = index + 1;
        } else {
            step.classList.add('pending');
            step.innerHTML = index + 1;
        }
    });
    
    // Update bar
    const percentage = (completed / total) * 100;
    elements.progressBar.style.width = `${percentage}%`;
    elements.progressCount.textContent = `${completed} / ${total} steps completed`;
}

// Utility Functions
function populateSelect(select, options, placeholder, formatter = null) {
    select.innerHTML = `<option value="">-- ${placeholder} --</option>`;
    options.forEach(option => {
        const opt = document.createElement('option');
        opt.value = option;
        opt.textContent = formatter ? formatter(option) : capitalizeFirst(option);
        select.appendChild(opt);
    });
}

function capitalizeFirst(str) {
    return str.charAt(0).toUpperCase() + str.slice(1);
}

function formatDeptType(type) {
    const map = {
        'dept': 'Department',
        's&h': 'Science & Humanities (S&H)',
        'default': 'Default'
    };
    return map[type] || capitalizeFirst(type);
}

function resetSubsequentFields(from) {
    if (from === 'regulation') {
        elements.categorySelect.innerHTML = '<option value="">-- Select Category --</option>';
        elements.categorySelect.disabled = true;
        state.category = '';
    }
    
    elements.deptTypeGroup.classList.add('hidden');
    elements.filesSection.classList.add('hidden');
    elements.submitSection.classList.add('hidden');
    state.deptType = 'default';
    state.files = {};
    state.requiredInputs = [];
}

// Alerts
function showAlert(message, type = 'error') {
    const icons = {
        success: '✅',
        error: '❌',
        warning: '⚠️'
    };
    
    elements.alertContainer.innerHTML = `
        <div class="alert alert-${type}">
            <span class="alert-icon">${icons[type]}</span>
            <span>${message}</span>
        </div>
    `;
    elements.alertContainer.classList.remove('hidden');
    
    // Auto-hide after 5 seconds
    setTimeout(hideAlert, 5000);
}

function hideAlert() {
    elements.alertContainer.classList.add('hidden');
    elements.alertContainer.innerHTML = '';
}
