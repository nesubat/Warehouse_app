/* =========================================
   1. GLOBAL UTILITIES (Runs on all pages)
========================================= */
document.addEventListener("DOMContentLoaded", function() {
    // Smooth scroll for any anchor links (e.g., href="#dashboard")
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            const targetId = this.getAttribute('href');
            if (targetId !== '#') {
                e.preventDefault();
                const targetElement = document.querySelector(targetId);
                if (targetElement) {
                    targetElement.scrollIntoView({ behavior: 'smooth' });
                }
            }
        });
    });
});


/* =========================================
   2. DASHBOARD LOGIC (index.html)
========================================= */
// Scroll Spy for Top Nav
window.addEventListener('scroll', function() {
    const dashSection = document.getElementById('dashboard-section');
    const navHome = document.getElementById('nav-home');
    const navDash = document.getElementById('nav-dash');

    if (dashSection && navHome && navDash) {
        if (window.scrollY >= dashSection.offsetTop - 150) {
            navHome.classList.remove('active');
            navDash.classList.add('active');
        } else {
            navDash.classList.remove('active');
            navHome.classList.add('active');
        }
    }
});

// Project Accordion Toggle
function toggleProject(projectName) {
    const content = document.getElementById('content-' + projectName);
    const icon = document.getElementById('icon-' + projectName);
    
    if (content.style.display === 'none') {
        content.style.display = 'block';
        icon.innerHTML = '▲';
        icon.style.color = '#3498db';
    } else {
        content.style.display = 'none';
        icon.innerHTML = '▼';
        icon.style.color = '#7f8c8d';
    }
}

// Preserve Scroll & Accordion State After Deletions
document.addEventListener("DOMContentLoaded", function() {
    const savedScroll = sessionStorage.getItem('dashboardScroll');
    if (savedScroll) {
        window.scrollTo(0, parseInt(savedScroll));
        sessionStorage.removeItem('dashboardScroll');
    }

    const openProject = sessionStorage.getItem('openProject');
    if (openProject) {
        const content = document.getElementById('content-' + openProject);
        const icon = document.getElementById('icon-' + openProject);
        if (content && icon) {
            content.style.display = 'block';
            icon.innerHTML = '▲';
            icon.style.color = '#3498db';
        }
        sessionStorage.removeItem('openProject');
    }
    
    const forms = document.querySelectorAll('form');
    forms.forEach(form => {
        if (form.action.includes('/delete')) {
            form.addEventListener('submit', function() {
                sessionStorage.setItem('dashboardScroll', window.scrollY);
                if (form.action.includes('/delete_file/')) {
                    const parentCard = form.closest('.project-card');
                    if (parentCard) {
                        const contentDiv = parentCard.querySelector('.accordion-content');
                        if (contentDiv && contentDiv.style.display === 'block') {
                            const projName = contentDiv.id.replace('content-', '');
                            sessionStorage.setItem('openProject', projName);
                        }
                    }
                }
            });
        }
    });
});


/* =========================================
   3. MATRIX ENGINE LOGIC (matrix.html)
========================================= */
function toggleAllPacks(masterCheckbox) {
    const parentList = masterCheckbox.closest('.pack-list');
    if (parentList) {
        const packCheckboxes = parentList.querySelectorAll('.pack-checkbox');
        packCheckboxes.forEach(function(checkbox) {
            checkbox.checked = masterCheckbox.checked;
        });
    }
}


/* =========================================
   4. SUB-GROUP LOGIC (subgroup.html)
========================================= */
let meta = null;

// Parse the JSON metadata safely on page load
document.addEventListener('DOMContentLoaded', () => {
    const metaTag = document.getElementById('meta-data');
    if (metaTag) {
        try {
            meta = JSON.parse(metaTag.textContent);
        } catch (e) {
            console.error("Could not parse metadata JSON");
        }
    }
});

function toggleTabBlock(checkbox) {
    if (!meta) return;
    const tabName = checkbox.value;
    const safeTabId = tabName.replace(/[^a-zA-Z0-9]/g, '_');
    const container = document.getElementById('tab-blocks-container');
    const submitBtn = document.getElementById('submit-btn');

    if (checkbox.checked) {
        let packOptions = '<div class="tab-checkbox-grid" style="margin-top: 5px;">';
        const packs = meta.tabs[tabName].packs;
        for (const [packName, packData] of Object.entries(packs)) {
            if (packData.is_selected) { 
                packOptions += `
                    <label class="tab-checkbox-label">
                        <input type="checkbox" name="target_pack_${tabName}[]" value="${packName}" onchange="togglePackBlock(this, '${tabName}', '${packName}')">
                        ${packName}
                    </label>
                `;
            }
        }
        packOptions += '</div>';

        const blockHTML = `
            <div class="tab-card" id="block_${safeTabId}" style="border-left: 5px solid #3498db;">
                <h3 class="tab-title">⚙️ Setup for Tab: <span style="color: #3498db;">${tabName}</span></h3>
                <input type="hidden" name="selected_tabs[]" value="${tabName}">

                <div class="input-field" style="width: 50%; min-width: 200px; margin-bottom: 20px;">
                    <label>Item Number Row (e.g. 9)</label>
                    <input type="number" name="item_row_${tabName}" style="width: 100%; padding: 10px; border: 1px solid #bdc3c7; border-radius: 5px; box-sizing: border-box;" required>
                </div>

                <div class="input-field" style="margin-bottom: 20px;">
                    <label>Parent Pack(s) to Sub-Group</label>
                    ${packOptions}
                </div>
                
                <div id="pack_blocks_container_${safeTabId}"></div>
            </div>
        `;
        container.insertAdjacentHTML('beforeend', blockHTML);
    } else {
        const block = document.getElementById(`block_${safeTabId}`);
        if (block) block.remove();
    }
    checkSubmitButton();
}

function togglePackBlock(checkbox, tabName, packName) {
    const safeTabId = tabName.replace(/[^a-zA-Z0-9]/g, '_');
    const safePackId = packName.replace(/[^a-zA-Z0-9]/g, '_');
    const container = document.getElementById(`pack_blocks_container_${safeTabId}`);

    if (checkbox.checked) {
        const blockHTML = `
            <div class="blueprint-card" id="block_${safeTabId}_${safePackId}" style="margin-top: 15px;">
                <h4 style="color: #2c3e50; font-size: 15px; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-bottom: 15px;">
                    Define Sub-Groups for: <span style="color: #e67e22;">${packName}</span>
                </h4>
                
                <div id="sg_container_${safeTabId}_${safePackId}">
                    <div class="input-group subgroup-row" style="margin-bottom: 15px; align-items: flex-end;">
                        <div class="input-field">
                            <label>Start Item #</label>
                            <input type="number" name="start_item_${tabName}_${packName}[]" required>
                        </div>
                        <div class="input-field">
                            <label>End Item #</label>
                            <input type="number" name="end_item_${tabName}_${packName}[]" required>
                        </div>
                    </div>
                </div>

                <button type="button" class="btn-file open" onclick="addSubGroupRow('${tabName}', '${packName}', '${safeTabId}', '${safePackId}')" style="margin-top: 5px;">
                    + Add Another Sub-Group for ${packName}
                </button>
            </div>
        `;
        container.insertAdjacentHTML('beforeend', blockHTML);
    } else {
        const block = document.getElementById(`block_${safeTabId}_${safePackId}`);
        if (block) block.remove();
    }
}

function addSubGroupRow(tabName, packName, safeTabId, safePackId) {
    const container = document.getElementById(`sg_container_${safeTabId}_${safePackId}`);
    const rowHTML = `
        <div class="input-group subgroup-row" style="margin-bottom: 15px; align-items: flex-end;">
            <div class="input-field">
                <label>Start Item #</label>
                <input type="number" name="start_item_${tabName}_${packName}[]" required>
            </div>
            <div class="input-field">
                <label>End Item #</label>
                <input type="number" name="end_item_${tabName}_${packName}[]" required>
            </div>
            <button type="button" class="btn-file delete" onclick="this.parentElement.remove()" style="height: 40px; min-width: 80px;">Remove</button>
        </div>
    `;
    container.insertAdjacentHTML('beforeend', rowHTML);
}

function checkSubmitButton() {
    const submitBtn = document.getElementById('submit-btn');
    if (submitBtn) {
        const checkedTabs = document.querySelectorAll('#tab-blocks-container > .tab-card');
        submitBtn.style.display = checkedTabs.length > 0 ? 'block' : 'none';
    }
}


/* =========================================
   5. UPLOAD FORM LOGIC (Step 1)
========================================= */
document.addEventListener("DOMContentLoaded", function() {
    // Only execute if we are actually on the Upload page
    const submitBtn = document.getElementById('submit_btn');
    if (!submitBtn) return;

    const dataStoreElement = document.getElementById('project-data-store');
    const projectDataRaw = dataStoreElement ? dataStoreElement.textContent : '{}';
    let projectData = {};
    try {
        projectData = JSON.parse(projectDataRaw);
    } catch (e) {
        console.error("Could not parse project data JSON");
    }

    function validateForm() {
        const dropdown = document.getElementById('existing_project_select');
        const textInput = document.getElementById('new_project_input');
        const fileInput = document.getElementById('excel_file_input');
        const helperText = document.getElementById('file-helper-text');

        let isValid = false;

        // SCENARIO A: Creating a New Project
        if (textInput && textInput.value.trim() !== "") {
            if (dropdown) {
                dropdown.value = "";
                dropdown.disabled = true;
                dropdown.style.backgroundColor = "#e9ecef";
            }
            if (helperText) helperText.innerHTML = "Required: Please upload the signature links Excel file for your new project.";
            
            if (fileInput && fileInput.files.length > 0) {
                isValid = true;
            }
        } 
        // SCENARIO B: Selecting an Existing Project
        else if (dropdown && dropdown.value !== "") {
            if (textInput) {
                textInput.value = "";
                textInput.disabled = true;
                textInput.style.backgroundColor = "#e9ecef";
            }
            
            const selectedProj = dropdown.value;
            const existingFile = projectData[selectedProj]; 
            
            if (existingFile) {
                if (helperText) helperText.innerHTML = `📂 <strong>Found:</strong> Upload another file otherwise system will use '<strong>${existingFile}</strong>'.`;
                isValid = true; 
            } else {
                if (helperText) helperText.innerHTML = "⚠️ No signature file found in this project. You MUST upload one below.";
                if (fileInput && fileInput.files.length > 0) {
                    isValid = true; 
                }
            }
        } 
        // SCENARIO C: Completely Empty Form
        else {
            if (dropdown) {
                dropdown.disabled = false;
                dropdown.style.backgroundColor = "#fff";
            }
            if (textInput) {
                textInput.disabled = false;
                textInput.style.backgroundColor = "#fff";
            }
            if (helperText) helperText.innerHTML = "Required if creating a new project. If using an existing project, upload a fresh Excel file or skip if already present.";
            
            isValid = false; 
        }

        // Lock or unlock the button
        if (isValid) {
            submitBtn.disabled = false;
            submitBtn.style.opacity = '1';
            submitBtn.style.cursor = 'pointer';
        } else {
            submitBtn.disabled = true;
            submitBtn.style.opacity = '0.5';
            submitBtn.style.cursor = 'not-allowed';
        }
    }

    const dropdown = document.getElementById('existing_project_select');
    const textInput = document.getElementById('new_project_input');
    const fileInput = document.getElementById('excel_file_input');
    
    if (dropdown) dropdown.addEventListener('change', validateForm);
    if (textInput) textInput.addEventListener('input', validateForm);
    if (fileInput) fileInput.addEventListener('change', validateForm);
    
    validateForm();
});


/* =========================================
   6. PDF LABEL SHUFFLER LOGIC (pdf.html)
========================================= */
document.addEventListener("DOMContentLoaded", function() {
    const fileInputs = document.querySelectorAll(".pdf-file-input");
    const masterCheckbox = document.getElementById("master-divider-checkbox");
    const packCheckboxes = document.querySelectorAll(".pack-divider-checkbox");

    // Only run if elements exist on the page
    if (fileInputs.length > 0 && masterCheckbox) {
        
        // 1. Listen for File Uploads to Show/Hide individual checkboxes
        fileInputs.forEach(input => {
            input.addEventListener("change", function() {
                const packName = this.getAttribute("data-pack");
                const container = document.getElementById("divider-container-" + packName);
                if (!container) return;
                
                const checkbox = container.querySelector(".pack-divider-checkbox");
                
                if (this.files && this.files.length > 0) {
                    container.style.display = "block"; // Reveal it
                } else {
                    container.style.display = "none";  // Hide it
                    if (checkbox) checkbox.checked = false; // Uncheck it if hidden
                }
            });
        });

        // 2. Master Checkbox Logic (Only affects VISIBLE checkboxes)
        masterCheckbox.addEventListener("change", function() {
            const isChecked = this.checked;
            packCheckboxes.forEach(checkbox => {
                // Only tick it if the container is currently visible (meaning a file is attached)
                const container = checkbox.closest("div");
                if (container && container.style.display === "block") {
                    checkbox.checked = isChecked;
                }
            });
        });
    }
});