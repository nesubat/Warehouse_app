/* =========================================
   WAREHOUSE AUTOMATION SUITE - MASTER SCRIPT
   Architecture: Event Delegation & State Management
========================================= */

document.addEventListener("DOMContentLoaded", function() {

    // =========================================
    // 1. GLOBAL UTILITIES (Smooth Scrolling)
    // =========================================
    document.body.addEventListener('click', function(e) {
        const anchor = e.target.closest('a[href^="#"]');
        if (anchor && anchor.getAttribute('href') !== '#') {
            e.preventDefault();
            const targetElement = document.querySelector(anchor.getAttribute('href'));
            if (targetElement) {
                targetElement.scrollIntoView({ behavior: 'smooth' });
            }
        }
    });

    // =========================================
    // 2. DASHBOARD LOGIC (index.html)
    // =========================================
    // A. Scroll Spy for Top Nav
    const dashSection = document.getElementById('dashboard-section');
    const navHome = document.getElementById('nav-home');
    const navDash = document.getElementById('nav-dash');
    
    if (dashSection && navHome && navDash) {
        window.addEventListener('scroll', function() {
            if (window.scrollY >= dashSection.offsetTop - 150) {
                navHome.classList.remove('active');
                navDash.classList.add('active');
            } else {
                navDash.classList.remove('active');
                navHome.classList.add('active');
            }
        });
    }

    // B. Accordion Toggle via Event Delegation
    document.body.addEventListener('click', function(e) {
        const header = e.target.closest('.accordion-header');
        if (header) {
            const content = header.nextElementSibling;
            const icon = header.querySelector('.accordion-icon');
            
            if (content && content.classList.contains('accordion-content')) {
                if (content.style.display === 'none' || content.style.display === '') {
                    content.style.display = 'block';
                    if (icon) {
                        icon.innerHTML = '▲';
                        icon.style.color = '#3498db';
                    }
                } else {
                    content.style.display = 'none';
                    if (icon) {
                        icon.innerHTML = '▼';
                        icon.style.color = '#7f8c8d';
                    }
                }
            }
        }
    });

    // C. Preserve State After Deletions
    const savedScroll = sessionStorage.getItem('dashboardScroll');
    if (savedScroll) {
        window.scrollTo(0, parseInt(savedScroll));
        sessionStorage.removeItem('dashboardScroll');
    }
    const openProject = sessionStorage.getItem('openProject');
    if (openProject) {
        const content = document.getElementById('content-' + openProject);
        if (content) {
            content.style.display = 'block';
            const icon = content.previousElementSibling.querySelector('.accordion-icon');
            if (icon) {
                icon.innerHTML = '▲';
                icon.style.color = '#3498db';
            }
        }
        sessionStorage.removeItem('openProject');
    }

    // Capture state right before form submission
    document.body.addEventListener('submit', function(e) {
        if (e.target.action && e.target.action.includes('/delete')) {
            sessionStorage.setItem('dashboardScroll', window.scrollY);
            if (e.target.action.includes('/delete_file/')) {
                const parentCard = e.target.closest('.project-card');
                if (parentCard) {
                    const contentDiv = parentCard.querySelector('.accordion-content');
                    if (contentDiv && contentDiv.style.display === 'block') {
                        const projName = contentDiv.id.replace('content-', '');
                        sessionStorage.setItem('openProject', projName);
                    }
                }
            }
        }
    });

    // D. Safely Handle "Open Locally" Fetch Requests
    document.body.addEventListener('click', function(e) {
        const openBtn = e.target.closest('a.btn-file.open');
        if (openBtn && openBtn.hasAttribute('href')) {
            e.preventDefault();
            fetch(openBtn.href).then(response => {
                if (!response.ok) alert("Could not open the file. It may have been moved or deleted.");
            }).catch(err => alert("Network error trying to open the file."));
        }
    });


    // =========================================
    // 3. SUB-GROUP ENGINE (Dynamic UI)
    // =========================================
    let meta = null;
    const metaTag = document.getElementById('meta-data');
    if (metaTag) {
        try {
            meta = JSON.parse(metaTag.textContent);
        } catch (e) {
            console.error("Could not parse metadata JSON");
        }
    }
    const selectedTabsState = new Set(); // State tracker
    // Listen for Checkbox Toggles
    document.body.addEventListener('change', function(e) {
        
        // --- RESTORED: Matrix "Select All Packs" Logic ---
        if (e.target.matches('input[id^="selectAllPacks_"]')) {
            const parentList = e.target.closest('.pack-list');
            if (parentList) {
                const packCheckboxes = parentList.querySelectorAll('.pack-checkbox');
                packCheckboxes.forEach(function(checkbox) {
                    checkbox.checked = e.target.checked;
                });
            }
        }
        // -------------------------------------------------

        // Tab Selector (Existing code...)
        if (e.target.matches('.tab-checkbox-grid input[type="checkbox"]:not([name^="target_pack"])')) {
            const tabName = e.target.value;
            const safeTabId = tabName.replace(/[^a-zA-Z0-9]/g, '_');
            const container = document.getElementById('tab-blocks-container');
            const submitBtn = document.getElementById('submit-btn');

            if (e.target.checked && meta) {
                selectedTabsState.add(tabName);
                
                let packOptions = '<div class="tab-checkbox-grid" style="margin-top: 5px;">';
                const packs = meta.tabs[tabName].packs;
                for (const [packName, packData] of Object.entries(packs)) {
                    if (packData.is_selected) { 
                        packOptions += `
                            <label class="tab-checkbox-label">
                                <!-- Notice: using data- attributes instead of onchange -->
                                <input type="checkbox" name="target_pack_${tabName}[]" value="${packName}" data-tab="${tabName}" data-pack="${packName}">
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
                selectedTabsState.delete(tabName);
                const block = document.getElementById(`block_${safeTabId}`);
                if (block) block.remove();
            }
            if (submitBtn) submitBtn.style.display = selectedTabsState.size > 0 ? 'block' : 'none';
        }
        
        // Pack Selector
        if (e.target.matches('input[name^="target_pack_"]')) {
            const tabName = e.target.dataset.tab;
            const packName = e.target.dataset.pack;
            const safeTabId = tabName.replace(/[^a-zA-Z0-9]/g, '_');
            const safePackId = packName.replace(/[^a-zA-Z0-9]/g, '_');
            const container = document.getElementById(`pack_blocks_container_${safeTabId}`);

            if (e.target.checked) {
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
                        <!-- Notice: using data- attributes instead of onclick -->
                        <button type="button" class="btn-file open add-subgroup-btn" data-tab="${tabName}" data-pack="${packName}" data-safetab="${safeTabId}" data-safepack="${safePackId}" style="margin-top: 5px;">
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
    });

    // Listen for Dynamic Button Clicks (Add/Remove Rows)
    document.body.addEventListener('click', function(e) {
        if (e.target.matches('.add-subgroup-btn')) {
            const tabName = e.target.dataset.tab;
            const packName = e.target.dataset.pack;
            const container = document.getElementById(`sg_container_${e.target.dataset.safetab}_${e.target.dataset.safepack}`);
            
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
                    <button type="button" class="btn-file delete remove-subgroup-btn" style="height: 40px; min-width: 80px;">Remove</button>
                </div>
            `;
            container.insertAdjacentHTML('beforeend', rowHTML);
        }
        
        if (e.target.matches('.remove-subgroup-btn')) {
            const row = e.target.closest('.subgroup-row');
            if (row) row.remove();
        }
    });


    // =========================================
    // 4. UPLOAD FORM (Step 1)
    // =========================================
    const submitBtn = document.getElementById('submit_btn');
    if (submitBtn) {
        const dataStoreElement = document.getElementById('project-data-store');
        const projectData = dataStoreElement ? JSON.parse(dataStoreElement.textContent || '{}') : {};
        const dropdown = document.getElementById('existing_project_select');
        const textInput = document.getElementById('new_project_input');
        const fileInput = document.getElementById('excel_file_input');
        const helperText = document.getElementById('file-helper-text');

        function validateForm() {
            let isValid = false;
            
            if (textInput && textInput.value.trim() !== "") {
                if (dropdown) {
                    dropdown.value = "";
                    dropdown.disabled = true;
                    dropdown.style.backgroundColor = "#e9ecef";
                }
                if (helperText) helperText.innerHTML = "Required: Please upload the signature links Excel file for your new project.";
                if (fileInput && fileInput.files.length > 0) isValid = true;
            } 
            else if (dropdown && dropdown.value !== "") {
                if (textInput) {
                    textInput.value = "";
                    textInput.disabled = true;
                    textInput.style.backgroundColor = "#e9ecef";
                }
                if (projectData[dropdown.value]) {
                    if (helperText) helperText.innerHTML = `📂 <strong>Found:</strong> Upload another file otherwise system will use '<strong>${projectData[dropdown.value]}</strong>'.`;
                    isValid = true; 
                } else {
                    if (helperText) helperText.innerHTML = "⚠️ No signature file found in this project. You MUST upload one below.";
                    if (fileInput && fileInput.files.length > 0) isValid = true; 
                }
            } 
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

        if (dropdown) dropdown.addEventListener('change', validateForm);
        if (textInput) textInput.addEventListener('input', validateForm);
        if (fileInput) fileInput.addEventListener('change', validateForm);
        validateForm();
    }


    // =========================================
    // 5. PDF LABEL SHUFFLER (Step 2)
    // =========================================
    const masterCheckbox = document.getElementById("master-divider-checkbox");
    if (masterCheckbox) {
        const fileInputs = document.querySelectorAll(".pdf-file-input");
        const packCheckboxes = document.querySelectorAll(".pack-divider-checkbox");

        fileInputs.forEach(input => {
            input.addEventListener("change", function() {
                const container = document.getElementById("divider-container-" + this.getAttribute("data-pack"));
                if (container) {
                    const checkbox = container.querySelector(".pack-divider-checkbox");
                    if (this.files && this.files.length > 0) {
                        container.style.display = "block";
                    } else {
                        container.style.display = "none";
                        if (checkbox) checkbox.checked = false;
                    }
                }
            });
        });

        masterCheckbox.addEventListener("change", function() {
            const isChecked = this.checked;
            packCheckboxes.forEach(checkbox => {
                const container = checkbox.closest("div");
                if (container && container.style.display === "block") {
                    checkbox.checked = isChecked;
                }
            });
        });
    }// =========================================
    // 6. MATRIX ANCHOR POINTS MEMORY
    // =========================================
    // Restores user's previous input for Start Cell, Job ID, and Store Col
    const previewForm = document.querySelector('form[action="/preview"]');
    if (previewForm) {
        // A. On page load, override default values with saved memory
        previewForm.querySelectorAll('input[name^="start_"]').forEach(el => {
            const saved = localStorage.getItem('anchor_start');
            if (saved && el.value === 'B8') el.value = saved;
        });
        previewForm.querySelectorAll('input[name^="job_"]').forEach(el => {
            const saved = localStorage.getItem('anchor_job');
            if (saved && el.value === 'E1') el.value = saved;
        });
        previewForm.querySelectorAll('input[name^="store_"]').forEach(el => {
            const saved = localStorage.getItem('anchor_store');
            if (saved && el.value === 'A') el.value = saved;
        });

        // B. On submit, save the current inputs to memory for next time
        previewForm.addEventListener('submit', function() {
            const firstStart = previewForm.querySelector('input[name^="start_"]');
            const firstJob = previewForm.querySelector('input[name^="job_"]');
            const firstStore = previewForm.querySelector('input[name^="store_"]');
            
            if (firstStart) localStorage.setItem('anchor_start', firstStart.value);
            if (firstJob) localStorage.setItem('anchor_job', firstJob.value);
            if (firstStore) localStorage.setItem('anchor_store', firstStore.value);
        });
    }

    // =========================================
    // 7. AUTO-SCROLL TO PREVIEWS
    // =========================================
    // If the preview section exists on load, smoothly scroll to it
    const previewSection = document.getElementById('preview-section');
    if (previewSection) {
        setTimeout(() => {
            previewSection.scrollIntoView({ behavior: 'smooth' });
        }, 150); // Slight delay ensures page is fully rendered before scrolling
    }

    // =========================================
    // 8. LOADING OVERLAYS (Global)
    // =========================================
    document.body.addEventListener('submit', function(e) {
        // Exclude overlay for Deletions AND Preview Generation
        if (e.target.action && (e.target.action.includes('/delete') || e.target.action.includes('/preview'))) {
            return; 
        }

        const overlay = document.getElementById('loading-overlay');
        if (overlay) {
            overlay.style.display = 'flex';
        }
    });

}); 