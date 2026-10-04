/* =========================================
   WAREHOUSE AUTOMATION SUITE - MASTER SCRIPT
   Architecture: Event Delegation & State Management
========================================= */

document.addEventListener("DOMContentLoaded", function() {

    // =========================================
    // 0. DISCARD UNGENERATED UPLOADS ON LEAVE
    // =========================================
    // A page with #discard-on-leave tells the server to delete its upload when the user leaves
    // without generating. Submitting the page's own forms (Preview, Generate, Recheck) doesn't
    // count as leaving, and Back/Forward keeps the page cached, so nothing is deleted then either.
    const discardInfo = document.getElementById('discard-on-leave');
    if (discardInfo) {
        let submitting = false;
        document.addEventListener('submit', () => { submitting = true; });
        window.addEventListener('pagehide', (e) => {
            if (submitting || e.persisted) return;
            const data = new URLSearchParams();
            data.append(discardInfo.dataset.field, discardInfo.dataset.value);
            navigator.sendBeacon(discardInfo.dataset.url, data);
        });
    }

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

        // --- Matrix "Sent to an installer" (per tab) ---
        // Ticks every pack in the tab and locks them, since every divider needs codes to carry
        // its barcodes. The server selects all packs for installer tabs itself, so the disabled
        // (and therefore unsubmitted) pack checkboxes don't matter.
        if (e.target.matches('.installer-checkbox')) {
            const card = e.target.closest('.blueprint-card');
            if (card) {
                card.querySelectorAll('.pack-checkbox, input[id^="selectAllPacks_"]').forEach(function(checkbox) {
                    if (e.target.checked) checkbox.checked = true;
                    checkbox.disabled = e.target.checked;
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
                        // Installer tabs: the job barcodes live on the dividers, so start with them on
                        if (this.dataset.installer === "true" && checkbox) checkbox.checked = true;
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
    // Remembers the Start Cell, Job ID Cell and Store Name Col typed for EACH tab and fills
    // them back in next time: first by tab name (the same file uploaded again), then by tab
    // position (a new campaign file from the same template, where the tab names changed),
    // then the first tab's values. Only fields still showing the built-in defaults are
    // filled, so values the server re-rendered after "Update Previews" are never touched.
    const previewForm = document.querySelector('form[action="/preview"]');
    if (previewForm) {
        const ANCHOR_KEY = 'anchor_points_by_tab';
        const ANCHOR_FIELDS = { start: 'B8', job: 'E1', store: 'A' };  // field -> built-in default
        const MAX_REMEMBERED_TABS = 200;

        const tabCards = Array.from(previewForm.querySelectorAll('.tab-card')).map(function(card) {
            const tabBox = card.querySelector('input[name="selected_tabs"]');
            return {
                tab: tabBox ? tabBox.value : null,
                start: card.querySelector('input[name^="start_"]'),
                job: card.querySelector('input[name^="job_"]'),
                store: card.querySelector('input[name^="store_"]')
            };
        }).filter(c => c.tab && c.start && c.job && c.store);

        let memory = { byName: {}, byIndex: [] };
        let legacy = {};
        try {
            const saved = JSON.parse(localStorage.getItem(ANCHOR_KEY));
            if (saved && saved.byName && Array.isArray(saved.byIndex)) memory = saved;
            // Older versions remembered a single set of anchors (the first tab's) under these keys
            legacy = {
                start: localStorage.getItem('anchor_start'),
                job: localStorage.getItem('anchor_job'),
                store: localStorage.getItem('anchor_store')
            };
        } catch (e) { /* storage blocked or corrupt - just keep the defaults */ }

        // A. On page load, fill each tab from its own memory
        tabCards.forEach(function(c, i) {
            const remembered = memory.byName[c.tab] || memory.byIndex[i] || memory.byIndex[0] || legacy;
            Object.keys(ANCHOR_FIELDS).forEach(function(field) {
                if (remembered[field] && c[field].value === ANCHOR_FIELDS[field]) c[field].value = remembered[field];
            });
        });

        // B. On submit, save every tab's current inputs for next time
        previewForm.addEventListener('submit', function() {
            memory.byIndex = tabCards.map(c => ({ start: c.start.value, job: c.job.value, store: c.store.value }));
            tabCards.forEach(function(c, i) {
                delete memory.byName[c.tab];  // re-insert so the most recently used tabs sit last
                memory.byName[c.tab] = memory.byIndex[i];
            });
            const names = Object.keys(memory.byName);
            names.slice(0, Math.max(0, names.length - MAX_REMEMBERED_TABS)).forEach(n => delete memory.byName[n]);
            try {
                localStorage.setItem(ANCHOR_KEY, JSON.stringify(memory));
            } catch (e) { /* storage blocked - nothing to remember with */ }
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
    // 8b. DRAG & DROP FILE INPUTS (Global)
    // =========================================
    // The file input itself covers the whole dropzone (transparent), so the
    // browser's native drag/drop-onto-input behavior already works. This just
    // adds the hover highlight, the "selected file" label underneath, and a
    // "Clear" button (shown once a file is picked) that empties the input.
    document.querySelectorAll('.dropzone').forEach(function(zone) {
        const input = zone.querySelector('.dropzone-input');
        const filenameEl = zone.querySelector('.dropzone-filename');
        if (!input) return;

        const clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.className = 'dropzone-clear';
        clearBtn.textContent = '✕ Clear';
        clearBtn.title = 'Remove the selected file';
        zone.appendChild(clearBtn);  // after the input, so it sits on top of it

        clearBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            input.value = '';
            // Fire 'change' so everything listening (filename label, divider checkbox,
            // Step 1 submit-button validation) updates as if the user removed the file.
            input.dispatchEvent(new Event('change', { bubbles: true }));
        });

        function showFilename() {
            if (filenameEl) {
                filenameEl.textContent = input.files.length ? `✓ ${input.files[0].name}` : '';
            }
            clearBtn.hidden = !input.files.length;
        }

        input.addEventListener('change', showFilename);
        input.addEventListener('dragenter', () => zone.classList.add('dropzone-active'));
        input.addEventListener('dragleave', () => zone.classList.remove('dropzone-active'));
        input.addEventListener('drop', () => zone.classList.remove('dropzone-active'));
        showFilename();
    });

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
// =========================================
// 9. DUPLICATE STORE NAME ALERT (Step 1)
// =========================================    
document.addEventListener('DOMContentLoaded', function() {
    // Auto-open the Duplicate Store modal if it exists on page render
    const dupeModal = document.getElementById('duplicate-modal');
    if (dupeModal) {
        dupeModal.showModal();
    }
});
// =========================================
// 10. LABEL MAKER DRAG AND DROP CONFIGURATOR
// =========================================    
document.addEventListener("DOMContentLoaded", function() {
    const layoutList = document.getElementById('attribute-list');
    const headersList = document.getElementById('mapped-headers-list');
    const hiddenInput = document.getElementById('attribute_order');
    
    if(!layoutList || !hiddenInput) return;
    
    const updateHiddenInput = () => {
        const items = [...layoutList.querySelectorAll('.sortable-item')].map(item => item.getAttribute('data-id'));
        hiddenInput.value = items.join(',');
    };
    updateHiddenInput();

    let draggedItem = null;

    document.querySelectorAll('.sortable-item').forEach(item => {
        item.addEventListener('dragstart', function(e) {
            draggedItem = this;
            setTimeout(() => this.style.opacity = '0.4', 0);
        });
        item.addEventListener('dragend', function(e) {
            this.style.opacity = '1';
            draggedItem = null;
            updateHiddenInput();
        });
    });

    document.querySelectorAll('.drag-container').forEach(container => {
        container.addEventListener('dragover', function(e) {
            e.preventDefault();
            const afterElement = getDragAfterElement(this, e.clientY);
            if (draggedItem) {
                if (afterElement == null) {
                    this.appendChild(draggedItem);
                } else {
                    this.insertBefore(draggedItem, afterElement);
                }
            }
        });
    });

    function getDragAfterElement(container, y) {
        const draggableElements = [...container.querySelectorAll('.sortable-item:not([style*="opacity: 0.4"])')];
        return draggableElements.reduce(function(closest, child) {
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            if (offset < 0 && offset > closest.offset) {
                return { offset: offset, element: child };
            } else {
                return closest;
            }
        }, { offset: Number.NEGATIVE_INFINITY }).element;
    }
});
// =========================================
// 11. COURIER CONSIGNMENTS (packing labels preview)
// =========================================
// Address corrections and per-consignment service codes are kept in the hidden #consignment-edits
// field as JSON, keyed by consignment id, and sent with Update Previews and Generate.
document.addEventListener('DOMContentLoaded', function () {
    const editsField = document.getElementById('consignment-edits');
    if (!editsField) return;

    let edits = {};
    try { edits = JSON.parse(editsField.value || '{}') || {}; } catch (e) { edits = {}; }
    const save = () => { editsField.value = JSON.stringify(edits); };

    const ADDRESS = ['receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode'];
    const rows = [...document.querySelectorAll('.pl-con-row')];
    const mainService = document.getElementById('service_code');
    const applyBtn = document.getElementById('apply-service');
    const stateFilter = document.getElementById('courier-state-filter');
    const textFilter = document.getElementById('courier-text-filter');
    const countLabel = document.getElementById('courier-filter-count');

    const editOf = (id) => (edits[id] = edits[id] || {});
    const tidy = (id) => { if (edits[id] && !Object.keys(edits[id]).length) delete edits[id]; save(); };

    function show(row, d) {
        row.querySelector('.pl-show-receiver').textContent = d.receiver;
        row.querySelector('.pl-show-contact').textContent = d.contact ? 'Attn ' + d.contact : '';
        row.querySelector('.pl-show-atl').hidden = !d.authority_to_leave;
        row.querySelector('.pl-show-address').textContent =
            [d.line1, d.line2, d.suburb, d.state, d.postcode].filter(Boolean).join(', ');
        row.dataset.state = (d.state || '').toUpperCase();
    }
    function fill(editRow, d) {
        editRow.querySelectorAll('[data-field]').forEach((input) => {
            if (input.type === 'checkbox') input.checked = !!d[input.dataset.field];
            else input.value = d[input.dataset.field] || '';
        });
    }
    function read(editRow) {
        const d = {};
        editRow.querySelectorAll('[data-field]').forEach((input) => {
            d[input.dataset.field] = input.type === 'checkbox' ? input.checked : input.value.trim();
        });
        d.state = (d.state || '').toUpperCase();
        return d;
    }

    rows.forEach((row) => {
        const id = row.dataset.id;
        const editRow = row.nextElementSibling;
        const original = JSON.parse(row.dataset.original);
        const current = () => {
            const d = { ...original };
            ADDRESS.concat('authority_to_leave').forEach((k) => { if (edits[id] && k in edits[id]) d[k] = edits[id][k]; });
            return d;
        };
        const close = () => { editRow.hidden = true; };

        row.querySelector('.pl-edit-btn').addEventListener('click', () => {
            editRow.hidden = !editRow.hidden;
            if (!editRow.hidden) { fill(editRow, current()); editRow.querySelector('input').focus(); }
        });

        editRow.querySelector('.pl-edit-save').addEventListener('click', () => {
            const d = read(editRow);
            const changed = ADDRESS.some((k) => (d[k] || '') !== (original[k] || '')) ||
                            d.authority_to_leave !== !!original.authority_to_leave;
            const keep = edits[id] && edits[id].service_code ? { service_code: edits[id].service_code } : {};
            edits[id] = changed ? { ...keep, ...d } : keep;
            tidy(id);
            show(row, changed ? d : original);
            row.querySelector('.pl-tag-edited').hidden = !changed;
            row.querySelector('.pl-tag-book').hidden = true;
            close();
            applyFilters();
        });
        editRow.querySelector('.pl-edit-cancel').addEventListener('click', close);
        editRow.querySelector('.pl-edit-reset').addEventListener('click', () => {
            fill(editRow, original);
        });

        // Typing a receiver or street offers matches from the address book; picking one fills the form
        attachAddressSuggestions(editRow, (entry) => fill(editRow, entry));

        // Enter inside the edit form saves the row instead of submitting the whole page
        editRow.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); editRow.querySelector('.pl-edit-save').click(); }
            if (e.key === 'Escape') close();
        });

        const service = row.querySelector('.pl-service-input');
        service.addEventListener('change', () => {
            const value = service.value;
            if (value) editOf(id).service_code = value; else if (edits[id]) delete edits[id].service_code;
            tidy(id);
        });
    });

    // Main service code: the default for rows left blank, and "Apply to shown" copies it into the visible rows
    // Each row's first option ("Same as main (CODE)") follows the main dropdown
    mainService.addEventListener('change', () => {
        const label = mainService.value ? `Same as main (${mainService.value})` : 'Same as main';
        rows.forEach((row) => { row.querySelector('.pl-service-input').options[0].textContent = label; });
    });
    applyBtn.addEventListener('click', () => {
        const value = mainService.value;
        if (!value) { mainService.focus(); return; }
        rows.filter((row) => !row.hidden).forEach((row) => {
            row.querySelector('.pl-service-input').value = value;
            editOf(row.dataset.id).service_code = value;
        });
        save();
    });

    function applyFilters() {
        const state = stateFilter.value;
        const text = textFilter.value.trim().toLowerCase();
        let shown = 0;
        rows.forEach((row) => {
            const visible = (!state || row.dataset.state === state) &&
                            (!text || row.textContent.toLowerCase().includes(text));
            row.hidden = !visible;
            if (!visible) row.nextElementSibling.hidden = true;
            if (visible) shown += 1;
        });
        countLabel.textContent = shown === rows.length ? `${rows.length} consignments` : `Showing ${shown} of ${rows.length}`;
        applyBtn.textContent = shown === rows.length ? 'Apply to all' : `Apply to ${shown} shown`;
    }
    stateFilter.addEventListener('change', applyFilters);
    textFilter.addEventListener('input', applyFilters);
    textFilter.addEventListener('keydown', (e) => { if (e.key === 'Enter') e.preventDefault(); });
    applyFilters();
    save();
});

// =========================================
// 12. ADDRESS BOOK
// =========================================
// Type-ahead suggestions from the address book, used in the consignment edit form. Waits for a short
// pause in typing, cancels a request that a newer keystroke made stale, and shows at most 6 matches.
function attachAddressSuggestions(container, onPick) {
    const list = container.querySelector('.pl-suggest');
    if (!list) return;
    let timer = null, controller = null, items = [], active = -1;

    const hide = () => { list.hidden = true; list.replaceChildren(); items = []; active = -1; };
    const highlight = (i) => {
        active = i;
        [...list.children].forEach((li, n) => li.setAttribute('aria-selected', String(n === i)));
    };
    const pick = (entry) => { onPick(entry); hide(); };
    const render = (rows) => {
        list.replaceChildren();
        items = rows;
        active = -1;
        if (!rows.length) { hide(); return; }
        rows.forEach((entry) => {
            const li = document.createElement('li');
            li.setAttribute('role', 'option');
            const name = document.createElement('strong');
            name.textContent = entry.receiver + (entry.contact ? ` · Attn ${entry.contact}` : '');
            const where = document.createElement('span');
            where.textContent = entry.address;
            li.append(name, where);
            li.addEventListener('mousedown', (e) => { e.preventDefault(); pick(entry); });
            list.append(li);
        });
        list.hidden = false;
    };

    container.querySelectorAll('[data-field="receiver"], [data-field="line1"]').forEach((input) => {
        input.setAttribute('autocomplete', 'off');
        input.addEventListener('input', () => {
            clearTimeout(timer);
            const q = input.value.trim();
            if (q.length < 2) { hide(); return; }
            timer = setTimeout(async () => {
                if (controller) controller.abort();
                controller = new AbortController();
                try {
                    const res = await fetch(`/api/addresses?limit=6&q=${encodeURIComponent(q)}`, { signal: controller.signal });
                    render((await res.json()).rows || []);
                } catch (err) {
                    if (err.name !== 'AbortError') hide();
                }
            }, 150);
        });
        input.addEventListener('keydown', (e) => {
            if (list.hidden) return;
            if (e.key === 'ArrowDown') { e.preventDefault(); highlight(Math.min(active + 1, items.length - 1)); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); highlight(Math.max(active - 1, 0)); }
            else if (e.key === 'Enter' && active >= 0) { e.preventDefault(); e.stopPropagation(); pick(items[active]); }
            else if (e.key === 'Escape') { e.stopPropagation(); hide(); }
        });
        input.addEventListener('blur', () => setTimeout(hide, 150));
    });
}

// The Address Book page: instant search, add, edit and delete. Only one page of results (50) is ever
// in the page, so it stays quick with 50k+ addresses; "Load more" fetches the next page.
document.addEventListener('DOMContentLoaded', function () {
    const page = document.getElementById('address-book');
    if (!page) return;

    const search = document.getElementById('ab-search');
    const body = document.getElementById('ab-rows');
    const more = document.getElementById('ab-more');
    const status = document.getElementById('ab-status');
    const summary = document.getElementById('ab-summary');
    const template = document.getElementById('ab-edit-template');
    const PAGE = 50;
    let query = '', offset = 0, timer = null, controller = null, total = Number(page.dataset.total || 0);

    // The table header sticks just under the nav bar; follow the nav's real height if it changes
    const nav = document.querySelector('.sticky-nav');
    if (nav) {
        const setNavHeight = () => document.documentElement.style.setProperty('--nav-height', `${nav.offsetHeight}px`);
        setNavHeight();
        new ResizeObserver(setNavHeight).observe(nav);
    }

    const say = (text, isError) => {
        status.textContent = text;
        status.classList.toggle('ab-status-error', !!isError);
    };
    const when = (ts) => ts ? new Date(ts * 1000).toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' }) : '—';
    const cell = (text, cls) => { const td = document.createElement('td'); td.textContent = text; if (cls) td.className = cls; return td; };

    function rowFor(entry) {
        const tr = document.createElement('tr');
        tr.className = 'ab-row';
        tr.dataset.id = entry.id;
        const actions = document.createElement('td');
        actions.className = 'ab-actions';
        const edit = document.createElement('button');
        edit.type = 'button'; edit.className = 'ab-icon-btn'; edit.textContent = '✏️';
        edit.title = 'Edit'; edit.setAttribute('aria-label', `Edit ${entry.receiver}`);
        edit.addEventListener('click', () => openEditor(tr, entry));
        const del = document.createElement('button');
        del.type = 'button'; del.className = 'ab-icon-btn'; del.textContent = '🗑️';
        del.title = 'Delete'; del.setAttribute('aria-label', `Delete ${entry.receiver}`);
        del.addEventListener('click', () => remove(tr, entry));
        actions.append(edit, del);
        const who = cell(entry.receiver);
        if (entry.contact) { const c = document.createElement('div'); c.className = 'ab-detail'; c.textContent = `Attn ${entry.contact}`; who.append(c); }
        tr.append(actions, who, cell([entry.line1, entry.line2].filter(Boolean).join(', ')), cell(entry.suburb),
                  cell(entry.state), cell(entry.postcode), cell(entry.country), cell(entry.authority_to_leave ? 'Y' : ''),
                  cell(entry.use_count, 'ab-num'), cell(when(entry.last_used_at)));
        return tr;
    }

    async function load(reset) {
        if (reset) offset = 0;
        if (controller) controller.abort();
        controller = new AbortController();
        try {
            const res = await fetch(`/api/addresses?limit=${PAGE}&offset=${offset}&q=${encodeURIComponent(query)}`, { signal: controller.signal });
            const data = await res.json();
            if (reset) body.replaceChildren();
            data.rows.forEach((entry) => body.append(rowFor(entry)));
            offset += data.rows.length;
            more.hidden = !data.has_more;
            const shown = body.querySelectorAll('.ab-row').length;
            summary.textContent = query
                ? (shown ? `${shown}${data.has_more ? '+' : ''} matching` : 'No matches')
                : (total ? `${total.toLocaleString()} addresses · most recently used first` : 'No addresses yet. They are added automatically each time you generate packing labels, or use "Add address".');
        } catch (err) {
            if (err.name !== 'AbortError') say('Could not load addresses. Is the app still running?', true);
        }
    }

    function editorFor(entry) {
        const tr = template.content.firstElementChild.cloneNode(true);
        tr.querySelectorAll('[data-field]').forEach((input) => {
            const v = entry[input.dataset.field];
            if (input.type === 'checkbox') input.checked = !!v; else input.value = v == null ? '' : v;
        });
        return tr;
    }

    function readEditor(tr) {
        const d = {};
        tr.querySelectorAll('[data-field]').forEach((input) => {
            d[input.dataset.field] = input.type === 'checkbox' ? input.checked : input.value.trim();
        });
        return d;
    }

    function wireEditor(tr, onSave, onCancel) {
        const err = tr.querySelector('.ab-edit-error');
        const save = async () => {
            const problem = await onSave(readEditor(tr));
            if (problem) err.textContent = problem;
        };
        tr.querySelector('.ab-save').addEventListener('click', save);
        tr.querySelector('.ab-cancel').addEventListener('click', onCancel);
        tr.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); save(); }
            if (e.key === 'Escape') onCancel();
        });
        tr.querySelector('input').focus();
    }

    async function send(url, method, data) {
        const res = await fetch(url, { method, headers: { 'Content-Type': 'application/json' }, body: data ? JSON.stringify(data) : undefined });
        const json = res.status === 204 ? {} : await res.json();
        return res.ok ? { ok: true, entry: json } : { ok: false, error: json.error || 'Something went wrong.' };
    }

    function openEditor(tr, entry) {
        if (tr.nextElementSibling && tr.nextElementSibling.classList.contains('ab-edit-row')) return;
        const editor = editorFor(entry);
        tr.after(editor);
        wireEditor(editor, async (data) => {
            const r = await send(`/api/addresses/${entry.id}`, 'PUT', data);
            if (!r.ok) return r.error;
            const fresh = rowFor(r.entry);
            fresh.classList.add('ab-flash');
            tr.replaceWith(fresh);
            editor.remove();
            say(`Saved ${r.entry.receiver}.`);
        }, () => editor.remove());
    }

    async function remove(tr, entry) {
        if (!confirm(`Delete ${entry.receiver}, ${entry.address} from the address book?`)) return;
        const r = await send(`/api/addresses/${entry.id}`, 'DELETE');
        if (!r.ok) { say(r.error, true); return; }
        const next = tr.nextElementSibling;
        if (next && next.classList.contains('ab-edit-row')) next.remove();
        tr.remove();
        total -= 1;
        say(`Deleted ${entry.receiver}.`);
    }

    document.getElementById('ab-add').addEventListener('click', () => {
        if (body.querySelector('.ab-edit-row.ab-new')) return;
        const editor = editorFor({ country: 'AU' });
        editor.classList.add('ab-new');
        body.prepend(editor);
        wireEditor(editor, async (data) => {
            const r = await send('/api/addresses', 'POST', data);
            if (!r.ok) return r.error;
            const fresh = rowFor(r.entry);
            fresh.classList.add('ab-flash');
            editor.replaceWith(fresh);
            total += 1;
            say(`Added ${r.entry.receiver}.`);
        }, () => editor.remove());
    });

    search.addEventListener('input', () => {
        clearTimeout(timer);
        timer = setTimeout(() => { query = search.value.trim(); load(true); }, 120);
    });
    search.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') e.preventDefault();
        if (e.key === 'Escape') { search.value = ''; query = ''; load(true); }
    });
    more.addEventListener('click', () => load(false));
    load(true);
});

