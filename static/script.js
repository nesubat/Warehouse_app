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
    // Files picked in an upload box survive a refresh: kept in this browser (IndexedDB) until the form is sent or
    // the box is cleared (a day at most), and put back in the box when the page opens again.
    const pickedFiles = (() => {
        const db = () => new Promise((resolve, reject) => {
            const req = indexedDB.open('warehouse-picked-files', 1);
            req.onupgradeneeded = () => req.result.createObjectStore('files');
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
        const run = async (mode, work) => {
            const store = (await db()).transaction('files', mode).objectStore('files');
            return new Promise((resolve, reject) => { const req = work(store); req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); });
        };
        const safe = (p) => p.catch(() => undefined);  // private window / storage blocked: just don't keep them
        return {
            get: (key) => safe(run('readonly', (s) => s.get(key))),
            set: (key, files) => safe(run('readwrite', (s) => s.put({ saved: Date.now(), files }, key))),
            drop: (key) => safe(run('readwrite', (s) => s.delete(key))),
        };
    })();

    document.querySelectorAll('.dropzone').forEach(function(zone) {
        const input = zone.querySelector('.dropzone-input');
        const filenameEl = zone.querySelector('.dropzone-filename');
        if (!input) return;
        const keepKey = `${location.pathname}|${input.name}`;

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

        // A box that takes several files (multiple) adds each pick or drop to what's already chosen
        let chosen = [];
        function showFilename() {
            if (filenameEl) {
                const names = [...input.files].map((f) => f.name);
                filenameEl.textContent = !names.length ? '' :
                    names.length === 1 ? `✓ ${names[0]}` : `✓ ${names.length} files: ${names.join(', ')}`;
            }
            clearBtn.hidden = !input.files.length;
        }
        function keepAdding() {
            if (!input.multiple) return;
            if (!input.files.length) { chosen = []; return; }  // cleared
            const seen = new Set(chosen.map((f) => `${f.name}|${f.size}`));
            [...input.files].forEach((f) => { if (!seen.has(`${f.name}|${f.size}`)) { chosen.push(f); seen.add(`${f.name}|${f.size}`); } });
            const all = new DataTransfer();
            chosen.forEach((f) => all.items.add(f));
            input.files = all.files;
        }

        // A picked file is read into memory straight away and that copy is what's sent. The browser otherwise reads
        // it from disk only when the form goes, and gives up (an error page, ERR_UPLOAD_FILE_CHANGED / ERR_FAILED)
        // if it changed in between, e.g. it was saved in Excel after picking it. A file that can't be read at all
        // says so here instead.
        const copies = new WeakSet();
        let reading = null;
        async function takeCopies() {
            const files = [...input.files];
            if (!files.length || files.every((f) => copies.has(f))) return;
            try {
                const read = await Promise.all(files.map(async (f) => (copies.has(f) ? f
                    : new File([await f.arrayBuffer()], f.name, { type: f.type, lastModified: f.lastModified }))));
                read.forEach((f) => copies.add(f));
                const all = new DataTransfer();
                read.forEach((f) => all.items.add(f));
                input.files = all.files;
                if (input.multiple) chosen = read;
                pickedFiles.set(keepKey, read);
            } catch (e) {
                input.value = '';
                chosen = [];
                showFilename();
                pickedFiles.drop(keepKey);
                if (filenameEl) filenameEl.textContent = `⚠ Couldn't read ${files.map((f) => f.name).join(', ')}. ` +
                    'If it’s open in Excel, save it (or close it), then pick it again.';
            }
        }
        input.addEventListener('change', () => {
            keepAdding(); showFilename();
            if (!input.files.length) { pickedFiles.drop(keepKey); return; }
            reading = takeCopies().finally(() => { reading = null; });
        });
        // Sent while a copy is still being read: wait for it, then send (with the button that was pressed)
        if (input.form) input.form.addEventListener('submit', (e) => {
            if (!reading) return;
            e.preventDefault();
            e.stopImmediatePropagation();
            const by = e.submitter;
            reading.then(() => { if (input.files.length || !input.required) input.form.requestSubmit(by || undefined); });
        }, true);
        input.addEventListener('dragenter', () => zone.classList.add('dropzone-active'));
        input.addEventListener('dragleave', () => zone.classList.remove('dropzone-active'));
        input.addEventListener('drop', () => zone.classList.remove('dropzone-active'));
        // Sent: nothing to keep (a check that stops the form runs first and leaves them kept)
        if (input.form) input.form.addEventListener('submit', () => pickedFiles.drop(keepKey));
        showFilename();
        // Picked before a refresh: put them back
        if (!input.files.length) {
            pickedFiles.get(keepKey).then((kept) => {
                if (!kept || input.files.length || Date.now() - kept.saved > 86400000 || !(kept.files || []).length) return;
                const all = new DataTransfer();
                kept.files.forEach((f) => all.items.add(f));
                input.files = all.files;
                chosen = [...kept.files];
                input.dispatchEvent(new Event('change', { bubbles: true }));
                if (filenameEl) filenameEl.textContent += ' (kept from before the refresh)';
            });
        }
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

    const ADDRESS = ['receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode', 'country'];
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
        const country = (d.country || 'AU').toUpperCase();
        row.querySelector('.pl-show-address').textContent =
            [d.line1, d.line2, d.suburb, d.state, d.postcode, country !== 'AU' ? country : ''].filter(Boolean).join(', ');
        row.dataset.state = (d.state || '').toUpperCase();
        row.dataset.country = country;
        document.dispatchEvent(new CustomEvent('pl-countries-changed'));  // the export details' count follows
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
        d.suburb = (d.suburb || '').toUpperCase();
        if ('country' in d) d.country = (d.country || '').toUpperCase();
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

        // The form's Service Code list mirrors the row's: filled from it when the form opens, copied back on Save
        const editService = editRow.querySelector('.pl-edit-service');
        row.querySelector('.pl-edit-btn').addEventListener('click', () => {
            editRow.hidden = !editRow.hidden;
            if (!editRow.hidden) {
                fill(editRow, current());
                if (editService) editService.value = row.querySelector('.pl-service-input').value;
                editRow.querySelector('input').focus();
            }
        });

        editRow.querySelector('.pl-edit-save').addEventListener('click', () => {
            const rowService = row.querySelector('.pl-service-input');
            if (editService && rowService.value !== editService.value) {
                rowService.value = editService.value;
                rowService.dispatchEvent(new Event('change'));  // the row's own handler keeps it in the edits
            }
            const d = read(editRow);
            const changed = ADDRESS.some((k) => (d[k] || '') !== (original[k] || '')) ||
                            d.authority_to_leave !== !!original.authority_to_leave;
            const keep = edits[id] && edits[id].service_code ? { service_code: edits[id].service_code } : {};
            // Saving counts as checked, so a row that isn't in the address book stops being flagged
            // (kept through Update Previews; the address is saved to the book on Generate)
            const checked = row.classList.contains('pl-needs-check') || (edits[id] && edits[id].checked) ? { checked: true } : {};
            edits[id] = changed ? { ...keep, ...checked, ...d } : { ...keep, ...checked };
            tidy(id);
            show(row, changed ? d : original);
            row.querySelector('.pl-tag-edited').hidden = !changed;
            row.querySelector('.pl-tag-book').hidden = true;
            markChecked(row);
            if (d.postcode) markAddressed(row);
            close();
            applyFilters();
        });
        editRow.querySelector('.pl-edit-cancel').addEventListener('click', close);

        // Closest saved addresses: clicking one fills the form; typing in any field re-ranks the list
        const closest = attachClosest(editRow, (entry) => fill(editRow, entry), () => read(editRow));
        editRow.querySelector('.pl-edit-reset').addEventListener('click', () => {
            fill(editRow, original);
            closest.refresh();  // the Excel address's closest matches, the address-book fill among them
        });

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

    // A flagged row the user has saved: drop the amber theme and update the count above the table
    const checkNote = document.getElementById('courier-check-note');
    function markChecked(row) {
        if (!row.classList.contains('pl-needs-check')) return;
        row.classList.remove('pl-needs-check');
        const tag = row.querySelector('.pl-tag-check');
        if (tag) tag.remove();
        const left = rows.filter((r) => r.classList.contains('pl-needs-check')).length;
        if (checkNote) {
            if (!left) checkNote.hidden = true;
            else checkNote.querySelector('strong').textContent =
                `${left} ${left === 1 ? 'address isn’t' : 'addresses aren’t'} in the address book yet.`;
        }
    }

    // A row that had no address and now has one: drop the red theme (the server regroups it on refresh)
    const missingNote = document.getElementById('courier-missing-note');
    function markAddressed(row) {
        if (!row.classList.contains('pl-no-address')) return;
        row.classList.remove('pl-no-address');
        const tag = row.querySelector('.pl-tag-missing');
        if (tag) tag.remove();
        const left = rows.filter((r) => r.classList.contains('pl-no-address')).length;
        if (missingNote && !left) missingNote.hidden = true;
        else if (missingNote) missingNote.querySelector('strong').textContent =
            `${left} consignment${left === 1 ? ' has' : 's have'} no address`;
    }

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
        // A group's heading goes with its rows
        document.querySelectorAll('.pl-group-row').forEach((g) => {
            g.hidden = rows.filter((r) => r.dataset.group === g.dataset.group).every((r) => r.hidden);
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
// The "Closest saved addresses" list in a consignment's ✏️ form. The server renders the first list; typing in
// any field asks /api/addresses/suggest again with everything typed so far (after a short pause, cancelling a
// request a newer keystroke made stale), so the most promising saved addresses stay on top. A click fills the form.
function attachClosest(container, onPick, readForm) {
    const box = container.querySelector('.pl-closest');
    if (!box) return { refresh() {} };
    const list = box.querySelector('.pl-closest-list');
    const title = box.querySelector('.pl-closest-title');
    const none = box.querySelector('.pl-closest-none');
    const flagged = ['unverified', 'missing'].includes(box.dataset.status);
    let timer = null, controller = null;

    const wire = (button) => {
        button.setAttribute('aria-pressed', 'false');
        button.addEventListener('click', () => {
            onPick(JSON.parse(button.dataset.entry));
            list.querySelectorAll('.pl-closest-item').forEach((b) => b.setAttribute('aria-pressed', String(b === button)));
        });
    };
    const span = (cls, text) => { const el = document.createElement('span'); el.className = cls; el.textContent = text; return el; };
    const render = (rows) => {
        list.replaceChildren();
        rows.forEach((s) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'pl-closest-item';
            button.dataset.entry = JSON.stringify({ receiver: s.receiver, contact: s.contact, line1: s.line1, line2: s.line2,
                suburb: s.suburb, state: s.state, postcode: s.postcode, authority_to_leave: s.authority_to_leave });
            const name = document.createElement('strong');
            name.textContent = s.receiver + (s.contact ? ` · Attn ${s.contact}` : '');
            button.append(span('pl-closest-why', s.why), name, span('pl-closest-address', s.address + (s.verified_at ? ' · ✓ Verified' : '')));
            wire(button);
            const li = document.createElement('li');
            li.append(button);
            list.append(li);
        });
        title.hidden = !rows.length;
        none.hidden = rows.length > 0 || !flagged;
    };
    const refresh = () => {
        clearTimeout(timer);
        timer = setTimeout(async () => {
            if (controller) controller.abort();
            controller = new AbortController();
            const form = readForm();
            const params = new URLSearchParams();
            ['receiver', 'contact', 'line1', 'line2', 'suburb', 'state', 'postcode'].forEach((k) => { if (form[k]) params.set(k, form[k]); });
            try {
                const res = await fetch(`/api/addresses/suggest?${params}`, { signal: controller.signal });
                render((await res.json()).rows || []);
            } catch (err) {
                if (err.name !== 'AbortError') title.hidden = true;
            }
        }, 200);
    };

    list.querySelectorAll('.pl-closest-item').forEach(wire);
    container.querySelectorAll('[data-field]').forEach((input) => {
        if (input.type !== 'checkbox') { input.setAttribute('autocomplete', 'off'); input.addEventListener('input', refresh); }
    });
    return { refresh };
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
        if (entry.verified_at) {
            const v = document.createElement('span');
            v.className = 'ab-verified';
            v.textContent = '✓ Verified';
            v.title = `Confirmed by the courier portal on ${when(entry.verified_at)}`;
            who.append(v);
        }
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

// =========================================
// 13. PACKING SPECS PAGE + PALLET WEIGHTS
// =========================================
// "+ Add spec" adds a blank row from the <template>; each row gets the next free index for its field names.
document.addEventListener('DOMContentLoaded', function () {
    const add = document.getElementById('ps-add');
    if (!add) return;
    const body = document.querySelector('#ps-named tbody');
    const template = document.getElementById('ps-row-template');
    let next = Number(add.dataset.next || 0);

    add.addEventListener('click', () => {
        const html = template.innerHTML.replaceAll('__i__', String(next));
        next += 1;
        const holder = document.createElement('tbody');
        holder.innerHTML = html.trim();
        const row = holder.firstElementChild;
        body.append(row);
        row.querySelector('.ps-name').focus();
    });
    body.addEventListener('click', (e) => {
        const drop = e.target.closest('.ps-drop-new');
        if (drop) drop.closest('tr').remove();
    });
});

// Packing Labels preview: a carton's weight typed in the consignment table is kept in the hidden
// #carton-weights field ({pack key: kg}) and sent with Update Previews and Generate. Clearing the box
// goes back to the Packing Specs weight.
document.addEventListener('DOMContentLoaded', function () {
    const field = document.getElementById('carton-weights');
    if (!field) return;
    let weights = {};
    try { weights = JSON.parse(field.value || '{}') || {}; } catch (e) { weights = {}; }

    document.querySelectorAll('.pl-weight-input').forEach((input) => {
        input.addEventListener('change', () => {
            const value = input.value.trim();
            const number = Number(value);
            const ok = value !== '' && Number.isFinite(number) && number > 0 && number <= 5000;
            input.classList.toggle('pl-weight-bad', value !== '' && !ok);
            const changed = ok && number !== Number(input.dataset.default);
            if (changed) weights[input.dataset.pack] = number;
            else delete weights[input.dataset.pack];
            if (value === '') input.value = input.dataset.default;
            input.classList.toggle('pl-weight-changed', changed);
            field.value = JSON.stringify(weights);
        });
        // Enter inside the weight box must not submit the whole page
        input.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); input.blur(); } });
    });
});

// =========================================
// 14. ENTER AND REQUIRED BOXES IN THE PACKING LABELS FORM
// =========================================
// Enter in a box would submit the form with its first button, Update Previews (the page reloads and scrolls,
// nothing is generated). Instead: Enter in the Project Name presses Generate, Enter in a Header Row presses
// Update Previews, and Enter anywhere else does nothing.
document.addEventListener('DOMContentLoaded', function () {
    const generate = document.querySelector('button[name="generate"]');
    const form = generate && generate.form;
    if (!form) return;
    const preview = form.querySelector('button[name="preview"]');
    // Generate: every required box filled in, else a message above the button naming them (and nothing is sent).
    // Update Previews: header rows must be numbers. Runs first (capture), so the loading overlay only shows when
    // the form really goes.
    const error = document.getElementById('pl-generate-error');
    const nameOf = (box) => {
        const label = box.id && form.querySelector(`label[for="${box.id}"]`);
        return (label ? label.textContent : box.name).replace(/[:*]\s*$/, '').trim();
    };
    form.addEventListener('submit', (e) => {
        const by = e.submitter && e.submitter.name;
        form.querySelectorAll('.pl-missing').forEach((b) => b.classList.remove('pl-missing'));
        let bad = [];
        if (by === 'generate') {
            bad = [...form.querySelectorAll('[required]')].filter((b) => !b.disabled && !String(b.value).trim());
        } else if (by === 'preview') {
            bad = [...form.querySelectorAll('input[name^="header_row_"]')].filter((b) => !(parseInt(b.value, 10) >= 1));
        }
        if (!bad.length) {
            if (error) error.hidden = true;
            return;
        }
        e.preventDefault();
        e.stopImmediatePropagation();
        bad.forEach((b) => b.classList.add('pl-missing'));
        const text = by === 'generate' ? `Fill in ${bad.map(nameOf).join(', ')} before generating.`
                                       : 'Each Header Row needs a row number (1 or more).';
        if (error && by === 'generate') { error.textContent = `⚠️ ${text}`; error.hidden = false; }
        else alert(text);
        bad[0].scrollIntoView({ block: 'center', behavior: 'smooth' });
        bad[0].focus({ preventScroll: true });
    }, true);
    form.addEventListener('input', (e) => { if (e.target.classList.contains('pl-missing') && String(e.target.value).trim()) e.target.classList.remove('pl-missing'); });
    form.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter' || e.isComposing) return;
        const box = e.target;
        if (!(box instanceof HTMLInputElement) || ['button', 'submit', 'checkbox', 'radio', 'file'].includes(box.type)) return;
        if (e.defaultPrevented) return;  // boxes with their own Enter (edit form, weights, filter)
        e.preventDefault();
        if (box.name === 'project_name') form.requestSubmit(generate);
        else if (box.name.startsWith('header_row_') && preview) form.requestSubmit(preview);
    });
});

// =========================================
// 15. UPDATE IN BULK: REVIEW CHOICES
// =========================================
// Each row of the portal file is a card with one choice: update the matched saved address, add it as new,
// merge it into one of the closest saved addresses, use one conflicting row, or skip. Tabs filter the cards,
// the bulk buttons set every shown card at once, and the bar by Apply says what will be saved.
document.addEventListener('DOMContentLoaded', function () {
    const cards = [...document.querySelectorAll('.av-card')];
    if (!cards.length) return;
    const plan = document.getElementById('av-plan');
    const apply = document.getElementById('av-apply');
    const tabs = [...document.querySelectorAll('.av-tab')];
    const radios = (card) => [...card.querySelectorAll('input[type="radio"]')];
    const chosen = (card) => radios(card).find((r) => r.checked);
    const suggested = new Map(cards.map((card) => [card, chosen(card)]));
    const OUTCOME = { update: 'Will update the saved address', new: 'Will be added as new', merge: 'Will be merged into a saved address', skip: 'Skipped' };

    function refresh() {
        const n = { update: 0, new: 0, merge: 0, skip: 0 };
        cards.forEach((card) => {
            const kind = chosen(card) ? chosen(card).dataset.kind : 'skip';
            n[kind] += 1;
            card.dataset.choice = kind;
            card.querySelector('.av-card-outcome').textContent = OUTCOME[kind];
            radios(card).forEach((r) => r.closest('.av-choice').classList.toggle('av-choice-on', r.checked));
        });
        const doing = n.update + n.new + n.merge;
        const parts = [n.update && `${n.update} update${n.update === 1 ? '' : 's'}`, n.new && `${n.new} new`,
                       n.merge && `${n.merge} merge${n.merge === 1 ? '' : 's'}`, n.skip && `${n.skip} skipped`].filter(Boolean);
        plan.textContent = parts.join(' · ');
        apply.textContent = doing ? `Apply ${doing} change${doing === 1 ? '' : 's'}` : 'Apply (nothing chosen)';
    }
    cards.forEach((card) => radios(card).forEach((r) => r.addEventListener('change', refresh)));

    tabs.forEach((tab) => tab.addEventListener('click', () => {
        tabs.forEach((t) => t.setAttribute('aria-selected', String(t === tab)));
        cards.forEach((card) => { card.hidden = tab.dataset.filter !== 'all' && card.dataset.status !== tab.dataset.filter; });
    }));
    document.querySelectorAll('.av-bulk-btn').forEach((button) => button.addEventListener('click', () => {
        cards.filter((card) => !card.hidden).forEach((card) => {
            const want = button.dataset.set;
            let pick = null;
            if (want === 'suggested') pick = suggested.get(card);
            else if (want === 'best-merge') pick = radios(card).find((r) => r.dataset.kind === 'merge');
            else pick = radios(card).find((r) => r.dataset.kind === want);
            if (pick) pick.checked = true;  // a card without that option keeps its choice
        });
        refresh();
    }));
    refresh();
});

// =========================================
// 16. MATCH COURIER LABELS BY HAND (stitch_match.html)
// =========================================
// Two areas, top to bottom:
//   Pairs          courier label beside its packing label, grouped by receiver and address. They start as the
//                  suggested pairs (an address with as many courier labels as packing labels is paired in order),
//                  amber, to check. ✓ confirms one: its two boxes close together into one green box, in place, and
//                  the page scrolls by exactly the distance to the next ✓, so it lands under the pointer (click,
//                  click, click in one spot). ✕ splits a pair; a matched pair shows ↶ Undo on hover.
//   Still to match each courier label with no pair beside its closest packing label still free ("closest 60%"),
//                  with ✓ to match them; "Other…" picks another packing label for it. Packing labels nobody's
//                  closest to are listed under it. Dragging (or clicking one card, then another) still works anywhere.
// Courier cards show the address printed on the label in bold and its Item Ref in bold red. "Match all 100%"
// confirms every sure suggestion at once. Save sends every pair ({courier key: item reference}), confirmed or
// still to check, and stitches again in the background behind an overlay.
document.addEventListener('DOMContentLoaded', function () {
    const root = document.getElementById('stitch-match');
    if (!root) return;
    let session = { unmatched: [], waiting: [] };
    try { session = JSON.parse(root.dataset.session || '{}'); } catch (e) { /* empty page */ }
    const pairList = document.getElementById('sm-pairs');
    const freeList = document.getElementById('sm-free');
    const packList = document.getElementById('sm-packs');
    const search = document.getElementById('sm-search');
    const saveBtn = document.getElementById('sm-save');
    const finaliseBtn = document.getElementById('sm-finalise');
    const sureBtn = document.getElementById('sm-sure');
    const overlay = document.getElementById('loading-overlay');
    const status = document.getElementById('sm-status');
    const hintbar = document.getElementById('sm-hintbar');
    const hintText = document.getElementById('sm-hint-text');
    const zoom = document.getElementById('sm-zoom');
    const motion = !window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    // Full-page overlay while saving / stitching again: one click only, and nothing to change under it
    const busy = (title, sub) => {
        document.getElementById('loading-text').textContent = title;
        if (sub) document.getElementById('loading-sub').textContent = sub;
        overlay.style.display = 'flex';
    };
    const idle = () => { overlay.style.display = 'none'; };
    const imageUrl = (i, size) => root.dataset.imageUrl.replace(/0\.png$/, `${i}.png`) + (size ? `?size=${size}` : '');
    const packs = Object.fromEntries((session.waiting || []).map((w) => [w.item_reference, w]));
    const indexOf = Object.fromEntries(session.unmatched.map((u, i) => [u.key, i]));
    const byKey = Object.fromEntries(session.unmatched.map((u) => [u.key, u]));
    const pairs = {};          // courier key -> item reference
    const suggested = {};      // courier key -> true while the pair is still the suggested one (to check)
    const fromSuggestion = {}; // courier key -> true when a confirmed pair was its suggestion (Undo puts it back to check)
    const choice = {};         // courier key -> packing label picked with "Other…" in Still to match
    let selected = null;       // {side: 'courier' | 'pack', id}
    let saved = true;
    let finished = false;      // saved and stitched: the page shows the result until it's opened again
    let renderTimer = null;

    const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
    const plural = (n, word) => `${n} ${word}${n === 1 ? '' : (/(s|x|ch|sh)$/.test(word) ? 'es' : 's')}`;
    const packName = (w) => `${w.store} · Label ${w.label} of ${w.of}${w.boxes > 1 ? ` · box ${w.box} of ${w.boxes}` : ''}`;
    const courierOf = (ref) => Object.keys(pairs).find((k) => pairs[k] === ref);
    const say = (text, bad) => { status.textContent = text; status.classList.toggle('sm-status-bad', !!bad); };
    const placeKey = (w) => (w.not_in_csv ? 'none' : String(w.consignment));
    const itemRef = (u) => u.item_ref || ((u.reason || '').match(/Item Ref '([^']*)'/) || [])[1] || '';
    const why = (u) => u.why || (u.reason || '').replace(/^Item Ref '[^']*' /, '');
    const scoreOf = (u, ref) => ((u.candidates || []).find((c) => c[0] === ref) || [null, 0])[1];
    const pct = (x) => `${Math.round((x || 0) * 100)}%`;
    const label = (u) => `courier label ${(u.snippet || [])[1] || u.file} ${u.page}`;

    function restoreSuggestions() {
        session.unmatched.forEach((u) => {
            if (u.suggestion && packs[u.suggestion] && !pairs[u.key] && !courierOf(u.suggestion)) { pairs[u.key] = u.suggestion; suggested[u.key] = true; }
        });
    }
    const sure = (u) => u.suggestion && packs[u.suggestion] && (u.score || 0) >= 0.995;
    const sureLeft = () => session.unmatched.filter((u) => sure(u) && !(pairs[u.key] === u.suggestion && !suggested[u.key])
                                                       && (!courierOf(u.suggestion) || courierOf(u.suggestion) === u.key));
    const rowOf = (list, key) => list.querySelector(`[data-key="${CSS.escape(key)}"].sm-pair`);
    const visibleRows = (list) => [...list.querySelectorAll('.sm-pair')].filter((r) => !r.hidden && !r.closest('[hidden]'));

    // ---------- keeping the next ✓ under the pointer ----------
    // anchor: where the clicked ✓ was on screen. The page scrolls by the distance from there to the next ✓.
    // smooth: a scroll the eye follows (✓ in place); instant: making up for rows that moved (a pair made in Still
    // to match goes up into Pairs, so everything under it shifts).
    function bringUnder(next, anchor, smooth) {
        if (!next) return;
        const tick = next.querySelector('.sm-tick');
        if (anchor != null && tick) {
            const delta = tick.getBoundingClientRect().top - anchor;
            if (Math.abs(delta) > 1) window.scrollBy({ top: delta, behavior: smooth && motion ? 'smooth' : 'instant' });
        }
        next.classList.remove('sm-next'); void next.offsetWidth; next.classList.add('sm-next');
        if (tick) tick.focus({ preventScroll: true });   // Enter again confirms the next one
    }
    const nextAfter = (rows, row, wanted) => {
        const at = rows.indexOf(row);
        return rows.slice(at + 1).find(wanted) || rows.slice(0, Math.max(0, at)).find(wanted) || null;
    };
    const toCheckRow = (r) => r.classList.contains('sm-pair-suggested');
    // Re-draws group headings and counts a moment after a ✓, once the boxes have closed; the page looks the same
    const renderSoon = () => { clearTimeout(renderTimer); renderTimer = setTimeout(() => render(), 650); };

    // ---------- pairing ----------
    // ✓ on a pair to check: its two boxes close together into one green box, right where it is
    function confirmRow(key, tick) {
        if (!suggested[key]) return;
        const row = rowOf(pairList, key);
        const anchor = tick ? tick.getBoundingClientRect().top : null;
        const next = row ? nextAfter(visibleRows(pairList), row, toCheckRow) : null;
        delete suggested[key]; fromSuggestion[key] = true; saved = false;
        if (row) setDone(row, byKey[key]);
        updateCounts();
        bringUnder(next, anchor, true);
        renderSoon();
    }
    // ✓ in Still to match, or dragged / clicked together: a matched pair. A packing label takes one courier label.
    function pair(key, ref, tick) {
        clearTimeout(renderTimer);
        const fromFree = tick && tick.closest('#sm-free');
        const anchor = tick ? tick.getBoundingClientRect().top : null;
        const list = fromFree ? freeList : pairList;
        const before = visibleRows(list);
        const row = tick ? tick.closest('.sm-pair') : null;
        const nextKey = row ? (nextAfter(before, row, fromFree ? () => true : toCheckRow) || {}).dataset?.key : null;
        const other = courierOf(ref);
        const wasToCheck = !!other && !!suggested[other];
        if (other && other !== key) { delete pairs[other]; delete suggested[other]; delete fromSuggestion[other]; }
        pairs[key] = ref; delete suggested[key]; delete choice[key];
        fromSuggestion[key] = byKey[key].suggestion === ref;
        selected = null; saved = false;
        render();
        const made = rowOf(pairList, key);
        if (made) made.classList.add('sm-arrived');
        if (nextKey) bringUnder(rowOf(list, nextKey), anchor, !fromFree);
        else if (wasToCheck) bringUnder(nextAfter(visibleRows(pairList), made, toCheckRow), null, true);
    }
    function unpair(key) { clearTimeout(renderTimer); delete pairs[key]; delete suggested[key]; delete fromSuggestion[key]; saved = false; render(); }
    // ↶ Undo on a matched pair: back to check if it was the suggestion, else both labels back to Still to match
    function undoPair(key) {
        clearTimeout(renderTimer);
        if (fromSuggestion[key] && pairs[key] === byKey[key].suggestion) { suggested[key] = true; delete fromSuggestion[key]; saved = false; render(); }
        else unpair(key);
        say('Match undone.');
    }

    // ---------- cards ----------
    function courierCard(u) {
        const i = indexOf[u.key];
        const card = el('article', 'sm-card sm-courier');
        card.draggable = true; card.tabIndex = 0; card.dataset.key = u.key;
        const ref = itemRef(u);
        card.dataset.search = [u.file, ref, why(u), ...(u.snippet || [])].join(' ').toLowerCase();
        const img = el('img', 'sm-thumb'); img.src = imageUrl(i); img.loading = 'lazy';
        img.alt = `Courier label (${u.file}, page ${u.page})`; img.title = `${u.file} · page ${u.page} — click to zoom`;
        img.addEventListener('click', (e) => { e.stopPropagation(); zoom.querySelector('img').src = imageUrl(i, 'zoom'); zoom.showModal(); });
        const body = el('div', 'sm-card-body');
        // The address printed on the label, in bold: what the packing label has to match
        const addr = el('div', 'sm-snippet');
        (u.snippet || []).forEach((line) => addr.append(el('div', null, line)));
        if (!(u.snippet || []).length) addr.append(el('div', 'ab-detail', 'No address text on this label'));
        body.append(addr);
        if (ref) body.append(el('div', 'sm-itemref', `Item Ref: ${ref}`));
        body.append(el('div', 'ab-detail sm-reason', why(u)));
        card.append(img, body);
        card.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/plain', u.key); card.classList.add('sm-dragging'); });
        card.addEventListener('dragend', () => card.classList.remove('sm-dragging'));
        card.addEventListener('click', () => choose('courier', u.key));
        card.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); choose('courier', u.key); } });
        card.classList.toggle('sm-selected', !!selected && selected.side === 'courier' && selected.id === u.key);
        return card;
    }
    function packCard(w) {
        const card = el('article', 'sm-card sm-pack');
        card.tabIndex = 0; card.dataset.ref = w.item_reference;
        card.dataset.search = [w.store, w.receiver, w.contact, w.address, w.packing_spec, w.item_reference, `label ${w.label}`].join(' ').toLowerCase();
        const body = el('div', 'sm-card-body');
        const title = el('div', 'sm-card-title', packName(w));
        title.append(el('span', 'sm-spec', w.packing_spec));
        body.append(title, el('div', 'sm-receiver', w.receiver + (w.contact ? ` · Attn ${w.contact}` : '')), el('div', 'sm-address', w.address));
        const tags = el('div', 'sm-tags');
        if (w.installer) tags.append(el('span', 'pl-courier-tag sm-tag-installer', 'Installer'));
        if (w.not_in_csv) tags.append(el('span', 'pl-courier-tag sm-tag-missing', 'Not in courier CSV'));
        if (w.page) tags.append(el('span', 'ab-detail', `packing label page ${w.page}`));
        body.append(tags);
        card.append(body);
        card.addEventListener('dragover', (e) => { e.preventDefault(); card.classList.add('sm-drop'); });
        card.addEventListener('dragleave', () => card.classList.remove('sm-drop'));
        card.addEventListener('drop', (e) => { e.preventDefault(); card.classList.remove('sm-drop'); const k = e.dataTransfer.getData('text/plain'); if (k) pair(k, w.item_reference); });
        card.addEventListener('click', () => choose('pack', w.item_reference));
        card.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); choose('pack', w.item_reference); } });
        card.classList.toggle('sm-selected', !!selected && selected.side === 'pack' && selected.id === w.item_reference);
        return card;
    }
    function choose(side, id) {
        if (selected && selected.side !== side) {
            pair(side === 'courier' ? id : selected.id, side === 'pack' ? id : selected.id);
            return;
        }
        selected = selected && selected.side === side && selected.id === id ? null : { side, id };
        render();
    }
    // A matched pair's middle: ✓, and ↶ Undo on hover (also used to turn a pair to check into a matched one in place)
    function setDone(row, u) {
        const w = packs[pairs[u.key]];
        row.classList.remove('sm-pair-suggested'); row.classList.add('sm-pair-done');
        const mid = row.querySelector('.sm-pair-mid');
        mid.replaceChildren(el('span', 'sm-done-tick', '✓'));
        const b = el('button', 'sm-undo', '↶ Undo'); b.type = 'button'; b.title = 'Undo this match';
        b.setAttribute('aria-label', `Undo the match of ${label(u)} with ${packName(w)}`);
        b.addEventListener('click', () => undoPair(u.key));
        row.append(b);
    }
    function pairRow(u, toCheck) {
        const w = packs[pairs[u.key]];
        const row = el('div', 'sm-pair sm-pair-suggested');
        row.dataset.key = u.key;
        const mid = el('div', 'sm-pair-mid');
        mid.append(el('span', 'sm-suggested-tag', `suggested${u.score ? ' ' + pct(u.score) : ''}`));
        const tick = el('button', 'sm-tick', '✓'); tick.type = 'button'; tick.title = 'Right pair: match them';
        tick.setAttribute('aria-label', `Match ${label(u)} with ${packName(w)}`);
        tick.addEventListener('click', () => confirmRow(u.key, tick));
        const x = el('button', 'sm-x', '✕'); x.type = 'button'; x.title = 'Wrong pair: split it';
        x.setAttribute('aria-label', `Split ${label(u)} from ${packName(w)}`);
        x.addEventListener('click', () => unpair(u.key));
        mid.append(tick, x);
        row.append(courierCard(u), mid, packCard(w));
        if (!toCheck) setDone(row, u);
        return row;
    }
    // Still to match: a courier label beside its closest free packing label, ✓ to match them, "Other…" to pick another
    function freeRow(u, ref, free) {
        const row = el('div', `sm-pair sm-pair-free${ref ? '' : ' sm-pair-none'}`);
        row.dataset.key = u.key;
        const mid = el('div', 'sm-pair-mid');
        const pick = el('select', 'sm-pick');
        pick.setAttribute('aria-label', `Packing label for ${label(u)}`);
        pick.append(new Option(ref ? 'Other…' : 'Pick…', ''));
        const near = (u.candidates || []).filter((c) => free.has(c[0]) && c[0] !== ref);
        const rest = [...free].filter((r) => r !== ref && !near.some((c) => c[0] === r))
            .sort((a, b) => (packs[a].page || 0) - (packs[b].page || 0));
        if (near.length) {
            const g = el('optgroup'); g.label = 'Closest addresses';
            near.forEach((c) => g.append(new Option(`${packName(packs[c[0]])} — ${pct(c[1])}`, c[0])));
            pick.append(g);
        }
        if (rest.length) {
            const g = el('optgroup'); g.label = 'All other packing labels';
            rest.forEach((r) => g.append(new Option(`${packName(packs[r])} — ${packs[r].receiver}`, r)));
            pick.append(g);
        }
        pick.addEventListener('change', () => { if (pick.value) { choice[u.key] = pick.value; render(); } });
        if (ref) {
            const sc = scoreOf(u, ref);
            mid.append(el('span', 'sm-closest-tag', choice[u.key] === ref ? 'picked' : `closest${sc ? ' ' + pct(sc) : ''}`));
            const tick = el('button', 'sm-tick', '✓'); tick.type = 'button'; tick.title = 'Match these two';
            tick.setAttribute('aria-label', `Match ${label(u)} with ${packName(packs[ref])}`);
            tick.addEventListener('click', () => pair(u.key, ref, tick));
            mid.append(tick);
        } else {
            mid.append(el('span', 'sm-closest-tag', 'no close match'));
        }
        mid.append(pick);
        const right = ref ? packCard(packs[ref]) : el('div', 'sm-card sm-none-box', 'No packing label at a matching address. Pick one, or drag this courier label onto one below.');
        row.append(courierCard(u), mid, right);
        return row;
    }

    // ---------- the page ----------
    function render() {
        clearTimeout(renderTimer);
        const focusKey = document.activeElement && document.activeElement.classList.contains('sm-tick')
            ? (document.activeElement.closest('.sm-pair') || {}).dataset?.key : null;
        const words = search.value.trim().toLowerCase().split(/\s+/).filter(Boolean);
        const hit = (text) => words.every((w) => text.includes(w));
        const rowHit = (row) => [...row.querySelectorAll('.sm-card[data-search]')].some((c) => hit(c.dataset.search));
        // Pairs, grouped by the packing label's receiver and address, side by side: to check, or matched
        pairList.replaceChildren();
        const groups = new Map();
        session.unmatched.forEach((u) => {
            if (!packs[pairs[u.key]]) return;
            const g = placeKey(packs[pairs[u.key]]);
            if (!groups.has(g)) groups.set(g, []);
            groups.get(g).push(u);
        });
        groups.forEach((list, g) => {
            const first = packs[pairs[list[0].key]];
            const waitingHere = (session.waiting || []).filter((w) => placeKey(w) === g).length;
            const checking = list.filter((u) => suggested[u.key]).length;
            const box = el('div', 'sm-group sm-pair-group');
            const head = el('div', 'sm-group-head');
            head.append(el('strong', null, g === 'none' ? 'Not in the courier CSV' : first.receiver),
                        el('span', 'ab-detail', `${g === 'none' ? '' : ' ' + first.address + ' · '}${plural(list.length, 'courier label')} ↔ ${plural(waitingHere, 'packing label')}${checking ? ` · ${checking} to check` : ' · ✓ all matched'}`));
            box.append(head);
            let shown = 0;
            list.sort((a, b) => (packs[pairs[a.key]].page || 0) - (packs[pairs[b.key]].page || 0)).forEach((u) => {
                const row = pairRow(u, !!suggested[u.key]);
                row.hidden = !rowHit(row);
                shown += !row.hidden;
                box.append(row);
            });
            box.hidden = !shown;
            box.classList.toggle('sm-group-done', !checking);
            pairList.append(box);
        });
        if (!groups.size) pairList.append(el('p', 'ps-hint sm-empty', finished ? '✅ Saved and stitched — see the result above. "Match the rest" opens what is still unmatched.'
                                                                              : 'No pairs yet: match the courier labels below with ✓, or drag one onto its packing label.'));
        // Still to match: every free courier label with its closest free packing label (a picked one first)
        freeList.replaceChildren();
        const free = new Set((session.waiting || []).map((w) => w.item_reference).filter((r) => !courierOf(r)));
        const freeCouriers = session.unmatched.filter((u) => !pairs[u.key]);
        const shownWith = {};
        const claimed = new Set();
        freeCouriers.forEach((u) => { if (choice[u.key] && free.has(choice[u.key]) && !claimed.has(choice[u.key])) { shownWith[u.key] = choice[u.key]; claimed.add(choice[u.key]); } });
        freeCouriers.forEach((u) => {
            if (shownWith[u.key]) return;
            const c = (u.candidates || []).find((x) => free.has(x[0]) && !claimed.has(x[0]));
            if (c) { shownWith[u.key] = c[0]; claimed.add(c[0]); }
        });
        freeCouriers.forEach((u) => {
            const row = freeRow(u, shownWith[u.key] || null, free);
            row.hidden = !rowHit(row);
            freeList.append(row);
        });
        // Packing labels nobody's closest to: still a drop target, or pick them with "Other…"
        packList.replaceChildren();
        const waitGroups = new Map();
        (session.waiting || []).filter((w) => free.has(w.item_reference) && !claimed.has(w.item_reference)).forEach((w) => {
            const g = placeKey(w);
            if (!waitGroups.has(g)) waitGroups.set(g, []);
            waitGroups.get(g).push(w);
        });
        waitGroups.forEach((list, g) => {
            const box = el('div', 'sm-group');
            const head = el('div', 'sm-group-head');
            head.append(el('strong', null, g === 'none' ? 'Not in the courier CSV' : list[0].receiver),
                        el('span', 'ab-detail', g === 'none' ? '' : ` ${list[0].address} · ${list.length} waiting`));
            box.append(head);
            let shown = 0;
            list.forEach((w) => { const card = packCard(w); card.hidden = !hit(card.dataset.search); shown += !card.hidden; box.append(card); });
            box.hidden = !shown;
            packList.append(box);
        });
        document.getElementById('sm-others').hidden = !waitGroups.size;
        document.getElementById('sm-right-count').textContent = `(${free.size - claimed.size})`;
        hintbar.hidden = !words.length;
        updateCounts();
        if (focusKey) {
            const t = pairList.querySelector(`.sm-pair[data-key="${CSS.escape(focusKey)}"] .sm-tick`) || freeList.querySelector(`.sm-pair[data-key="${CSS.escape(focusKey)}"] .sm-tick`);
            if (t) t.focus({ preventScroll: true });
        }
    }
    // Counts and buttons, without re-drawing the cards
    function updateCounts() {
        const toCheck = Object.keys(suggested).length;
        const n = Object.keys(pairs).length;
        const restCouriers = session.unmatched.length - n, restPacks = (session.waiting || []).length - n;
        document.getElementById('sm-paired-count').textContent =
            `(${n - toCheck} of ${plural(session.unmatched.length, 'courier label')} matched${toCheck ? `, ${toCheck} to check` : ''})`;
        document.getElementById('sm-left-count').textContent = `(${plural(restCouriers, 'courier label')}, ${plural(restPacks, 'packing label')})`;
        document.getElementById('sm-rest').hidden = !restCouriers && !restPacks;
        const shownCouriers = visibleRows(freeList).length;
        const shownPacks = [...document.querySelectorAll('#sm-rest .sm-pack')].filter((c) => !c.hidden && !c.closest('[hidden]')).length;
        hintText.textContent = `Shown still to match: ${plural(shownCouriers, 'courier label')} and ${plural(shownPacks, 'packing label')}.`;
        const matchShown = document.getElementById('sm-match-shown');
        matchShown.disabled = !shownCouriers || shownCouriers !== shownPacks;
        matchShown.textContent = shownCouriers === shownPacks ? `Match these ${shownPacks} in order` : 'Counts differ — narrow the search';
        const sures = sureLeft().length;
        document.querySelectorAll('.sm-sure-btn').forEach((b) => {
            b.disabled = !sures || root.dataset.running === '1';
            b.textContent = `✓ Match all 100% (${sures})`;
            b.title = sures ? `Match the ${plural(sures, 'pair')} whose courier label address matched the packing label 100%`
                            : 'No pair to check matched 100% — check the amber ones with ✓ or ✕';
        });
        saveBtn.disabled = finished || !session.unmatched.length || root.dataset.running === '1';
        saveBtn.textContent = saved ? '💾 Saved' : `💾 Save changes (${plural(n, 'pair')})`;
        finaliseBtn.disabled = finished || root.dataset.running === '1';
        finaliseBtn.textContent = finished ? '✅ Finalised' : '✅ Finalise';
    }

    // Search shows the same number of courier and packing labels: pair them in order (courier labels in file order,
    // packing labels in packing label order)
    document.getElementById('sm-match-shown').addEventListener('click', () => {
        const couriers = visibleRows(freeList).map((r) => r.dataset.key);
        const refs = [...document.querySelectorAll('#sm-rest .sm-pack')].filter((c) => !c.hidden && !c.closest('[hidden]'))
            .map((c) => c.dataset.ref).sort((a, b) => (packs[a].page || 0) - (packs[b].page || 0));
        if (!couriers.length || couriers.length !== refs.length) return;
        couriers.forEach((k, i) => { pairs[k] = refs[i]; delete suggested[k]; delete choice[k]; fromSuggestion[k] = byKey[k].suggestion === refs[i]; });
        saved = false; render();
        say(`${plural(refs.length, 'pair')} matched in order: the first courier label shown with the first packing label, and so on.`);
    });
    const matchAllSure = () => {
        const list = sureLeft();
        if (!list.length) return;
        list.forEach((u) => { pairs[u.key] = u.suggestion; delete suggested[u.key]; fromSuggestion[u.key] = true; });
        saved = false; render();
        say(`${plural(list.length, 'pair')} matched 100%.${Object.keys(suggested).length ? ' The amber ones left need a look.' : ''}`);
        const first = visibleRows(pairList).find(toCheckRow);
        if (first) { first.scrollIntoView({ behavior: motion ? 'smooth' : 'instant', block: 'center' }); bringUnder(first, null); }
    };
    document.querySelectorAll('.sm-sure-btn').forEach((b) => b.addEventListener('click', matchAllSure));
    document.getElementById('sm-restore').addEventListener('click', () => { restoreSuggestions(); render(); say('Suggested pairs put back where their labels were free.'); });
    document.getElementById('sm-clear').addEventListener('click', () => {
        Object.keys(pairs).forEach((k) => { delete pairs[k]; delete suggested[k]; delete fromSuggestion[k]; }); selected = null; saved = true; render();
        say('Every pair split.');
    });
    search.addEventListener('input', () => render());
    search.addEventListener('keydown', (e) => { if (e.key === 'Escape') { search.value = ''; render(); } });
    zoom.querySelector('.sm-zoom-close').addEventListener('click', () => zoom.close());
    zoom.addEventListener('click', (e) => { if (e.target === zoom) zoom.close(); });

    // Courier labels still matching nothing after this save: ask whether to add them 4-up at the end
    const ask = document.getElementById('sm-ask');
    const askUser = (left) => new Promise((resolve) => {
        document.getElementById('sm-ask-text').textContent =
            `${plural(left, 'courier label')} will still match no packing label. Add ${left === 1 ? 'it' : 'them'} 4-up at the end of the Complete Labels PDF (in cut-and-stack order), or leave ${left === 1 ? 'it' : 'them'} out?`;
        const done = (answer) => { ask.close(); resolve(answer); };
        ask.querySelectorAll('[data-answer]').forEach((b) => { b.onclick = () => done(b.dataset.answer); });
        ask.oncancel = (e) => { e.preventDefault(); done('cancel'); };
        ask.showModal();
    });
    // Save changes: the pairs are kept (stitch_matches.json); nothing is stitched until Finalise
    const post = async (url, body) => {
        const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(data.error || 'Could not save.');
        return data;
    };
    saveBtn.addEventListener('click', async () => {
        if (root.dataset.running === '1') return;
        saveBtn.disabled = true;
        try {
            const data = await post(root.dataset.saveUrl, { matches: pairs });
            saved = true;
            say(`💾 Saved ${plural(data.saved, 'match')}. ` + (data.complete
                ? 'Every courier label now has its packing label: press Finalise to stitch the Complete Labels.'
                : `Still ${plural(data.unmatched, 'courier label')} without a packing label. Nothing is stitched until you press Finalise.`));
        } catch (err) { say(err.message, true); }
        render();
    });
    // Finalise: stitch once with these pairs, in the background behind the overlay
    finaliseBtn.addEventListener('click', async () => {
        if (root.dataset.running === '1') return;   // one click only
        const left = session.unmatched.length - Object.keys(pairs).length;
        let addUnmatched = true;
        if (left > 0) {
            const answer = await askUser(left);
            if (answer === 'cancel') return;
            addUnmatched = answer === 'add';
        }
        if (root.dataset.running === '1') return;
        root.dataset.running = '1'; render();
        busy('Stitching the Complete Labels…', 'Rendering each courier label sharp enough to scan - a large job can take a minute.');
        say('Stitching the Complete Labels…');
        try {
            await post(root.dataset.finaliseUrl, { matches: pairs, add_unmatched: addUnmatched });
            saved = true;
            poll();
        } catch (err) { root.dataset.running = ''; idle(); render(); say(err.message, true); }
    });
    // The finished file: Download and Open file, nothing else to do on this page
    function showResult(data) {
        const link = root.dataset.downloadUrl.replace('__FILE__', encodeURIComponent(data.output));
        status.replaceChildren(el('span', null, `✅ ${data.output}: ${plural(data.placed, 'courier label')} placed`
            + (data.unmatched ? `, ${data.unmatched} unmatched${data.added_unmatched ? ' (4-up at the end)' : ' (left out)'}. ` : '. ')));
        const a = el('a', 'ab-btn', '⬇ Download'); a.href = link;
        const open = el('a', 'ab-btn ab-btn-plain btn-file open', '📂 Open file');
        open.href = root.dataset.openUrl.replace('__FILE__', encodeURIComponent(data.output));
        status.append(a, ' ', open);
        if ((data.locked || []).length) status.append(el('div', 'ps-hint', `Still open in another program, so not removed: ${data.locked.join(', ')}.`));
        status.classList.remove('sm-status-bad');
        Object.keys(pairs).forEach((k) => { delete pairs[k]; delete suggested[k]; delete fromSuggestion[k]; });
        session.unmatched = []; session.waiting = []; finished = true;
        render();
        status.scrollIntoView({ behavior: motion ? 'smooth' : 'instant', block: 'center' });
    }
    async function poll() {
        try {
            // The server answers when the stitch is done (or after 25 s): one request at a time, not one a second
            const data = await (await fetch(`${root.dataset.statusUrl}?wait=1`)).json();
            if (data.state === 'running') { setTimeout(poll, 100); return; }
            root.dataset.running = '';
            idle();
            if (data.state === 'error') { render(); say(data.error || 'Stitching failed.', true); return; }
            showResult(data);
        } catch (e) { setTimeout(poll, 2500); }
    }
    // More courier label PDFs: the earlier Complete Labels go, the matching is worked out again with them; when
    // everything matches it's stitched straight away, else the page opens again with what's left
    document.getElementById('sm-add-files').addEventListener('change', async (e) => {
        const files = [...e.target.files];
        if (!files.length) return;
        if (!saved && Object.keys(pairs).length && !confirm("Matches you haven't saved will be lost. Add the PDFs anyway?")) { e.target.value = ''; return; }
        const form = new FormData();
        files.forEach((f) => form.append('labels', f));
        root.dataset.running = '1'; render();
        busy(`Adding ${plural(files.length, 'PDF')} and matching again…`, 'Every courier label is read again with the new ones.');
        try {
            const res = await fetch(root.dataset.addUrl, { method: 'POST', body: form });
            const data = await res.json().catch(() => ({}));
            if (!res.ok) throw new Error(data.error || 'Could not add the PDFs.');
            saved = true;
            if (data.complete) { busy('Everything matches now: stitching the Complete Labels…'); poll(); }
            else window.location.reload();  // what's left to match
        } catch (err) { root.dataset.running = ''; idle(); render(); say(err.message, true); }
        e.target.value = '';
    });
    window.addEventListener('beforeunload', (e) => { if (!saved && Object.keys(pairs).length) { e.preventDefault(); e.returnValue = ''; } });
    // Matches saved earlier come back as matched pairs; the rest start as suggestions to check
    session.unmatched.forEach((u) => {
        if (u.matched_to && packs[u.matched_to] && !courierOf(u.matched_to)) {
            pairs[u.key] = u.matched_to;
            if (u.suggestion === u.matched_to) fromSuggestion[u.key] = true;
        }
    });
    const savedBefore = Object.keys(pairs).length;
    restoreSuggestions();
    saved = Object.keys(pairs).length === savedBefore;
    render();
    if (!session.unmatched.length) say('Every courier label is matched. Nothing to do here.');
    else if (Object.keys(pairs).length) {
        const fresh = Object.keys(pairs).length - savedBefore;
        say([savedBefore ? `${plural(savedBefore, 'match')} saved earlier.` : '',
             fresh ? `${plural(fresh, 'pair')} suggested from the addresses on the courier labels — ✓ each one (or Match all 100%).` : '',
             'Then Finalise to stitch the Complete Labels.'].filter(Boolean).join(' '));
    }
});

// =========================================
// 17. STITCH LABELS FORM: one click only
// =========================================
// The first click disables the button and shows the loading overlay (the global submit handler, section 8);
// a second click can't send the PDFs again. Nothing chosen: say so instead of submitting.
document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('st-form');
    if (!form) return;
    const button = document.getElementById('st-submit');
    const error = document.getElementById('st-form-error');
    form.addEventListener('submit', (e) => {
        const files = form.querySelector('.dropzone-input');
        const ticked = form.querySelectorAll('input[name="existing"]:checked').length;
        if (form.dataset.sent === '1' || (!(files && files.files.length) && !ticked)) {
            e.preventDefault();
            e.stopImmediatePropagation();  // no overlay for a form that isn't sent
            if (form.dataset.sent !== '1') error.hidden = false;
            return;
        }
        error.hidden = true;
        form.dataset.sent = '1';
        button.disabled = true;
        button.textContent = '🧵 Stitching…';
    }, true);  // capture: runs before the overlay handler on <body>
    // Coming back with the browser's Back button: ready to use again
    window.addEventListener('pageshow', () => {
        form.dataset.sent = '';
        button.disabled = false;
        button.textContent = '🧵 Stitch Labels';
        const overlay = document.getElementById('loading-overlay');
        if (overlay) overlay.style.display = '';
    });
});

// =========================================
// 18. COLUMN MAPPING AND SIMILAR ADDRESSES (packing labels preview)
// =========================================
// Column mapping: each field -> the distro header it's matched to -> that header's cell. Picking another column (or
// "Not used") shows its cell straight away; Update Previews reads the file with it. What's picked is remembered in
// this browser for the file and tab (localStorage), so it comes back after a refresh or when the same file is
// scanned again: it's sent with the first preview of the file even before the mapping is on the page.
// Packing Labels and Courier Import keep theirs apart (data-colmap-key): their fields differ.
document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('pl-form');
    if (!form) return;
    const KEY = form.dataset.colmapKey || 'pl_column_mapping';
    const file = (form.querySelector('input[name="filename"]') || {}).value || '';
    const load = () => { try { return JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { return {}; } };
    const store = (all) => { try { localStorage.setItem(KEY, JSON.stringify(all)); } catch (e) { /* storage blocked */ } };

    document.querySelectorAll('.pl-colmap').forEach((panel) => {
        const row = panel.dataset.headerRow;
        const letterOf = (select) => (select.value === '-' ? '' : (select.value || select.dataset.auto));
        const headerOf = (select) => {
            const letter = letterOf(select);
            if (!letter) return '';
            if (!select.value) return select.dataset.autoText;
            return (select.selectedOptions[0].text || '').replace(/ \([A-Z]+\)$/, '');
        };
        // A row of a field still to map: its cell follows the pick
        const showRow = (select) => {
            const item = select.closest('.pl-colmap-item');
            const letter = letterOf(select);
            item.querySelector('.pl-colmap-cell').textContent = letter ? `${letter}${row}` : '—';
            item.classList.toggle('pl-colmap-hand', !!select.value);
        };
        // A chip of a mapped field: green (found by the rules), blue (set by hand) or grey (not used)
        const showChip = (select) => {
            const item = select.closest('.pl-colmap-item');
            const chip = item.querySelector('.pl-chip');
            const letter = letterOf(select);
            const state = select.value === '-' ? 'off' : (select.value ? 'hand' : 'auto');
            item.classList.remove('pl-chip-auto', 'pl-chip-hand', 'pl-chip-off', 'pl-chip-none');
            item.classList.add(letter || state === 'off' ? `pl-chip-${state}` : 'pl-chip-none');
            chip.querySelector('.pl-chip-mark').textContent = !letter && state !== 'off' ? '!' : { auto: '✓', hand: '✎', off: '⊘' }[state];
            chip.querySelector('.pl-chip-cell').textContent = state === 'off' ? 'not used' : (letter ? `${letter}${row}` : 'not found');
            chip.title = letter ? `Excel header: “${headerOf(select)}” · cell ${letter}${row} — click to change` : 'No column — click to pick one';
        };
        panel.querySelectorAll('.pl-row-item .pl-colmap-select').forEach((select) => select.addEventListener('change', () => showRow(select)));
        panel.querySelectorAll('.pl-chip-item').forEach((item) => {
            const chip = item.querySelector('.pl-chip');
            const select = item.querySelector('.pl-colmap-select');
            const close = () => { select.hidden = true; chip.hidden = false; };
            chip.addEventListener('click', () => { chip.hidden = true; select.hidden = false; select.focus(); });
            select.addEventListener('change', () => { showChip(select); close(); chip.focus(); });
            select.addEventListener('blur', close);
            select.addEventListener('keydown', (e) => { if (e.key === 'Escape') { close(); chip.focus(); } });
        });
        panel.querySelector('.pl-colmap-reset').addEventListener('click', () => {
            panel.querySelectorAll('.pl-colmap-select').forEach((select) => {
                select.value = '';
                if (select.closest('.pl-chip-item')) showChip(select); else showRow(select);
            });
        });
    });

    form.addEventListener('submit', () => {
        if (!file) return;
        const all = load();
        const mine = all[file] || {};
        const onPage = new Set();
        form.querySelectorAll('.pl-colmap').forEach((panel) => {
            const tab = panel.dataset.tab;
            onPage.add(tab);
            mine[tab] = {};
            panel.querySelectorAll('.pl-colmap-select').forEach((select) => {
                if (select.value) mine[tab][select.name.split('::')[2]] = select.value;
            });
            if (!Object.keys(mine[tab]).length) delete mine[tab];
        });
        // Tabs without the mapping on the page yet (the first preview of a file): send what was picked last time
        form.querySelectorAll('input[name="all_tabs"]').forEach((input) => {
            const tab = input.value;
            if (onPage.has(tab) || !mine[tab]) return;
            Object.entries(mine[tab]).forEach(([field, value]) => {
                const hidden = document.createElement('input');
                hidden.type = 'hidden'; hidden.name = `colmap::${tab}::${field}`; hidden.value = value;
                form.appendChild(hidden);
            });
        });
        if (Object.keys(mine).length) all[file] = mine; else delete all[file];
        store(all);
    });

    // One receiver's consignments at different addresses sit next to each other under a heading: the address picked
    // there goes on all of them (as an edit, like ✏️) and the preview is updated at once, so they become one
    const editsField = document.getElementById('consignment-edits');
    const preview = form.querySelector('button[name="preview"]');
    document.querySelectorAll('.pl-group-merge').forEach((button) => {
        button.addEventListener('click', () => {
            if (!editsField || !preview) return;
            const group = JSON.parse(button.dataset.group);
            const picked = button.closest('.pl-group-head').querySelector('.pl-group-keep');
            const keep = group[picked ? Number(picked.value) : 0];
            // They become one consignment, so one service: the kept address's. Say so when theirs differed.
            const services = [...new Set(group.map((m) => m.service_used).filter(Boolean))];
            if (services.length > 1 && !confirm(`These consignments use different services: ${group.map((m) => `#${m.number} ${m.service_used}`).join(', ')}.\n\n`
                + `Merged, they all go with the service of the address you keep (#${keep.number}: ${keep.service_used}). Merge?`)) return;
            let edits = {};
            try { edits = JSON.parse(editsField.value || '{}') || {}; } catch (e) { edits = {}; }
            group.forEach((member) => member.sources.forEach((id) => {
                edits[id] = { ...(edits[id] || {}), checked: true, ...keep.fields };
                if (keep.service) edits[id].service_code = keep.service; else delete edits[id].service_code;  // '' = same as main
            }));
            editsField.value = JSON.stringify(edits);
            button.disabled = true;
            button.textContent = 'Merging…';
            form.requestSubmit(preview);
        });
    });
});

// =========================================
// 19. COURIER CSV CHOICE (packing labels preview)
// =========================================
// Open360, OpenFreight or both, just above the project name. Nothing is ticked: it's chosen for every job, and
// Generate needs one (a message says so, like section 14's required boxes; the server refuses too).
document.addEventListener('DOMContentLoaded', function () {
    const box = document.getElementById('pl-csv-choice');
    if (!box) return;
    const ticks = [...box.querySelectorAll('input[name="courier_csv"]')];
    ticks.forEach((t) => t.addEventListener('change', () => box.classList.remove('pl-missing')));
    const form = box.closest('form');
    form.addEventListener('submit', (e) => {
        if (!e.submitter || e.submitter.name !== 'generate' || ticks.some((t) => t.checked)) return;
        e.preventDefault();
        e.stopImmediatePropagation();
        box.classList.add('pl-missing');
        const error = document.getElementById('pl-generate-error');
        if (error) { error.textContent = '⚠️ Choose the courier CSV to make: Open360, OpenFreight or both.'; error.hidden = false; }
        box.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }, true);
});

// =========================================
// 20. PROJECT NAME LENGTH
// =========================================
// Project names are at most 50 characters (they go into the folder and every file name, and Excel can't open a
// file whose full path is too long). The box shows what's left as you type.
document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('input[data-count][maxlength]').forEach((box) => {
        const max = Number(box.maxLength);
        const note = document.createElement('span');
        note.className = 'name-count';
        note.setAttribute('aria-live', 'polite');
        box.insertAdjacentElement('afterend', note);
        const show = () => {
            const left = max - box.value.length;
            note.textContent = `${box.value.length} / ${max} characters`;
            note.classList.toggle('name-count-full', left <= 0);
        };
        box.addEventListener('input', show);
        show();
    });
});

// =========================================
// 21. EXPORT DETAILS (packing labels preview)
// =========================================
// For deliveries outside Australia in the OpenFreight CSV: shown under the Courier CSV buttons only while OpenFreight
// is chosen and the job has consignments outside Australia (counted again when a ✏️ edit changes a country, so
// it appears as soon as an address is made overseas, and goes when none are left).
// The values start as the saved defaults and can be changed for this job; "Save as default" keeps them for future
// jobs (/api/export-defaults) and this job uses them too; "Reset to default" puts the saved ones back.
document.addEventListener('DOMContentLoaded', function () {
    const panel = document.getElementById('pl-export');
    if (!panel) return;
    const openfreight = document.querySelector('input[name="courier_csv"][value="openfreight"]');
    const fields = [...panel.querySelectorAll('[data-export]')];
    const status = document.getElementById('pl-export-status');
    const count = document.getElementById('pl-export-count');
    const show = () => {
        const n = [...document.querySelectorAll('.pl-con-row')].filter((r) => (r.dataset.country || 'AU') !== 'AU').length;
        panel.hidden = !(openfreight && openfreight.checked && n > 0);
        if (count) count.textContent = `${n} consignment${n === 1 ? '' : 's'}`;
    };
    if (openfreight) openfreight.addEventListener('change', show);
    document.addEventListener('pl-countries-changed', show);
    show();
    const read = () => Object.fromEntries(fields.map((f) => [f.dataset.export, f.value]));
    const fill = (values) => fields.forEach((f) => { if (values[f.dataset.export] !== undefined) f.value = values[f.dataset.export]; });
    document.getElementById('pl-export-save').addEventListener('click', async () => {
        try {
            const res = await fetch(panel.dataset.defaultsUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(read()) });
            const data = await res.json();
            fill(data.saved);
            // The fields keep what was saved, so this job's Generate uses the same values
            status.textContent = data.problems && data.problems.length
                ? `Saved as the default, and used for this job (not a number, kept as typed: ${data.problems.join(', ')}).`
                : '✓ Saved as the default, and used for this job.';
        } catch (e) { status.textContent = 'Could not save the defaults.'; }
    });
    document.getElementById('pl-export-reset').addEventListener('click', async () => {
        try {
            fill(await (await fetch(panel.dataset.defaultsUrl)).json());
            status.textContent = 'Back to the saved defaults.';
        } catch (e) { status.textContent = 'Could not read the defaults.'; }
    });
});
