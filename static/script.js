// A task behind the loading overlay that failed comes back with X-Log-Url (app.py _failed): the page goes to the Logs
// page at what went wrong, as a form's failure does
(() => {
    const fetchFromPage = window.fetch.bind(window);
    window.fetch = async (...args) => {
        const res = await fetchFromPage(...args);
        const to = res.headers.get('X-Log-Url');
        if (to && to.startsWith('/logs')) window.location.assign(to);
        return res;
    };
})();

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
    // Back from 'Create … CSV' (?open=<project>&new=<file>): that project open, at its courier CSVs, the new file marked
    const params = new URLSearchParams(location.search);
    if (params.get('open')) {
        const content = document.getElementById('content-' + params.get('open'));
        if (content) {
            content.style.display = 'block';
            const icon = content.previousElementSibling.querySelector('.accordion-icon');
            if (icon) { icon.innerHTML = '▲'; icon.style.color = '#3498db'; }
            const target = content.querySelector('.file-row-new') || document.getElementById('csv-' + params.get('open')) || content;
            requestAnimationFrame(() => target.scrollIntoView({ block: 'center', behavior: 'instant' }));
        }
        history.replaceState(null, '', location.pathname + '#dashboard-section');  // a refresh doesn't jump again
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
    // Sub-Group page (sub-group.html): the blocks it builds, by tab and pack name, and two helpers that only ever set
    // text and properties (sgEl), so names from the client's file can't become HTML
    const sgTabs = new Map();  // tab name -> {block, packs (its packs' container), packBlocks: pack name -> block}
    function sgEl(tag, props, text) {
        const n = Object.assign(document.createElement(tag), props || {});
        if (props && props.style) n.style.cssText = props.style;
        if (text !== undefined) n.textContent = text;
        return n;
    }
    function sgRow(tabName, packName, removable) {
        const row = sgEl('div', { className: 'input-group subgroup-row', style: 'margin-bottom: 15px; align-items: flex-end;' });
        for (const [label, field] of [['Start Item #', 'start_item'], ['End Item #', 'end_item']]) {
            const cell = sgEl('div', { className: 'input-field' });
            cell.append(sgEl('label', {}, label), sgEl('input', { type: 'number', name: `${field}_${tabName}_${packName}[]`, required: true }));
            row.append(cell);
        }
        if (removable) {
            const remove = sgEl('button', { type: 'button', className: 'btn-file delete', style: 'height: 40px; min-width: 80px;' }, 'Remove');
            remove.addEventListener('click', () => row.remove());
            row.append(remove);
        }
        return row;
    }

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

        // Tab Selector: a block per ticked tab, with its packs to pick. Tab and pack names come from the client's Excel
        // file, so every one goes in as text or a property, never as HTML: a pack called 12" Poster keeps its quote in
        // the field names app.py reads back (start_item_<tab>_<pack>[]), and nothing in a name can run as script
        if (e.target.matches('.tab-checkbox-grid input[type="checkbox"]:not([name^="target_pack"])')) {
            const tabName = e.target.value;
            const container = document.getElementById('tab-blocks-container');
            const submitBtn = document.getElementById('submit-btn');

            if (e.target.checked && meta) {
                selectedTabsState.add(tabName);
                const block = sgEl('div', { className: 'tab-card', style: 'border-left: 5px solid #3498db;' });
                const title = sgEl('h3', { className: 'tab-title' }, '⚙️ Setup for Tab: ');
                title.append(sgEl('span', { style: 'color: #3498db;' }, tabName));
                const rowField = sgEl('div', { className: 'input-field', style: 'width: 50%; min-width: 200px; margin-bottom: 20px;' });
                rowField.append(sgEl('label', {}, 'Item Number Row (e.g. 9)'),
                    sgEl('input', { type: 'number', name: `item_row_${tabName}`, required: true,
                                   style: 'width: 100%; padding: 10px; border: 1px solid #bdc3c7; border-radius: 5px; box-sizing: border-box;' }));
                const grid = sgEl('div', { className: 'tab-checkbox-grid', style: 'margin-top: 5px;' });
                for (const [packName, packData] of Object.entries(meta.tabs[tabName].packs)) {
                    if (!packData.is_selected) continue;
                    const label = sgEl('label', { className: 'tab-checkbox-label' });
                    const box = sgEl('input', { type: 'checkbox', name: `target_pack_${tabName}[]`, value: packName });
                    box.dataset.tab = tabName;
                    box.dataset.pack = packName;
                    label.append(box, ' ', document.createTextNode(packName));
                    grid.append(label);
                }
                const packField = sgEl('div', { className: 'input-field', style: 'margin-bottom: 20px;' });
                packField.append(sgEl('label', {}, 'Parent Pack(s) to Sub-Group'), grid);
                const packs = sgEl('div');
                block.append(title, sgEl('input', { type: 'hidden', name: 'selected_tabs[]', value: tabName }), rowField, packField, packs);
                sgTabs.set(tabName, { block, packs, packBlocks: new Map() });
                container.append(block);
            } else {
                selectedTabsState.delete(tabName);
                const tab = sgTabs.get(tabName);
                if (tab) { tab.block.remove(); sgTabs.delete(tabName); }
            }
            if (submitBtn) submitBtn.style.display = selectedTabsState.size > 0 ? 'block' : 'none';
        }

        // Pack Selector: a block per ticked pack, with its sub-group rows
        if (e.target.matches('input[name^="target_pack_"]')) {
            const tab = sgTabs.get(e.target.dataset.tab);
            const packName = e.target.dataset.pack;
            if (!tab) return;
            if (e.target.checked) {
                const block = sgEl('div', { className: 'blueprint-card', style: 'margin-top: 15px;' });
                const head = sgEl('h4', { style: 'color: #2c3e50; font-size: 15px; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-bottom: 15px;' },
                                  'Define Sub-Groups for: ');
                head.append(sgEl('span', { style: 'color: #e67e22;' }, packName));
                const rows = sgEl('div');
                rows.append(sgRow(e.target.dataset.tab, packName, false));
                const add = sgEl('button', { type: 'button', className: 'btn-file open', style: 'margin-top: 5px;' },
                                 `+ Add Another Sub-Group for ${packName}`);
                add.addEventListener('click', () => rows.append(sgRow(e.target.dataset.tab, packName, true)));
                block.append(head, rows, add);
                tab.packBlocks.set(packName, block);
                tab.packs.append(block);
            } else {
                const block = tab.packBlocks.get(packName);
                if (block) { block.remove(); tab.packBlocks.delete(packName); }
            }
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

        const jobField = document.getElementById('job-number-field');
        const jobInput = document.getElementById('job_number_input');
        const JOB = /^[A-Za-z]{1,3}\d{4,8}(-\d{1,4})?$/;  // as app.py checks it (JOB_NUMBER)

        function validateForm() {
            let isValid = false;
            // A new project asks for its job number (the dashboard shows it)
            if (jobField) jobField.hidden = !(textInput && textInput.value.trim() !== "");

            if (textInput && textInput.value.trim() !== "") {
                if (dropdown) {
                    dropdown.value = "";
                    dropdown.disabled = true;
                    dropdown.style.backgroundColor = "#e9ecef";
                }
                if (helperText) helperText.innerHTML = "Required: Please upload the signature links Excel file for your new project.";
                if (fileInput && fileInput.files.length > 0) isValid = !jobInput || JOB.test(jobInput.value.replace(/\s+/g, ''));
            } 
            else if (dropdown && dropdown.value !== "") {
                if (textInput) {
                    textInput.value = "";
                    textInput.disabled = true;
                    textInput.style.backgroundColor = "#e9ecef";
                }
                if (projectData[dropdown.value]) {
                    if (helperText) {  // the file's name goes in as text
                        const name = document.createElement('strong');
                        name.textContent = projectData[dropdown.value];
                        helperText.replaceChildren('📂 ', Object.assign(document.createElement('strong'), { textContent: 'Found:' }),
                            " Upload another file otherwise system will use '", name, "'.");
                    }
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
        if (jobInput) jobInput.addEventListener('input', validateForm);
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
            const text = document.getElementById('loading-text');
            if (text && !text.dataset.idle) text.dataset.idle = text.textContent;
            if (text && e.target.action && e.target.action.includes('/courier-csv')) text.textContent = 'Making the courier CSV…';
            // A form in a pop-up (Generate Labels from …): the pop-up closes, or it would sit over the loading screen
            // and its button could be pressed again
            const dialog = e.target.closest('dialog');
            if (dialog && dialog.open && !e.defaultPrevented) dialog.close();
            overlay.style.display = 'flex';
            if (!e.defaultPrevented) followProgress(e.target, overlay);
        }
    });

    // Back to this page with the browser's Back button: it comes back as it was left, loading screen and all. The
    // work it was waiting for is long done, so the screen goes and its progress checks stop
    window.addEventListener('pageshow', (e) => {
        if (!e.persisted) return;
        const overlay = document.getElementById('loading-overlay');
        if (!overlay) return;
        clearTimeout(progressTimer);
        overlay.style.display = 'none';
        const text = document.getElementById('loading-text');
        if (text && text.dataset.idle) text.textContent = text.dataset.idle;
        const box = overlay.querySelector('.lp-box');
        if (box) box.remove();
    });

    // Live progress on the overlay (progress.py): the form carries a progress_id, and until the next page arrives the
    // overlay asks the app twice a second how far it is: the step, a bar when the step counts its work (label 12 of
    // 40), the time so far and the last lines of the log that run wrote
    let progressTimer = null;
    const newProgressId = () => (window.crypto && crypto.randomUUID ? crypto.randomUUID()
        : Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 12)).replace(/[^A-Za-z0-9-]/g, '');
    function followProgress(form, overlay) {
        if ((form.getAttribute('method') || 'get').toLowerCase() !== 'post') return;
        const id = newProgressId();
        let input = form.querySelector('input[name="progress_id"]');
        if (!input) {
            input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'progress_id';
            form.append(input);
        }
        input.value = id;  // the browser builds what it sends after this handler, so it goes with the form
        watchProgress(overlay, id);
    }
    // Pages that send their work by script (the matching page's Finalise) use the same overlay: a new id, sent as the
    // X-Progress-Id header, then followed here
    window.whProgress = { newId: newProgressId, watch: (overlay, id) => watchProgress(overlay, id),
                          stop: () => clearTimeout(progressTimer) };

    function watchProgress(overlay, id) {
        let box = overlay.querySelector('.lp-box');
        if (!box) {
            const make = (tag, cls) => { const n = document.createElement(tag); n.className = cls; return n; };
            box = make('div', 'lp-box');
            const track = make('div', 'lp-bar');
            track.setAttribute('role', 'progressbar');
            track.setAttribute('aria-valuemin', '0');
            track.setAttribute('aria-valuemax', '100');
            track.setAttribute('aria-label', 'Progress');
            track.append(make('div', 'lp-fill'));
            const meta = make('div', 'lp-meta');
            meta.append(make('span', 'lp-step'), make('span', 'lp-pct'));
            const log = make('ol', 'lp-lines');
            log.setAttribute('aria-live', 'polite');
            box.append(track, meta, log);
            overlay.append(box);
        }
        const bar = box.querySelector('.lp-bar'), fill = box.querySelector('.lp-fill');
        const step = box.querySelector('.lp-step'), pct = box.querySelector('.lp-pct'), lines = box.querySelector('.lp-lines');
        box.classList.add('lp-waiting');
        fill.style.width = '0%';
        bar.removeAttribute('aria-valuenow');
        step.textContent = 'Sending the file to the app…';
        pct.textContent = '';
        lines.replaceChildren();

        const started = Date.now();
        clearTimeout(progressTimer);
        const poll = async () => {
            let finished = false;
            try {
                const res = await fetch(`/api/progress/${id}`, { cache: 'no-store' });
                if (res.ok) {
                    const d = await res.json();
                    finished = d.finished;
                    box.classList.toggle('lp-waiting', !d.counted);
                    if (d.counted) {
                        fill.style.width = `${d.percent}%`;
                        bar.setAttribute('aria-valuenow', Math.round(d.percent));
                    }
                    step.textContent = d.step;
                    pct.textContent = (d.counted ? `${Math.round(d.percent)}% · ` : '') + `${Math.round(d.seconds)}s`;
                    // Every line of the run so far, the fine detail too; kept at the newest unless scrolled up to read
                    const atEnd = lines.scrollHeight - lines.scrollTop - lines.clientHeight < 24;
                    lines.replaceChildren(...d.lines.map(x => {
                        const li = document.createElement('li');
                        li.className = `lp-line lv-${x.level}`;
                        li.style.paddingLeft = `${Math.min(x.depth, 4) * 12}px`;
                        const t = document.createElement('span');
                        t.className = 'lp-line-time';
                        t.textContent = x.clock || x.time;
                        li.append(t, document.createTextNode(x.text));
                        return li;
                    }));
                    if (atEnd) lines.scrollTop = lines.scrollHeight;
                } else {
                    pct.textContent = `${Math.round((Date.now() - started) / 1000)}s`;  // still sending, or not followed
                }
            } catch (err) {
                // the app is busy or the next page is on its way: try again
            }
            if (!finished && Date.now() - started < 30 * 60 * 1000) progressTimer = setTimeout(poll, 500);
        };
        progressTimer = setTimeout(poll, 300);
    }

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
// 10. LABEL MAKER DRAG AND DROP CONFIGURATOR + LIVE CELL PREVIEW
// =========================================
// One configurator per tab (.pl-configurator, data-tab): each tab has its own blocks, as its rows of the file differ.
// A tab's blocks are dragged between its .pl-headers-pool (not used) and its .pl-attr-list (top to bottom = the order
// on the label); the order goes to the server in its .pl-attr-order ('attribute_order::<tab>'). A block stays in its
// own tab's configurator. While a block is dragged, the other blocks slide out of its way, and the tab's live preview
// (.pl-cell-preview) redraws its first package group's item boxes with the order as it stands - the same sizes and
// rules as the PDF (packing_label_generator._layout_cell / _draw_cell), drawn on canvases so it keeps up with the mouse.
document.addEventListener("DOMContentLoaded", function() {
    document.querySelectorAll('.pl-configurator').forEach(configurator);
});
function configurator(root) {
    const layoutList = root.querySelector('.pl-attr-list');
    const headersList = root.querySelector('.pl-headers-pool');
    const hiddenInput = root.querySelector('.pl-attr-order');

    if (!layoutList || !hiddenInput) return;
    const lists = [layoutList, headersList].filter(Boolean);
    const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    const currentOrder = () => [...layoutList.querySelectorAll(':scope > .sortable-item')].map(item => item.getAttribute('data-id'));
    const updateHiddenInput = () => { hiddenInput.value = currentOrder().join(','); };
    updateHiddenInput();

    const preview = cellPreview(root.querySelector('.pl-cell-preview'), () => currentOrder().filter(id => id && id !== 'empty'));

    // Any change to the cell's blocks (a drag in progress included): keep the order and the preview up to date
    new MutationObserver(() => { updateHiddenInput(); preview.schedule(); }).observe(layoutList, { childList: true });

    let draggedItem = null;

    lists.forEach(list => list.querySelectorAll(':scope > .sortable-item').forEach(item => {
        item.addEventListener('dragstart', function(e) {
            draggedItem = this;
            e.dataTransfer.effectAllowed = 'move';
            try { e.dataTransfer.setData('text/plain', this.getAttribute('data-id') || ''); } catch (err) { /* Firefox needs data to start a drag */ }
            // After the browser has taken its drag image (and only if the drag hasn't already ended)
            requestAnimationFrame(() => { if (draggedItem === this) this.classList.add('is-dragging'); });
        });
        item.addEventListener('dragend', function() {
            this.classList.remove('is-dragging');
            draggedItem = null;
            updateHiddenInput();
            preview.schedule();
        });
    }));

    lists.forEach(list => {
        list.addEventListener('dragover', function(e) {
            if (!draggedItem) return;
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            const afterElement = getDragAfterElement(this, e.clientY);
            // Only move when the block's slot actually changes: dragover fires many times a second
            const inPlace = draggedItem.parentNode === this &&
                (afterElement ? draggedItem.nextElementSibling === afterElement : !draggedItem.nextElementSibling);
            if (inPlace) return;
            slideWhileMoving(() => afterElement ? this.insertBefore(draggedItem, afterElement) : this.appendChild(draggedItem));
        });
        list.addEventListener('drop', function(e) { if (draggedItem) e.preventDefault(); });  // no navigating to the dropped text
    });

    // The block the dragged one goes before, from the blocks' layout positions (offsetTop isn't moved by the
    // slide animation's transforms, so a block mid-slide doesn't make the slot flicker back and forth)
    function getDragAfterElement(container, clientY) {
        const rect = container.getBoundingClientRect();
        const y = clientY - rect.top - container.clientTop + container.scrollTop;
        for (const child of container.querySelectorAll(':scope > .sortable-item')) {
            if (child === draggedItem) continue;
            if (y < child.offsetTop + child.offsetHeight / 2) return child;
        }
        return null;
    }

    // FLIP: note where the other blocks are, move the dragged block, then slide each one from its old place to its new one
    function slideWhileMoving(move) {
        const others = lists.flatMap(l => [...l.querySelectorAll(':scope > .sortable-item')]).filter(el => el !== draggedItem);
        const before = new Map(others.map(el => [el, el.getBoundingClientRect().top]));
        move();
        if (reduceMotion) return;
        others.forEach(el => {
            if (!el.animate) return;
            el.getAnimations().forEach(a => a.cancel());
            const dy = before.get(el) - el.getBoundingClientRect().top;
            if (Math.abs(dy) < 1) return;
            el.animate([{ transform: `translateY(${dy}px)` }, { transform: 'translateY(0)' }],
                       { duration: 160, easing: 'cubic-bezier(0.2, 0.7, 0.3, 1)' });
        });
    }
}

// The live cell preview: draws each item box of a tab's first package group (its .pl-cell-preview's data-cells, from
// packing_label_generator.cell_preview) on its own canvas, in PDF points, following _layout_cell and
// _draw_cell line for line. getOrder() gives the blocks top to bottom. Returns { schedule }: redraws once
// on the next animation frame however often it's asked.
function cellPreview(box, getOrder) {
    let data = null;
    try { data = box ? JSON.parse(box.dataset.cells || 'null') : null; } catch (e) { data = null; }
    if (!box || !data || !data.items || !data.items.length) return { schedule: () => {} };

    const [CW, CH] = data.cell_pt;
    const LINE = 1.2;
    const FONT = 'Helvetica, Arial, sans-serif';  // the PDF's helv/hebo; Arial has the same letter widths
    const TEXT_STYLES = {  // packing_label_generator.TEXT_STYLES: base size, bold, colour
        desc: [9, false, '#000'], dimension: [9, true, '#000'], material: [9, false, '#000'],
        install: [9, true, 'rgb(204,0,0)'], notes: [9, false, '#000'],
    };

    let frame = 0;
    const schedule = () => { if (!frame) frame = requestAnimationFrame(() => { frame = 0; render(); }); };

    // Each item's images, loaded once; a picture that can't be read shows "Image Error", as in the PDF
    const images = data.items.map(item => (item.thumbnails || []).map(src => {
        const img = { el: null, failed: !src };
        if (src) {
            img.el = new Image();
            img.el.onload = schedule;
            img.el.onerror = () => { img.failed = true; schedule(); };
            img.el.src = src;
        }
        return img;
    }));
    const canvases = data.items.map((item, i) => {
        const canvas = document.createElement('canvas');
        canvas.setAttribute('role', 'img');
        canvas.setAttribute('aria-label', `Item ${i + 1}: ${item.job_no || 'no job number'}`);
        box.appendChild(canvas);
        return canvas;
    });

    function setFont(ctx, bold, size) { ctx.font = `${bold ? 'bold ' : ''}${size}px ${FONT}`; }

    // _wrap: words onto lines no wider than width; a word too long for a line is cut
    function wrap(ctx, text, bold, size, width) {
        setFont(ctx, bold, size);
        const w = s => ctx.measureText(s).width;
        const lines = [];
        for (const paragraph of String(text).split(/\r\n|\r|\n/)) {
            let line = '';
            for (let word of paragraph.split(/\s+/).filter(Boolean)) {
                while (w(word) > width) {
                    let cut = word.length;
                    while (cut > 1 && w(word.slice(0, cut)) > width) cut--;
                    if (line) { lines.push(line); line = ''; }
                    lines.push(word.slice(0, cut));
                    word = word.slice(cut);
                }
                const candidate = `${line} ${word}`.trim();
                if (line && w(candidate) > width) { lines.push(line); line = word; }
                else line = candidate;
            }
            if (line) lines.push(line);
        }
        return lines;
    }

    // _draw_lines: each line centred between x0 and x1
    function drawLines(ctx, lines, x0, x1, y, bold, size, colour) {
        setFont(ctx, bold, size);
        ctx.fillStyle = colour;
        ctx.textBaseline = 'alphabetic';
        for (const line of lines) {
            const lw = ctx.measureText(line).width;
            ctx.fillText(line, x0 + (x1 - x0 - lw) / 2, y + size * 0.92);
            y += size * LINE;
        }
    }

    function roundRect(ctx, x0, y0, x1, y1, r) {
        r = Math.min(r, (x1 - x0) / 2, (y1 - y0) / 2);
        ctx.beginPath();
        ctx.moveTo(x0 + r, y0);
        ctx.arcTo(x1, y0, x1, y1, r);
        ctx.arcTo(x1, y1, x0, y1, r);
        ctx.arcTo(x0, y1, x0, y0, r);
        ctx.arcTo(x0, y0, x1, y0, r);
        ctx.closePath();
    }

    // _layout_cell: stacks the blocks top-down and returns the height used; draws only when draw is true
    // One block of an item box at y; returns the y after it (and its gap). Draws only when draw is true
    function block(ctx, item, imgs, attr, textScale, thumbScale, draw, y) {
        const pad = 4, x0 = pad, x1 = CW - pad, innerW = x1 - x0, gap = 3 * textScale;
        if (attr === 'thumbnail') {
            if (!imgs.length) return y;
            const h = CH * 0.35 * thumbScale;
            if (draw) {
                // One image: centred, 70% of the box width. Several: share the full width, up to 3 a row
                const n = imgs.length, perRow = Math.min(n, 3), rows = Math.ceil(n / perRow);
                const areaW = n === 1 ? CW * 0.7 : innerW;
                const ax = (CW - areaW) / 2;
                const slotW = (areaW - (perRow - 1) * 2) / perRow;
                const slotH = (h - (rows - 1) * 2) / rows;
                imgs.forEach((img, i) => {
                    const sx = ax + (i % perRow) * (slotW + 2);
                    const sy = y + Math.floor(i / perRow) * (slotH + 2);
                    if (img.failed) {
                        setFont(ctx, false, 8);
                        ctx.fillStyle = '#000';
                        const tw = ctx.measureText('Image Error').width;
                        ctx.fillText('Image Error', sx + (slotW - tw) / 2, sy + 8);
                    } else if (img.el && img.el.complete && img.el.naturalWidth) {
                        const s = Math.min(slotW / img.el.naturalWidth, slotH / img.el.naturalHeight);
                        const w = img.el.naturalWidth * s, ih = img.el.naturalHeight * s;
                        ctx.drawImage(img.el, sx + (slotW - w) / 2, sy + (slotH - ih) / 2, w, ih);
                    }
                });
            }
            y += h + gap;
        } else if (attr === 'job_no') {
            if (!item.job_no) return y;
            const size = 14 * textScale;
            const lines = wrap(ctx, item.job_no, true, size, innerW - 4);
            const h = lines.length * size * LINE + 6 * textScale;
            if (draw) {
                roundRect(ctx, 3, y, CW - 3, y + h, 3);
                ctx.fillStyle = '#000';
                ctx.fill();
                drawLines(ctx, lines, x0, x1, y + 3 * textScale, true, size, '#fff');
            }
            y += h + gap;
        } else if (attr === 'barcode') {
            const modules = item.modules;
            if (!modules) return y;
            // Code 128 needs blank modules either side; the bars get as wide as the box allows
            const moduleW = Math.min(data.barcode_max_module, innerW / (modules.length + 2 * data.barcode_quiet));
            const barH = 22 * textScale, size = 6.5 * textScale;
            const h = barH + 2 + size * LINE;
            if (draw) {
                const bx = (CW - modules.length * moduleW) / 2;
                ctx.fillStyle = '#000';
                let runStart = -1;
                for (let i = 0; i <= modules.length; i++) {
                    const bar = i < modules.length && modules[i] === '1';
                    if (bar && runStart < 0) runStart = i;
                    else if (!bar && runStart >= 0) { ctx.fillRect(bx + runStart * moduleW, y, (i - runStart) * moduleW, barH); runStart = -1; }
                }
                drawLines(ctx, [item.barcode], x0, x1, y + barH + 2, false, size, '#000');
            }
            y += h + gap;
        } else if (attr === 'qty') {
            if (!item.qty) return y;
            const size = 20 * textScale, border = 2.5 * textScale;
            const lines = wrap(ctx, item.qty, true, size, innerW - 2 * border - 8);
            setFont(ctx, true, size);
            const widest = Math.max(...lines.map(l => ctx.measureText(l).width));
            const boxW = Math.min(innerW, Math.max(70 * textScale, widest + 16 * textScale));
            let h = lines.length * size * LINE + 6 * textScale;
            if (draw) {
                const bx0 = (CW - boxW) / 2, half = border / 2;
                ctx.lineWidth = border;
                ctx.strokeStyle = '#000';
                ctx.strokeRect(bx0 + half, y + half, boxW - border, h - border);
                drawLines(ctx, lines, bx0, bx0 + boxW, y + 3 * textScale, true, size, '#000');
            }
            // The note written with the quantity ('4 (1 of each kind)'), under the box as the PDF has it
            if (item.qty_note) {
                const noteSize = 9 * textScale;
                const noteLines = wrap(ctx, `(${item.qty_note})`, true, noteSize, innerW);
                if (draw) drawLines(ctx, noteLines, x0, x1, y + h + 2 * textScale, true, noteSize, '#000');
                h += 2 * textScale + noteLines.length * noteSize * LINE;
            }
            y += h + gap;
        } else if (TEXT_STYLES[attr] || /^row_\d+$/.test(attr)) {  // a row of the file reads like a description
            const val = item[attr];
            if (!val) return y;
            const [base, bold, colour] = TEXT_STYLES[attr] || TEXT_STYLES.desc;
            const size = base * textScale;
            const lines = wrap(ctx, val, bold, size, innerW);
            if (draw) drawLines(ctx, lines, x0, x1, y, bold, size, colour);
            y += lines.length * size * LINE + gap;
        }
        return y;
    }
    // slots ({block: height}, planCells): each block takes the same height in every box, empty or not (plan_cells);
    // heights: filled with what each block of this item needs
    function layout(ctx, item, imgs, order, textScale, thumbScale, draw, slots, heights) {
        const pad = 4, gap = 3 * textScale;
        let y = pad;
        for (const attr of order) {
            if (slots) {
                if (!slots[attr]) continue;  // no box has this block
                block(ctx, item, imgs, attr, textScale, thumbScale, draw, y);
                y += slots[attr] + gap;
                continue;
            }
            const next = block(ctx, item, imgs, attr, textScale, thumbScale, draw, y);
            if (heights && next > y) heights[attr] = next - y - gap;
            y = next;
        }
        return y - gap + pad;
    }

    // _draw_cell: shrink the text first (down to 65%), then the images, until the whole stack fits the box
    // plan_cells: one layout for every box shown (each block as tall as the tallest item needs it), at the largest
    // scales the stack fits; an item without a value for a block leaves its space, so the blocks line up
    function scales(step) {
        const s = 1 - step * 0.01;
        return [Math.max(s, 0.65), s >= 0.65 ? 1 : s / 0.65];
    }
    function planCells(ctx, shown, order) {
        const at = (step) => {
            const [ts, th] = scales(step), slots = {};
            shown.forEach(([item, imgs]) => {
                const used = {};
                layout(ctx, item, imgs, order, ts, th, false, null, used);
                Object.entries(used).forEach(([k, h]) => { slots[k] = Math.max(slots[k] || 0, h); });
            });
            const gap = 3 * ts, hs = Object.values(slots);
            return { ts, th, slots, total: 8 + hs.reduce((sum, h) => sum + h + gap, 0) - (hs.length ? gap : 0) };
        };
        let lo = 0, hi = 70;
        if (at(0).total <= CH) hi = 0;
        else while (lo < hi) { const mid = (lo + hi) >> 1; if (at(mid).total <= CH) hi = mid; else lo = mid + 1; }
        return at(hi);
    }
    function drawCell(ctx, item, imgs, order, plan) {
        layout(ctx, item, imgs, order, plan.ts, plan.th, true, plan.slots);
        roundRect(ctx, 0.375, 0.375, CW - 0.375, CH - 0.375, data.radius_pt);
        ctx.lineWidth = 0.75;
        ctx.strokeStyle = '#000';
        ctx.stroke();
    }

    // Code boxes stay on the labels grouped or not (Label order -> Don't group): no column of a pack is left off
    function render() {
        const order = getOrder();
        const dpr = window.devicePixelRatio || 1;
        const shown = data.items.map((item, i) => [item, images[i]]);
        const plan = planCells(canvases[0].getContext('2d'), shown, order);
        canvases.forEach((canvas, i) => {
            const cssW = canvas.clientWidth || 160;
            const pxW = Math.round(cssW * dpr), pxH = Math.round(cssW * CH / CW * dpr);
            if (canvas.width !== pxW || canvas.height !== pxH) { canvas.width = pxW; canvas.height = pxH; }
            const ctx = canvas.getContext('2d');
            ctx.setTransform(1, 0, 0, 1, 0, 0);
            ctx.clearRect(0, 0, pxW, pxH);
            ctx.setTransform(pxW / CW, 0, 0, pxW / CW, 0, 0);  // draw in PDF points
            drawCell(ctx, data.items[i], images[i], order, plan);
        });
    }

    if (window.ResizeObserver) new ResizeObserver(schedule).observe(box);
    schedule();
    return { schedule };
}
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
        renderAddressParts(row.querySelector('.pl-show-address'), d);
        row.dataset.missing = missingParts(d).join(',');
        row.dataset.state = (d.state || '').toUpperCase();
        row.dataset.postcode = String(d.postcode || '').trim();
        row.dataset.country = country;
        document.dispatchEvent(new CustomEvent('pl-countries-changed'));  // the export details' count follows
    }
    // The fields of a row's ✏️ editor (in the row itself: the receiver's cell and the address cell)
    function fill(row, d) {
        row.querySelectorAll('input[data-field]').forEach((input) => {
            if (input.type === 'checkbox') input.checked = !!d[input.dataset.field];
            else input.value = d[input.dataset.field] || '';
        });
    }
    function read(row) {
        const d = {};
        row.querySelectorAll('input[data-field]').forEach((input) => {
            d[input.dataset.field] = input.type === 'checkbox' ? input.checked : input.value.trim();
        });
        d.state = (d.state || '').toUpperCase();
        d.suburb = (d.suburb || '').toUpperCase();
        if ('country' in d) d.country = (d.country || '').toUpperCase();
        return d;
    }

    const currentOf = new Map();  // row -> its address as it stands (Excel + edits), for merging
    rows.forEach((row) => {
        const id = row.dataset.id;
        const original = JSON.parse(row.dataset.original);
        const current = () => {
            const d = { ...original };
            ADDRESS.concat('authority_to_leave').forEach((k) => { if (edits[id] && k in edits[id]) d[k] = edits[id][k]; });
            return d;
        };
        currentOf.set(row, current);
        // ✏️ opens the row in place: the receiver and the address fold away and their fields unfold where they were
        // shown (styles.css .pl-fold / .pl-unfold). What's folded is inert: no tabbing into hidden fields.
        const folds = [...row.querySelectorAll('.pl-fold')], unfolds = [...row.querySelectorAll('.pl-unfold')];
        const editing = () => row.classList.contains('pl-editing');
        function setEditing(on) {
            row.classList.toggle('pl-editing', on);
            folds.forEach((x) => { x.inert = on; });
            unfolds.forEach((x) => { x.inert = !on; });
        }
        const close = () => setEditing(false);

        row.querySelector('.pl-edit-btn').addEventListener('click', () => {
            if (editing()) { close(); return; }
            fill(row, current());
            setEditing(true);
            row.querySelector('input[data-field="line1"]').focus({ preventScroll: true });
        });

        // What an edit carries over from before: the service picked, Keep separate (same-name groups) and the
        // Keep as is choices (they name the exact address they're for, so a changed address simply doesn't match)
        const carried = (key) => Object.fromEntries(['service_code', 'separate', 'ref_keep', 'book_keep', 'no_fill']
            .filter((k) => edits[key] && edits[key][k]).map((k) => [k, edits[key][k]]));

        row.querySelector('.pl-edit-save').addEventListener('click', () => {
            const d = read(row);
            const changed = ADDRESS.some((k) => (d[k] || '') !== (original[k] || '')) ||
                            d.authority_to_leave !== !!original.authority_to_leave;
            const keep = carried(id);
            // Saving counts as checked, so a row that isn't in the address book stops being flagged
            // (kept through Update Previews; the address is saved to the book on Generate)
            const checked = row.classList.contains('pl-needs-check') || (edits[id] && edits[id].checked) ? { checked: true } : {};
            const wasFilled = !!(edits[id] && edits[id].source === 'official');
            edits[id] = changed ? { ...keep, ...checked, ...d } : { ...keep, ...checked };
            if (wasFilled && !changed) edits[id].no_fill = true;  // back to the Excel address: not filled in again
            tidy(id);
            show(row, changed ? d : original);
            row.querySelector('.pl-tag-edited').hidden = !changed;
            row.querySelector('.pl-tag-book').hidden = true;
            row.querySelector('.pl-tag-filled').hidden = true;
            refreshGroup(row);
            markChecked(row);
            if (d.postcode) markAddressed(row);
            close();
            applyFilters();
            PL_REF.recheck(row, changed ? d : original, edits[id] && edits[id].ref_keep);
        });
        row.querySelector('.pl-edit-cancel').addEventListener('click', close);

        // A suggestion from the official address data applied in one click: an edit like ✏️ Save, counted as checked
        row.addEventListener('pl-ref-use', (e) => {
            const d = { ...current(), ...e.detail.fields };
            if ('suburb' in e.detail.fields) d.suburb = (d.suburb || '').toUpperCase();
            const keep = carried(id);
            // From the address book: tagged as such (and its options are done with); else an edit
            const fromBook = e.detail.source === 'book';
            edits[id] = { ...keep, checked: true, ...d, ...(fromBook ? { source: 'book' } : {}) };
            tidy(id);
            show(row, d);
            row.querySelector('.pl-tag-edited').hidden = fromBook;
            row.querySelector('.pl-tag-book').hidden = !fromBook;
            row.querySelector('.pl-tag-filled').hidden = true;
            refreshGroup(row);
            if (fromBook) row.querySelector('.pl-book-opts').replaceChildren();
            if (editing()) fill(row, d);  // the open editor shows what was just applied
            markChecked(row);
            if (d.postcode) markAddressed(row);
            applyFilters();
            PL_REF.recheck(row, d, null);
        });
        // Keep as is for the address book's option: not offered again for this exact address (Update Previews too)
        row.addEventListener('pl-book-keep', () => {
            const d = current();
            editOf(id).book_keep = Object.fromEntries(['line1', 'line2', 'suburb', 'state', 'postcode'].map((k) => [k, d[k] || '']));
            save();
            row.querySelector('.pl-book-opts').replaceChildren();
        });
        // Keep as is: this exact address stops being flagged (the address book remembers it on Generate)
        row.addEventListener('pl-ref-keep', () => {
            const d = current();
            editOf(id).ref_keep = Object.fromEntries(['line1', 'line2', 'suburb', 'state', 'postcode'].map((k) => [k, d[k] || '']));
            save();
            PL_REF.render(row, { level: 'kept', message: 'Kept as is', fields: {}, options: [] });
        });
        // Clear: the address fields only; receiver and contact stay
        row.querySelector('.pl-edit-clear').addEventListener('click', () => {
            ['line1', 'line2', 'suburb', 'state', 'postcode', 'country'].forEach((k) => {
                const input = row.querySelector(`input[data-field="${k}"]`);
                if (input) input.value = '';
            });
            const first = row.querySelector('input[data-field="line1"]');
            first.focus();
            first.dispatchEvent(new Event('input', { bubbles: true }));
        });
        // Suggestions while typing: verified saved addresses, then other saved ones, then the official data
        attachFieldSuggest(row);

        // Closest saved addresses: clicking one fills the fields; typing in any field re-ranks the list
        const closest = attachClosest(row, (entry) => fill(row, entry), () => read(row));
        row.querySelector('.pl-edit-reset').addEventListener('click', () => {
            fill(row, original);
            closest.refresh();  // the Excel address's closest matches, the address-book fill among them
        });

        // Enter in the editor's fields saves the row instead of submitting the whole page; Esc closes it
        unfolds.forEach((box) => box.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.target.matches('input')) { e.preventDefault(); row.querySelector('.pl-edit-save').click(); }
            if (e.key === 'Escape') close();
        }));

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

    // The text a row shows: not the ✏️ editor's labels, other receivers' saved addresses offered as options, or
    // anything hidden, so a search (and "Apply to N shown") finds the consignments it names and no others
    // A drop-down (the service) shows only what's chosen, not its whole list
    const SEARCH_SKIP = '.pl-editor, .pl-book-opts, .pl-ref, [hidden], select';
    function searchText(row) {
        const parts = [];
        const shown = (el) => {
            const skip = el.closest(SEARCH_SKIP);
            return !skip || skip === row || !row.contains(skip);  // the row itself may be hidden by this filter
        };
        const walker = document.createTreeWalker(row, NodeFilter.SHOW_TEXT);
        for (let n = walker.nextNode(); n; n = walker.nextNode()) {
            if (n.parentElement && shown(n.parentElement)) parts.push(n.nodeValue);
        }
        row.querySelectorAll('select').forEach((select) => {
            const holder = select.parentElement;
            if (holder && shown(holder)) parts.push(...[...select.selectedOptions].map((o) => o.textContent));
        });
        return parts.join(' ').toLowerCase();
    }

    // A courier's area in the State filter ('Flash-NSW'): Australian consignments whose postcode is in its ranges
    function inArea(option) {
        let ranges = [];
        try { ranges = JSON.parse(option.dataset.postcodes || '[]'); } catch (e) { ranges = []; }
        return (row) => {
            const pc = row.dataset.postcode || '';
            return (row.dataset.country || 'AU') === 'AU' && /^\d{4}$/.test(pc) &&
                ranges.some(([from, to]) => Number(pc) >= from && Number(pc) <= to);
        };
    }
    function applyFilters() {
        const state = stateFilter.value;
        const area = state.startsWith('region:') ? inArea(stateFilter.selectedOptions[0]) : null;
        const text = textFilter.value.trim().toLowerCase();
        let shown = 0;
        rows.forEach((row) => {
            const visible = (!state || (area ? area(row) : row.dataset.state === state)) &&
                            (!text || searchText(row).includes(text));
            row.hidden = !visible;
            if (visible) shown += 1;
        });
        // A group's heading goes with its rows
        document.querySelectorAll('.pl-group-row').forEach((g) => {
            g.hidden = rows.filter((r) => r.dataset.group === g.dataset.group).every((r) => r.hidden);
        });
        countLabel.textContent = shown === rows.length ? `${rows.length} consignments` : `Showing ${shown} of ${rows.length}`;
        applyBtn.textContent = shown === rows.length ? 'Apply to all' : `Apply to ${shown} shown`;
    }
    // One receiver's consignments at different addresses, under a heading. Merge into one happens on the page at
    // once: the address picked goes on every one of them (as an edit, like ✏️), their items move into the kept row and
    // the other rows go. Nothing is sent until Generate (or Update Previews), where build_consignments makes them one
    // consignment because they now share an address. Keep separate: each stays at its own address, the heading goes.
    const oneLine = (d) => ['line1', 'line2', 'suburb', 'state', 'postcode'].map((k) => d[k]).filter(Boolean).join(', ');
    const groupRows = (head) => rows.filter((r) => r.dataset.group === head.dataset.group);
    // The Keep list follows the rows: an address typed or fixed in a member row since the page opened is offered
    function refreshGroup(row) {
        const head = row.dataset.group && document.querySelector(`.pl-group-row[data-group="${row.dataset.group}"]`);
        if (!head) return;
        const group = JSON.parse(head.querySelector('.pl-group-merge').dataset.group);
        const select = head.querySelector('.pl-group-keep');
        const was = select.value;
        select.replaceChildren();
        const offered = group.map((m, i) => {
            const r = rows.find((x) => x.dataset.id === m.id);
            return r && { m, i, d: currentOf.get(r)() };
        }).filter((x) => x && oneLine(x.d));
        offered.sort((a, b) => !a.d.postcode - !b.d.postcode);  // complete addresses first, as the page starts
        offered.forEach(({ m, i, d }) => {
            const o = document.createElement('option');
            o.value = String(i);
            o.textContent = `#${m.number} ${oneLine(d)}${d.postcode ? '' : ' (no postcode)'}`;
            select.append(o);
        });
        if ([...select.options].some((o) => o.value === was)) select.value = was;
        const any = select.options.length > 0;
        select.closest('.pl-group-keep-label').hidden = !any;
        head.querySelector('.pl-group-none').hidden = any;
        head.querySelector('.pl-group-merge').disabled = !any;
    }
    function removeRow(row) {
        row.remove();
        rows.splice(rows.indexOf(row), 1);
        document.dispatchEvent(new CustomEvent('pl-row-removed', { detail: { row } }));  // the To check list drops it
    }
    function endGroup(head) {
        groupRows(head).forEach((r) => { r.classList.remove('pl-in-group'); delete r.dataset.group; });
        head.remove();
    }
    document.querySelectorAll('.pl-group-row').forEach((head) => {
        const select = head.querySelector('.pl-group-keep');
        select.addEventListener('focus', () => refreshGroup(head));
        head.querySelector('.pl-group-merge').addEventListener('click', (e) => {
            const group = JSON.parse(e.currentTarget.dataset.group);
            if (!select.options.length) return;
            const keep = group[Number(select.value)];
            const pairs = group.map((m) => [m, rows.find((r) => r.dataset.id === m.id)]).filter(([, r]) => r);
            const members = pairs.map(([, r]) => r);
            const numberOf = (r) => pairs.find(([, x]) => x === r)[0].number;
            const keepRow = members.find((r) => r.dataset.id === keep.id);
            if (!keepRow) return;
            // They become one consignment, so one service: the kept row's. Say so when theirs differ.
            const serviceOf = (r) => r.querySelector('.pl-service-input').value || mainService.value;
            const services = [...new Set(members.map(serviceOf).filter(Boolean))];
            if (services.length > 1 && !confirm(`These consignments use different services: ${members.map((r) => `#${numberOf(r)} ${serviceOf(r) || '-'}`).join(', ')}.\n\n`
                + `Merged, they all go with the service of the address you keep (#${keep.number}: ${serviceOf(keepRow) || '-'}). Merge?`)) return;
            const kept = currentOf.get(keepRow)();
            const fields = Object.fromEntries(['line1', 'line2', 'suburb', 'state', 'postcode', 'country'].map((k) => [k, kept[k] || '']));
            const service = keepRow.querySelector('.pl-service-input').value;
            group.forEach((m) => m.sources.forEach((sid) => {
                const edit = { ...(edits[sid] || {}), checked: true, ...fields };
                delete edit.source; delete edit.filled;  // the address is now the one picked here
                if (service) edit.service_code = service; else delete edit.service_code;  // '' = same as main
                edits[sid] = edit;
            }));
            save();
            // The kept row takes every item; the others go
            const items = keepRow.querySelector('.pl-items');
            const others = members.filter((r) => r !== keepRow);
            others.forEach((r) => { r.querySelectorAll('.pl-carton').forEach((x) => items.append(x)); removeRow(r); });
            keepRow.querySelector('.pl-items-count').textContent = items.querySelectorAll('.pl-carton').length;
            const note = document.createElement('div');
            note.className = 'issue-detail pl-merged-note';
            note.textContent = `Merged with ${others.map((r) => `#${numberOf(r)}`).join(', ')}: one consignment when you generate.`;
            keepRow.querySelector('.pl-show-address').after(note);
            endGroup(head);
            show(keepRow, kept);
            markChecked(keepRow);
            if (kept.postcode) markAddressed(keepRow);
            applyFilters();
        });
        head.querySelector('.pl-group-separate').addEventListener('click', (e) => {
            const group = JSON.parse(e.currentTarget.dataset.group);
            group.forEach((m) => m.sources.forEach((sid) => { editOf(sid).separate = true; }));
            save();
            endGroup(head);
            // The file's warning about this store's addresses goes from the To check list too
            document.dispatchEvent(new CustomEvent('pl-kept-separate', { detail: { receiver: group[0].receiver } }));
            applyFilters();
        });
    });

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
    const picked = new Set();  // ids ticked in select mode

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
        const pick = document.createElement('input');
        pick.type = 'checkbox'; pick.className = 'ab-pick'; pick.checked = picked.has(entry.id);
        pick.setAttribute('aria-label', `Select ${entry.receiver}`);
        pick.addEventListener('click', (e) => tick(tr, pick.checked, e.shiftKey));
        tr.classList.toggle('ab-picked', pick.checked);
        actions.append(pick, edit, del);
        const who = cell(entry.receiver);
        if (entry.verified_at) {
            const v = document.createElement('span');
            v.className = 'ab-verified';
            v.textContent = '✓ Verified';
            v.title = `Confirmed by the courier portal on ${when(entry.verified_at)}`;
            who.append(v);
        }
        if (entry.contact) { const c = document.createElement('div'); c.className = 'ab-detail'; c.textContent = `Attn ${entry.contact}`; who.append(c); }
        const cells = { line1: cell([entry.line1, entry.line2].filter(Boolean).join(', ')), suburb: cell(entry.suburb),
                        state: cell(entry.state), postcode: cell(entry.postcode), country: cell(entry.country) };
        tr.append(actions, who, cells.line1, cells.suburb, cells.state, cells.postcode, cells.country,
                  cell(entry.authority_to_leave ? 'Y' : ''), cell(entry.use_count, 'ab-num'), cell(when(entry.last_used_at)));
        showReference(tr, entry, cells);
        return tr;
    }

    // The official address data's check of an entry: the cells that differ underlined in amber, the reason and one-click
    // fixes under the address, Keep as is. An entry the courier portal verified isn't flagged.
    function showReference(tr, entry, cells) {
        const ref = entry.ref || {};
        const flagged = ref.status === 'mismatch' && !entry.verified_at;
        tr.classList.toggle('ab-needs-look-row', flagged);
        Object.entries(ref.fields || {}).forEach(([field, f]) => {
            const td = cells[field === 'line2' ? 'line1' : field];
            if (!td || entry.verified_at) return;
            td.classList.add(`ab-flag-${f.level}`);
            td.title = f.text;
        });
        if (!flagged && ref.status !== 'note') return;
        const box = document.createElement('div');
        box.className = flagged ? 'ab-ref ab-ref-bad' : 'ab-ref ab-ref-note';
        const text = document.createElement('span');
        text.textContent = (flagged ? '⚠ ' : 'ⓘ ') + ref.message;
        box.append(text);
        if (flagged) {
            (ref.options || []).forEach((o) => {
                const b = document.createElement('button');
                b.type = 'button'; b.className = 'ab-btn ab-btn-small'; b.textContent = `Use ${o.label}`;
                b.title = 'Change this entry to the official suburb, state and postcode';
                b.addEventListener('click', async () => {
                    const r = await send(`/api/addresses/${entry.id}`, 'PUT', { ...entry, ...o.fields });
                    if (!r.ok) { say(r.error, true); return; }
                    const fresh = rowFor(r.entry); fresh.classList.add('ab-flash'); tr.replaceWith(fresh);
                    say(`Saved ${r.entry.receiver}.`); refreshCount();
                });
                box.append(b);
            });
            const keep = document.createElement('button');
            keep.type = 'button'; keep.className = 'ab-btn ab-btn-small ab-btn-plain'; keep.textContent = 'Keep as is';
            keep.title = 'This address is right: stop flagging it (until it changes)';
            keep.addEventListener('click', async () => {
                const r = await send(`/api/addresses/${entry.id}/keep`, 'POST');
                if (!r.ok) { say(r.error, true); return; }
                const fresh = rowFor(r.entry); tr.replaceWith(fresh);
                say(`Kept ${r.entry.receiver} as is.`); refreshCount();
            });
            box.append(keep);
        }
        cells.line1.append(box);
    }

    // "Needs a look": only the addresses that don't match the official data (they're listed first anyway)
    const needsBtn = document.getElementById('ab-needs-look');
    let needsOnly = false;
    function showCount(n) {
        needsBtn.hidden = !n && !needsOnly;
        needsBtn.querySelector('span').textContent = `(${n})`;
    }
    async function refreshCount() {
        try { showCount((await (await fetch('/api/addresses?limit=1&q=%20')).json()).needs_look); } catch (e) { /* not important */ }
    }
    needsBtn.addEventListener('click', () => {
        needsOnly = !needsOnly;
        needsBtn.setAttribute('aria-pressed', String(needsOnly));
        load(true);
    });

    async function load(reset) {
        if (reset) offset = 0;
        if (controller) controller.abort();
        controller = new AbortController();
        try {
            const res = await fetch(`/api/addresses?limit=${PAGE}&offset=${offset}&q=${encodeURIComponent(query)}${needsOnly ? '&needs=1' : ''}`, { signal: controller.signal });
            const data = await res.json();
            showCount(data.needs_look);
            // Being checked against newly loaded official data: said once, no polling (refresh to see the results)
            if (data.checking && reset) say('Checking the addresses against the newly loaded official data: refresh the page to see the results.');
            if (reset) body.replaceChildren();
            data.rows.forEach((entry) => body.append(rowFor(entry)));
            offset += data.rows.length;
            if (page.classList.contains('ab-selecting')) showPicked();
            more.hidden = !data.has_more;
            const shown = body.querySelectorAll('.ab-row').length;
            summary.textContent = query
                ? (shown ? `${shown}${data.has_more ? '+' : ''} matching` : 'No matches')
                : (total ? `${total.toLocaleString()} addresses · ${data.needs_look ? 'needing a look first, then ' : ''}most recently used first` : 'No addresses yet. They are added automatically each time you generate packing labels, or use "Add address".');
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
        // Clear: the address fields only (receiver and contact stay); suggestions while typing
        tr.querySelector('.ab-clear').addEventListener('click', () => {
            ['line1', 'line2', 'suburb', 'state', 'postcode', 'country'].forEach((k) => { tr.querySelector(`[data-field="${k}"]`).value = ''; });
            tr.querySelector('[data-field="line1"]').focus();
        });
        attachFieldSuggest(tr);
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

    // Select mode: tick addresses (shift-click ticks every one between), then delete them all at once. What's ticked
    // stays ticked through a search or Load more, so several searches can make one selection.
    const selectBtn = document.getElementById('ab-select');
    const bar = document.getElementById('ab-select-bar');
    const pickedCount = document.getElementById('ab-select-count');
    const deleteBtn = document.getElementById('ab-select-delete');
    const all = document.getElementById('ab-select-all');
    let lastTicked = null;
    function showPicked() {
        pickedCount.textContent = `${picked.size} selected`;
        deleteBtn.disabled = !picked.size;
        deleteBtn.textContent = picked.size ? `🗑️ Delete ${picked.size} selected` : '🗑️ Delete selected';
        const boxes = [...body.querySelectorAll('.ab-row .ab-pick')];
        all.checked = boxes.length > 0 && boxes.every((b) => b.checked);
        all.indeterminate = !all.checked && boxes.some((b) => b.checked);
    }
    function setPicked(tr, on) {
        const id = Number(tr.dataset.id);
        if (on) picked.add(id); else picked.delete(id);
        tr.querySelector('.ab-pick').checked = on;
        tr.classList.toggle('ab-picked', on);
    }
    function tick(tr, on, range) {
        const rows = [...body.querySelectorAll('.ab-row')];
        if (range && lastTicked && rows.includes(lastTicked)) {
            const [a, b] = [rows.indexOf(lastTicked), rows.indexOf(tr)].sort((x, y) => x - y);
            rows.slice(a, b + 1).forEach((r) => setPicked(r, on));
        } else {
            setPicked(tr, on);
        }
        lastTicked = tr;
        showPicked();
    }
    function selecting(on) {
        page.classList.toggle('ab-selecting', on);
        selectBtn.setAttribute('aria-pressed', String(on));
        bar.hidden = !on;
        if (!on) {
            picked.clear();
            body.querySelectorAll('.ab-row').forEach((r) => setPicked(r, false));
            lastTicked = null;
        }
        showPicked();
    }
    selectBtn.addEventListener('click', () => selecting(!page.classList.contains('ab-selecting')));
    document.getElementById('ab-select-cancel').addEventListener('click', () => selecting(false));
    all.addEventListener('change', () => {
        body.querySelectorAll('.ab-row').forEach((r) => setPicked(r, all.checked));
        showPicked();
    });
    deleteBtn.addEventListener('click', async () => {
        const n = picked.size;
        if (!n || !confirm(`Delete ${n} address${n === 1 ? '' : 'es'} from the address book? This can't be undone `
            + '(Back up & move can save a copy first).')) return;
        deleteBtn.disabled = true;
        const r = await send('/api/addresses/delete', 'POST', { ids: [...picked] });
        if (!r.ok) { say(r.error, true); showPicked(); return; }
        body.querySelectorAll('.ab-row').forEach((tr) => {
            if (!picked.has(Number(tr.dataset.id))) return;
            const next = tr.nextElementSibling;
            if (next && next.classList.contains('ab-edit-row')) next.remove();
            tr.remove();
        });
        total -= r.entry.deleted;
        say(`Deleted ${r.entry.deleted} address${r.entry.deleted === 1 ? '' : 'es'}.`);
        selecting(false);
        load(true);
    });

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
    // ?spec=NAME (the preview's "Add its size" link): that spec's row, highlighted, the cursor in its first empty box.
    // Names compare as the app does: capitals, spaces and punctuation don't matter
    const wanted = new URLSearchParams(location.search).get('spec');
    const key = (t) => String(t || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const row = wanted && [...body.querySelectorAll('tr[data-spec]')].find((r) => key(r.dataset.spec) === key(wanted));
    if (row) {
        row.classList.add('ps-focus');
        row.scrollIntoView({ block: 'center' });
        const box = [...row.querySelectorAll('input.ps-num')].find((i) => !i.value) || row.querySelector('input.ps-num');
        if (box) box.focus({ preventScroll: true });
    }
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
    // The overlay shows the work's progress and log lines, as Generate's does: an id sent with the request (header)
    const watch = () => {
        if (!window.whProgress) return '';
        const id = window.whProgress.newId();
        window.whProgress.watch(overlay, id);
        return id;
    };
    const idle = () => { overlay.style.display = 'none'; if (window.whProgress) window.whProgress.stop(); };
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
    // Packs booked again (a stitch built on an earlier one): new courier page key -> 'new' | 'keep'
    const rebooked = session.rebooked || [];
    const rebook = {};
    rebooked.forEach((r) => { if (r.choice === 'new' || r.choice === 'keep') rebook[r.new.key] = r.choice; });
    const undecided = () => rebooked.filter((r) => !rebook[r.new.key]);

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

    // ---------- packs booked again: the new courier label or the earlier one ----------
    const rebookImage = (i, which, size) => root.dataset.imageUrl.replace(/0\.png$/, `${i}.png`)
        + `?list=rebooked&which=${which}${size ? `&size=${size}` : ''}`;
    function rebookCard(i, which, side, title) {
        const card = el('article', 'sm-card');
        const img = el('img', 'sm-thumb'); img.src = rebookImage(i, which); img.loading = 'lazy';
        img.alt = `${title} courier label (${side.file}, page ${side.page})`; img.title = 'Click to zoom';
        img.addEventListener('click', () => { zoom.querySelector('img').src = rebookImage(i, which, 'zoom'); zoom.showModal(); });
        const body = el('div', 'sm-card-body');
        body.append(el('div', 'sm-card-title', `${title} courier label`), el('div', 'ab-detail', `${side.file} · page ${side.page}`));
        card.append(img, body);
        return card;
    }
    function renderRebooked() {
        const box = document.getElementById('sm-rebooked');
        box.hidden = !rebooked.length;
        if (!rebooked.length) return;
        const list = document.getElementById('sm-rebooked-list');
        list.replaceChildren();
        rebooked.forEach((r, i) => {
            const row = el('div', 'sm-pair sm-rebook-row');
            const mid = el('div', 'sm-rebook-choice');
            mid.append(el('div', 'sm-card-title', r.slot));
            [['new', 'Use the new label'], ['keep', 'Keep the earlier one']].forEach(([value, text]) => {
                const on = rebook[r.new.key] === value;
                const b = el('button', on ? 'ab-btn' : 'ab-btn ab-btn-plain', `${on ? '✓ ' : ''}${text}`);
                b.type = 'button';
                b.setAttribute('aria-pressed', String(on));
                b.disabled = finished || root.dataset.running === '1';
                b.addEventListener('click', () => { rebook[r.new.key] = value; saved = false; render(); });
                mid.append(b);
            });
            row.classList.toggle('sm-pair-suggested', !rebook[r.new.key]);
            row.append(rebookCard(i, 'earlier', r.earlier, 'Earlier'), mid, rebookCard(i, 'new', r.new, 'New'));
            list.append(row);
        });
        const left = undecided().length;
        document.getElementById('sm-rebooked-count').textContent = `(${rebooked.length}${left ? `, ${left} to choose` : ''})`;
    }
    const chooseAll = (value) => { rebooked.forEach((r) => { rebook[r.new.key] = value; }); saved = false; render(); };
    document.getElementById('sm-rebook-new').addEventListener('click', () => chooseAll('new'));
    document.getElementById('sm-rebook-keep').addEventListener('click', () => chooseAll('keep'));

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
        renderRebooked();
        saveBtn.disabled = finished || (!session.unmatched.length && !rebooked.length) || root.dataset.running === '1';
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
    const post = async (url, body, progressId) => {
        const headers = { 'Content-Type': 'application/json', ...(progressId ? { 'X-Progress-Id': progressId } : {}) };
        const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
        const data = await res.json().catch(() => ({}));
        if (!res.ok) throw new Error(data.error || 'Could not save.');
        return data;
    };
    saveBtn.addEventListener('click', async () => {
        if (root.dataset.running === '1') return;
        saveBtn.disabled = true;
        try {
            const data = await post(root.dataset.saveUrl, { matches: pairs, rebook });
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
        if (undecided().length) {  // a pack booked again needs its choice first
            say(`Choose the new or the earlier courier label for ${plural(undecided().length, 'pack')} booked again first.`, true);
            document.getElementById('sm-rebooked').scrollIntoView({ behavior: motion ? 'smooth' : 'instant', block: 'start' });
            return;
        }
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
        const progressId = watch();
        say('Stitching the Complete Labels…');
        try {
            await post(root.dataset.finaliseUrl, { matches: pairs, add_unmatched: addUnmatched, rebook }, progressId);
            saved = true;
            poll();
        } catch (err) { root.dataset.running = ''; idle(); render(); say(err.message, true); }
    });
    // The finished file: Download and Open file, nothing else to do on this page
    function showResult(data) {
        const link = root.dataset.downloadUrl.replace('__FILE__', encodeURIComponent(data.output));
        status.replaceChildren(el('span', null, `${(data.without || []).length ? '⚠️' : '✅'} ${data.output}: ${plural(data.placed, 'courier label')} placed`
            + (data.built_on ? ` (built on ${data.built_on}, ${data.new} new)` : '')
            + (data.unmatched ? `, ${data.unmatched} unmatched${data.added_unmatched ? ' (4-up at the end)' : ' (left out)'}. ` : '. ')));
        const a = el('a', 'ab-btn', '⬇ Download'); a.href = link;
        const open = el('a', 'ab-btn ab-btn-plain btn-file open', '📂 Open file');
        open.href = root.dataset.openUrl.replace('__FILE__', encodeURIComponent(data.output));
        status.append(a, ' ', open);
        if ((data.locked || []).length) status.append(el('div', 'ps-hint', `Still open in another program, so not removed: ${data.locked.join(', ')}.`));
        status.classList.remove('sm-status-bad');
        // Packing labels that got no courier label: each with its page in the new file, so they can be found and booked
        const without = data.without || [];
        if (without.length) {
            const box = el('div', 'sm-missing');
            box.append(el('strong', null, `⚠ ${plural(without.length, 'packing label')} without a courier label in this file:`));
            const ul = el('ul');
            without.forEach((m) => ul.append(el('li', null, `${m.slot}${m.page ? ` · page ${m.page}` : ''}`)));
            box.append(ul, el('p', 'ps-hint', 'Book them on the courier portal, then use ➕ Add courier label PDFs and Finalise again.'));
            status.append(box);
        }
        const notInCsv = data.not_in_csv || [];
        if (notInCsv.length) status.append(el('div', 'ps-hint', `Not in the courier CSV (no complete address), so no courier label: ${notInCsv.map((m) => `${m.slot}${m.page ? ` (page ${m.page})` : ''}`).join(' · ')}.`));
        Object.keys(pairs).forEach((k) => { delete pairs[k]; delete suggested[k]; delete fromSuggestion[k]; });
        session.unmatched = []; session.waiting = []; rebooked.splice(0); finished = true;
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
            if (data.state === 'error') {
                if (data.log_url) { window.location.assign(data.log_url); return; }  // the Logs page, at what failed
                render(); say(data.error || 'Stitching failed.', true); return;
            }
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
        const progressId = watch();
        try {
            const res = await fetch(root.dataset.addUrl, { method: 'POST', body: form,
                                                          headers: progressId ? { 'X-Progress-Id': progressId } : {} });
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
    if (undecided().length) say(`${plural(undecided().length, 'pack')} got another courier label in the new PDFs: choose the new or the earlier one under Booked again, then Finalise.`);
    else if (!session.unmatched.length) say('Every courier label is matched. Nothing to do here.');
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
    // Build on: the PDFs the picked stitch already has are part of it (greyed, not ticked again)
    const base = () => form.querySelector('input[name="base"]:checked');
    const markBase = () => {
        let inBase = [];
        try { inBase = JSON.parse((base() || {}).dataset ? base().dataset.files || '[]' : '[]'); } catch (err) { inBase = []; }
        form.querySelectorAll('input[name="existing"]').forEach((box) => {
            const has = inBase.includes(box.value);
            box.disabled = has;
            if (has) box.checked = false;
            const note = box.parentElement.querySelector('.st-in-base');
            if (note) note.hidden = !has;
        });
    };
    form.querySelectorAll('input[name="base"]').forEach((r) => r.addEventListener('change', markBase));
    markBase();
    form.addEventListener('submit', (e) => {
        const files = form.querySelector('input[name="labels"]');
        const ticked = form.querySelectorAll('input[name="existing"]:checked').length;
        const onTop = base() && base().value !== 'none';  // stitching again on an earlier stitch needs no new PDF
        if (form.dataset.sent === '1' || (!(files && files.files.length) && !ticked && !onTop)) {
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

    // Merge into one and Keep separate (one receiver at different addresses): section 11, on the page at once
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
    const stores = document.getElementById('pl-export-stores');
    const show = () => {
        const abroad = [...document.querySelectorAll('.pl-con-row')].filter((r) => (r.dataset.country || 'AU') !== 'AU');
        const n = abroad.length;
        panel.hidden = !(openfreight && openfreight.checked && n > 0);
        if (count) count.textContent = `${n} consignment${n === 1 ? '' : 's'}`;
        // Which ones: each store with its number and country, a click goes to its row
        if (stores) {
            stores.replaceChildren();
            abroad.forEach((r, i) => {
                const b = document.createElement('button');
                b.type = 'button';
                b.className = 'pl-export-store';
                const num = (r.querySelector('.pl-courier-num') || {}).textContent || '';
                const name = (r.querySelector('.pl-show-receiver') || {}).textContent || '';
                b.textContent = `#${num.replace(/\D/g, '')} ${name.trim()} (${r.dataset.country})`;
                b.title = 'Go to this consignment';
                b.addEventListener('click', () => { r.scrollIntoView({ block: 'center', behavior: 'smooth' }); r.classList.remove('pl-flash'); void r.offsetWidth; r.classList.add('pl-flash'); });
                stores.append(i ? ', ' : ': ', b);
            });
        }
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

// =========================================
// 22. ADDRESS DATA (address_data.html)
// =========================================
// Official address data: check for updates, download & load one country in the background (progress polled every 2s),
// the LINZ key, a lookup tester and a check of the whole address book.
document.addEventListener('DOMContentLoaded', function () {
    const page = document.getElementById('address-data');
    if (!page) return;
    const $ = (id) => document.getElementById(id);
    const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const ICON = { true: '✓', false: '✕', null: 'ℹ' };
    const STATUS = { verified: '✅ Verified', locality: '🟢 Suburb, postcode and state agree', mismatch: '⚠️ Doesn’t match',
                     unknown: '❓ Not found', no_data: '— No data loaded' };

    // Progress of a download / load
    let polling = null;
    function showJob(job) {
        const box = $('ad-progress');
        if (!job || !job.state || job.state === 'idle') { box.hidden = true; return; }
        box.hidden = false;
        box.dataset.state = job.state;
        $('ad-step').textContent = `${job.country}: ${job.state === 'failed' ? 'Failed' : job.step || ''}`;
        $('ad-count').textContent = job.total ? `${Number(job.done).toLocaleString()} / ${Number(job.total).toLocaleString()} ${job.unit}`
                                              : (job.done ? `${Number(job.done).toLocaleString()} ${job.unit}` : '');
        $('ad-bar').style.width = job.total ? `${Math.min(100, (100 * job.done) / job.total)}%` : (job.state === 'running' ? '100%' : '0');
        $('ad-bar').classList.toggle('ad-bar-busy', job.state === 'running' && !job.total);
        $('ad-message').textContent = job.message || '';
        page.querySelectorAll('[data-load]').forEach((b) => { b.disabled = job.state === 'running'; });
    }
    async function poll() {
        try {
            const job = await (await fetch('/api/address-data/status')).json();
            showJob(job);
            if (job.state === 'running') { polling = setTimeout(poll, 2000); return; }
            if (job.state === 'done') setTimeout(() => location.reload(), 1500);  // show the new counts
        } catch (e) { polling = setTimeout(poll, 5000); }
    }
    let job = {};
    try { job = JSON.parse(page.dataset.job || '{}'); } catch (e) { job = {}; }
    showJob(job);
    if (job.state === 'running') poll();

    page.querySelectorAll('[data-load]').forEach((button) => button.addEventListener('click', async () => {
        const country = button.dataset.load;
        if (!confirm(country === 'AU' ? 'Download G-NAF (about 1.9 GB) and load it? It takes a while; the current data stays in use until it’s done.'
                                      : 'Download the LINZ NZ addresses and load them? The current data stays in use until it’s done.')) return;
        const res = await fetch('/api/address-data/load', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ country }) });
        const data = await res.json();
        if (data.error) { alert(data.error); return; }
        clearTimeout(polling); poll();
    }));

    $('ad-check-updates').addEventListener('click', async (e) => {
        const button = e.currentTarget;
        button.disabled = true; button.textContent = 'Checking…';
        try {
            const data = await (await fetch('/api/address-data/updates')).json();
            Object.entries(data).forEach(([c, r]) => {
                const el = page.querySelector(`.ad-card[data-country="${c}"] [data-latest]`);
                if (!el) return;
                el.textContent = r.error ? `Couldn't check: ${r.error}`
                    : r.newer ? `🆕 ${c === 'AU' ? `Newer release: ${r.latest} (${r.size_mb} MB)` : 'Worth refreshing: LINZ updates every week'}${r.loaded ? '' : ' (not loaded yet)'}`
                              : `✓ Up to date (${r.latest})`;
                el.classList.toggle('ad-newer', !!r.newer);
                if (c === 'NZ' && r.layer) showLayerCheck(el, r.layer);
            });
            $('ad-last-checked').textContent = 'Last checked just now';
        } catch (err) { alert('Could not check for updates.'); }
        button.disabled = false; button.textContent = '🔄 Check for updates';
    });

    $('ad-key-save').addEventListener('click', async () => {
        const res = await fetch('/api/address-data/key', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ linz_key: $('ad-linz-key').value }) });
        $('ad-key-status').textContent = (await res.json()).saved ? '✓ Saved' : 'Removed';
    });

    // NZ Addresses' LINZ layer: blank = looked up in LINZ's catalogue, a number = set by hand
    $('ad-layer-save').addEventListener('click', async () => {
        const res = await fetch('/api/address-data/layer', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ linz_layer: $('ad-linz-layer').value.trim() }) });
        const data = await res.json();
        $('ad-layer-status').textContent = data.error ? `⚠ ${data.error}`
            : data.layer ? `✓ Saved: the next NZ download uses layer ${data.layer}` : '✓ Automatic: looked up in LINZ’s catalogue';
    });

    // After "Check for updates": what LINZ lists now against the layer the loaded data came from
    function showLayerCheck(el, info) {
        const note = document.createElement('div');
        if (info.moved) {
            note.textContent = `⚠ LINZ has moved NZ Addresses: it's layer ${info.catalogue} now, your NZ data came from layer ${info.used}. `
                + (info.hand && info.hand !== info.catalogue ? `Your layer box says ${info.hand}: clear it (or type ${info.catalogue}), save, then Download & load.`
                                                             : 'Download & load to switch to it.');
            note.className = 'ad-newer';
        } else if (info.catalogue) {
            note.textContent = `✓ LINZ layer ${info.catalogue}${info.used ? ` (same as your NZ data)` : ''}`
                + (info.hand && info.hand !== info.catalogue ? `. Your layer box says ${info.hand}: clear it unless you need that one.` : '');
        } else {
            note.textContent = `Couldn't reach LINZ's catalogue to check the layer number.`;
        }
        el.appendChild(note);
    }

    function renderResult(r) {
        return `<div class="ad-status ad-status-${r.status}">${STATUS[r.status] || r.status} <span class="issue-detail">${esc(r.country)} · ${r.ms} ms</span></div>`
            + `<ul class="ad-checks">${r.checks.map((c) => `<li class="ad-check-${c.ok}"><span>${ICON[c.ok]}</span> <strong>${esc(c.what)}:</strong> ${esc(c.text)}</li>`).join('')}</ul>`
            + (r.suggestions.length ? `<div class="issue-detail">Did you mean: ${r.suggestions.map((s) => esc([s.suburb, s.state, s.postcode].filter(Boolean).join(' '))).join(' · ')}</div>` : '');
    }
    $('ad-lookup').addEventListener('submit', async (e) => {
        e.preventDefault();
        const q = new URLSearchParams(new FormData(e.target));
        const r = await (await fetch(`/api/address-data/lookup?${q}`)).json();
        $('ad-result').innerHTML = renderResult(r);
        $('ad-result').hidden = false;
    });

    $('ad-check-book').addEventListener('click', async (e) => {
        const button = e.currentTarget;
        button.disabled = true; button.textContent = 'Checking…';
        try {
            const data = await (await fetch('/api/address-data/check-book', { method: 'POST' })).json();
            const order = ['mismatch', 'unknown', 'locality', 'verified', 'no_data'];
            const rows = data.results.sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status));
            $('ad-book-result').innerHTML =
                `<p><strong>${rows.length} addresses in ${data.seconds}s</strong> (${data.average_ms} ms each): `
                + order.filter((s) => data.counts[s]).map((s) => `${STATUS[s]} ${data.counts[s]}`).join(' · ') + '</p>'
                + '<table class="ad-book-table"><thead><tr><th>Receiver</th><th>Address</th><th>Result</th></tr></thead><tbody>'
                + rows.map((r) => `<tr><td>${esc(r.receiver)}</td><td>${esc(r.address)}</td><td>${renderResult(r)}</td></tr>`).join('')
                + '</tbody></table>';
            $('ad-book-result').hidden = false;
        } catch (err) { alert('Could not check the address book.'); }
        button.disabled = false; button.textContent = 'Check all saved addresses';
    });
});

// =========================================
// 23. OFFICIAL ADDRESS DATA IN THE PREVIEW AND THE FORMS
// =========================================
// renderAddressParts: a consignment's address as one span per part (line1, line2, suburb, state, postcode, country), so
// a mismatch with the official data underlines only the part that's wrong.
// The parts the courier needs (suburb, postcode; state in Australia) missing from an address that has something in
// it: ['suburb'] for 'Home Centre, Shop 5/752 Stafford Rd, Everton Park, QLD, 4053'. Same rules as the template's
// missing_parts macro.
function missingParts(d) {
    const country = (d.country || 'AU').toUpperCase();
    if (!['line1', 'line2', 'suburb', 'postcode'].some((k) => d[k])) return [];
    return (country === 'AU' ? ['suburb', 'state', 'postcode'] : ['suburb', 'postcode']).filter((k) => !d[k]);
}
function renderAddressParts(el, d) {
    const country = (d.country || 'AU').toUpperCase();
    const missing = missingParts(d);
    const parts = ['line1', 'line2', 'suburb', 'state', 'postcode'].filter((k) => d[k] || missing.includes(k));
    parts.push('country');  // always shown, AU too, as the courier CSV carries it
    el.replaceChildren();
    parts.forEach((k, i) => {
        const span = document.createElement('span');
        span.className = 'pl-af';
        span.dataset.f = k;
        if (k === 'country') span.textContent = country;
        else if (d[k]) span.textContent = d[k];
        else {  // missing: where it belongs, in red
            span.classList.add('pl-af-missing');
            span.title = `Missing: the courier needs the ${k}`;
            span.textContent = `no ${k}`;
        }
        el.append(span);
        if (i < parts.length - 1) el.append(', ');
    });
}

// PL_REF: what the official address data says about each consignment (the server works it out with the preview and
// again after an edit). Soft on purpose:
//   mismatch  the part that differs underlined in amber; under the address the reason, "Use …" fixes and Keep as is
//             (only a grey note when the address is saved in the address book: the book's address may well be right)
//   note      the street isn't on record, in the suburb or next to it: flagged like a mismatch, with the closest
//             streets as fixes and Keep as is (a new street or centre the official data hasn't caught up with); only a
//             grey line when the address is saved in the address book
//   ok        a small ✓ after the address; a number not on record only shows when hovering the street
// One instance of each change: when the address book, the saved addresses and the official data suggest the same
// change, or one's change is part of another's (a suburb hint inside "Use EVERTON PARK QLD 4053"), only the fuller one
// is shown, with its own buttons, and says who else suggests it. Run after each part of a row's suggestions is drawn.
function dedupeSuggestions(row) {
    const norm = (v) => String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toUpperCase();
    const covers = (big, small) => Object.keys(small).length > 0 && Object.entries(small).every(([k, v]) => norm(big[k]) === norm(v));
    const items = [];
    row.querySelectorAll('.pl-book-opt').forEach((el) => { if (el._fields) items.push({ el, fields: el._fields, source: 'Address book', rank: 0 }); });
    row.querySelectorAll('.pl-book-hint').forEach((el) => { if (el._fields) items.push({ el, fields: el._fields, source: 'Saved addresses', rank: 1 }); });
    row.querySelectorAll('.pl-ref .pl-ref-use').forEach((el) => { if (el._fields) items.push({ el, fields: el._fields, source: 'Official data', rank: 2 }); });
    items.forEach((it) => { it.el.hidden = false; it.also = []; });
    row.querySelectorAll('.pl-also').forEach((x) => x.remove());
    const size = (it) => Object.keys(it.fields).length;
    items.forEach((a) => {
        const b = items.find((x) => x !== a && !x.el.hidden && covers(x.fields, a.fields) &&
            (size(x) > size(a) || (size(x) === size(a) && x.rank < a.rank)));
        if (!b) return;
        a.el.hidden = true;
        if (!b.also.includes(a.source) && a.source !== b.source) b.also.push(a.source);
    });
    items.filter((it) => it.also.length).forEach((it) => {
        const tag = document.createElement('span');
        tag.className = 'pl-also';
        tag.textContent = ` · also from ${it.also.join(' and ')}`;
        tag.title = 'The same change is suggested there too: shown once';
        if (it.el.classList.contains('pl-ref-use')) it.el.after(tag);
        else (it.el.querySelector('.pl-book-opt-text, span') || it.el).append(tag);
    });
}

const PL_REF = (() => {
    const refNote = () => document.getElementById('courier-ref-note');
    function count() {
        const note = refNote();
        if (!note) return;
        const n = [...document.querySelectorAll('.pl-con-row')].filter((r) => ['mismatch', 'note'].includes(r.dataset.refLevel) && r.dataset.refSoft !== 'true').length;
        note.hidden = !n;
        note.querySelector('strong').textContent = n;
        note.querySelector('span').textContent = n === 1 ? 'address doesn’t' : 'addresses don’t';
    }
    function button(text, cls, title, onClick) {
        const b = document.createElement('button');
        b.type = 'button'; b.className = `pl-btn-small ${cls}`; b.textContent = text; b.title = title;
        b.addEventListener('click', onClick);
        return b;
    }
    function render(row, ref) {
        ref = ref || {};
        row.dataset.refLevel = ref.level || 'none';
        row.dataset.refSoft = String(!!ref.soft);
        row.dataset.refMessage = ref.message || '';
        const address = row.querySelector('.pl-show-address');
        const box = row.querySelector('.pl-ref');
        if (!address || !box) return;
        address.querySelectorAll('.pl-af').forEach((span) => {
            span.classList.remove('pl-af-bad', 'pl-af-soft', 'pl-af-note', 'pl-af-info');
            if (!span.classList.contains('pl-af-missing')) span.removeAttribute('title');
        });
        address.querySelectorAll('.pl-ref-ok').forEach((x) => x.remove());
        box.replaceChildren();
        box.className = 'pl-ref';
        const flag = ref.level === 'mismatch' ? (ref.soft ? 'soft' : 'bad') : (ref.level === 'note' && !ref.soft ? 'bad' : null);
        Object.entries(ref.fields || {}).forEach(([field, f]) => {
            let span = address.querySelector(`.pl-af[data-f="${field}"]`) || (field === 'line1' && address.querySelector('.pl-af[data-f="line2"]'));
            if (!span && field === 'country') {  // AU isn't shown: add it, to have something to underline
                address.append(', ');
                span = document.createElement('span');
                span.className = 'pl-af'; span.dataset.f = 'country'; span.textContent = row.dataset.country || 'AU';
                address.append(span);
            }
            if (!span) return;
            const level = f.level === 'bad' ? (flag === 'soft' ? 'soft' : 'bad') : f.level;
            span.classList.add(`pl-af-${level}`);
            span.title = f.text;
        });
        if (ref.level === 'ok' || ref.level === 'verified') {
            const ok = document.createElement('span');
            ok.className = 'pl-ref-ok';
            ok.textContent = ' ✓';
            ok.title = ref.level === 'verified' ? 'Verified by the courier portal' : 'Matches the official address data';
            address.append(ok);
        }
        if (flag) {
            box.classList.add(flag === 'bad' ? 'pl-ref-bad' : 'pl-ref-soft');
            const text = document.createElement('span');
            text.className = 'pl-ref-text';
            text.textContent = (flag === 'bad' ? '⚠ ' : 'Official data: ') + ref.message;
            box.append(text);
            (ref.options || []).forEach((o) => {
                const use = button(`Use ${o.label}`, 'pl-ref-use', 'Apply this to the address (like ✏️ Save)',
                    () => row.dispatchEvent(new CustomEvent('pl-ref-use', { detail: { fields: o.fields } })));
                use._fields = o.fields || {};
                box.append(use);
            });
            box.append(button('Keep as is', 'pl-btn-plain', 'This address is right: stop flagging it',
                () => row.dispatchEvent(new CustomEvent('pl-ref-keep'))));
        } else if (ref.level === 'note') {
            box.classList.add('pl-ref-note');
            box.textContent = `ⓘ ${ref.message}`;
        } else if (ref.level === 'kept') {
            box.classList.add('pl-ref-note');
            box.textContent = 'Kept as is (official data differs)';
        }
        count();
        dedupeSuggestions(row);
    }
    let timers = new WeakMap();
    function recheck(row, d, keep) {
        clearTimeout(timers.get(row));
        timers.set(row, setTimeout(async () => {
            try {
                const res = await fetch('/api/address-data/assess', { method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ address: d, keep: keep || null }) });
                render(row, await res.json());
            } catch (e) { /* leave it as it was */ }
        }, 150));
    }
    return { render, recheck, count };
})();

document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('.pl-con-row[data-ref]').forEach((row) => {
        let ref = {};
        try { ref = JSON.parse(row.dataset.ref || '{}'); } catch (e) { ref = {}; }
        PL_REF.render(row, ref);
    });
});

// attachFieldSuggest: suggestions under an address field while it's typed in (the ✏️ form, the Address Book editor).
// Saved addresses the courier verified first, then other saved addresses (as whole addresses), then the official data
// for that field (a suburb fills suburb, state and postcode; a street fills the line; with no suburb yet, a whole
// address). ↑ ↓ to move, Enter to pick, Esc to close. Picking fills the fields and re-ranks "Closest saved addresses".
function attachFieldSuggest(container) {
    const FIELDS = ['line1', 'line2', 'suburb', 'state', 'postcode', 'country'];
    const inputs = Object.fromEntries(FIELDS.map((k) => [k, container.querySelector(`[data-field="${k}"]`)]).filter(([, el]) => el));
    const SOURCE = { verified: '✓ Verified', saved: 'Saved', official: 'Official' };
    let menu = null, items = [], active = -1, timer = null, controller = null, applying = false;

    const close = () => { if (menu) menu.remove(); menu = null; items = []; active = -1; };
    const read = () => Object.fromEntries(Object.entries(inputs).map(([k, el]) => [k, el.value.trim()]));
    // Whose address it is: a whole saved address is only suggested in the suburb or postcode field when it's theirs
    const receiverInput = container.querySelector('input[data-field="receiver"]');
    function apply(option) {
        applying = true;
        Object.entries(option.fields).forEach(([k, v]) => {
            const el = inputs[k];
            if (!el) return;
            el.value = ['suburb', 'state', 'country'].includes(k) ? String(v || '').toUpperCase() : (v || '');
            el.dispatchEvent(new Event('input', { bubbles: true }));
        });
        applying = false;
        close();
    }
    function highlight(i) {
        active = i;
        items.forEach((el, n) => el.classList.toggle('af-active', n === i));
        if (items[i]) items[i].scrollIntoView({ block: 'nearest' });
    }
    function show(input, options) {
        close();
        if (!options.length || document.activeElement !== input) return;
        menu = document.createElement('div');
        menu.className = 'af-menu';
        menu.setAttribute('role', 'listbox');
        options.forEach((o) => {
            const el = document.createElement('div');
            el.className = `af-item af-${o.source}`;
            el.setAttribute('role', 'option');
            const label = document.createElement('span'); label.className = 'af-label'; label.textContent = o.label;
            const detail = document.createElement('span'); detail.className = 'af-detail'; detail.textContent = o.detail || '';
            const source = document.createElement('span'); source.className = 'af-source'; source.textContent = SOURCE[o.source] || '';
            el.append(label, detail, source);
            el.addEventListener('mousedown', (e) => { e.preventDefault(); apply(o); });  // before the field loses focus
            menu.append(el);
            items.push(el);
        });
        const box = input.getBoundingClientRect();
        menu.style.left = `${box.left + window.scrollX}px`;
        menu.style.top = `${box.bottom + window.scrollY + 2}px`;
        menu.style.minWidth = `${Math.max(box.width, 260)}px`;
        document.body.append(menu);
        menu._options = options;
    }
    Object.entries(inputs).forEach(([field, input]) => {
        input.setAttribute('autocomplete', 'off');
        input.addEventListener('input', () => {
            if (applying) return;
            clearTimeout(timer);
            const value = input.value.trim();
            if (!value && field !== 'state' && field !== 'country') { close(); return; }
            timer = setTimeout(async () => {
                if (controller) controller.abort();
                controller = new AbortController();
                const params = new URLSearchParams({ field, value, ...read(), receiver: receiverInput ? receiverInput.value.trim() : '' });
                try {
                    const res = await fetch(`/api/address-data/complete?${params}`, { signal: controller.signal });
                    show(input, (await res.json()).options || []);
                } catch (e) { /* aborted by a newer keystroke, or offline */ }
            }, 180);
        });
        input.addEventListener('keydown', (e) => {
            if (!menu) return;
            if (e.key === 'ArrowDown') { e.preventDefault(); highlight(Math.min(items.length - 1, active + 1)); }
            else if (e.key === 'ArrowUp') { e.preventDefault(); highlight(Math.max(0, active - 1)); }
            else if (e.key === 'Enter' && active >= 0) { e.preventDefault(); e.stopPropagation(); apply(menu._options[active]); }
            else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); }
        });
        input.addEventListener('blur', () => setTimeout(close, 150));
    });
    window.addEventListener('scroll', close, { passive: true });
}

// =========================================
// 24. BACK UP & MOVE DATA (data_transfer.html)
// =========================================
// "Everything" ticks every item of its list; sending with nothing ticked says so instead of going to the server.
document.addEventListener('DOMContentLoaded', function () {
    [['dt-export', 'dt-export-all'], ['dt-import', 'dt-import-all']].forEach(([formId, allId]) => {
        const form = document.getElementById(formId);
        const all = document.getElementById(allId);
        if (!form || !all) return;
        const boxes = [...form.querySelectorAll('input[name="item"]:not(:disabled)')];
        const error = form.querySelector('.dt-error');
        all.addEventListener('change', () => boxes.forEach((b) => { b.checked = all.checked; }));
        boxes.forEach((b) => b.addEventListener('change', () => {
            all.checked = boxes.every((x) => x.checked);
            if (error) error.hidden = true;
        }));
        form.addEventListener('submit', (e) => {
            if (boxes.some((b) => b.checked)) {
                // An export is a download: the page stays, so the loading overlay mustn't stay up
                if (formId === 'dt-export') setTimeout(() => { const o = document.getElementById('loading-overlay'); if (o) o.style.display = 'none'; }, 50);
                return;
            }
            e.preventDefault();
            e.stopImmediatePropagation();
            if (error) error.hidden = false;
        }, true);
    });
});

// =========================================
// 25. ADDRESS BOOK OPTIONS AND THE "TO CHECK" LIST (packing labels preview)
// =========================================
// The preview shows each Excel address as written. What the address book knows about it (the address it learned for
// this spelling, who is at an address with no name, a name-only receiver's saved addresses) is offered under the
// address, after any problem, with one click to use it: nothing is filled in by itself.
// The "To check" list above the table gathers everything to look at, ranked in three tiers and grouped by problem (one
// line per problem, its consignments under it): Needs attention first (the file's major problems, packing specs with no
// size, no address, the official data saying a part of the address doesn't exist there), Worth a check (the other
// checks, address-book options and hints), For information (new to the address book). A click on a consignment goes
// to its row; it follows every fix.
document.addEventListener('DOMContentLoaded', function () {
    const panel = document.getElementById('pl-issues');
    const rows = [...document.querySelectorAll('.pl-con-row')];
    if (!rows.length) return;

    const button = (text, cls, title, onClick) => {
        const b = document.createElement('button');
        b.type = 'button'; b.className = `pl-btn-small ${cls}`; b.textContent = text; b.title = title;
        b.addEventListener('click', onClick);
        return b;
    };
    rows.forEach((row) => {
        // The address book's saved address for this receiver / spelling: what differs from the address shown is marked
        const box = row.querySelector('.pl-book-opts');
        if (box) {
            let options = [];
            try { options = JSON.parse(box.dataset.options || '[]'); } catch (e) { options = []; }
            options.forEach((o) => {
                const line = document.createElement('div');
                line.className = 'pl-book-opt';
                const what = document.createElement('span');
                what.className = 'pl-book-opt-text';
                const name = o.receiver && o.receiver !== row.querySelector('.pl-show-receiver').textContent ? ` · ${o.receiver}` : '';
                what.append('📇 Address book: ');
                const strong = document.createElement('strong');
                (o.parts && o.parts.length ? o.parts : [[o.label, false]]).forEach(([text, changed]) => {
                    if (!changed) { strong.append(text); return; }
                    const mark = document.createElement('mark');
                    mark.className = 'pl-diff';
                    mark.textContent = text;
                    strong.append(mark);
                });
                what.append(strong, `${name}${o.contact ? ` · Attn ${o.contact}` : ''} · ${o.why}${o.verified ? ' · ✓ Verified' : ''}`);
                line.title = o.changed && o.changed.length ? `Differs from the address shown in: ${o.changed.join(', ')}` : '';
                line._fields = o.fields || {};
                line.append(what,
                    button('Use', '', 'Put this saved address in, instead of the one shown',
                        () => row.dispatchEvent(new CustomEvent('pl-ref-use', { detail: { fields: o.fields, source: 'book' } }))),
                    button('Keep as is', 'pl-btn-plain', 'The address shown is right: stop offering this one for it',
                        () => row.dispatchEvent(new CustomEvent('pl-book-keep'))));
                box.append(line);
            });
        }
        // What the saved addresses say about the suburb: each hint changes only its own field
        const hintsBox = row.querySelector('.pl-book-hints');
        if (hintsBox) {
            let hints = [];
            try { hints = JSON.parse(hintsBox.dataset.hints || '[]'); } catch (e) { hints = []; }
            hints.forEach((h) => {
                const line = document.createElement('div');
                line.className = 'pl-book-hint';
                line.dataset.hintField = h.field;
                line._fields = { [h.field]: h.value };
                const what = document.createElement('span');
                const strong = document.createElement('strong');
                strong.textContent = h.label;
                what.append('📇 Saved addresses: ', strong, ` · ${h.why}`);
                line.append(what, button(`Use ${h.value}`, '', `Change only the ${h.field} to ${h.value}`,
                    () => row.dispatchEvent(new CustomEvent('pl-ref-use', { detail: { fields: { [h.field]: h.value } } }))));
                hintsBox.append(line);
            });
        }
        // A fix applied (an option, a hint, an official suggestion): hints about the fields it changed are done with
        row.addEventListener('pl-ref-use', (e) => {
            const changed = Object.keys(e.detail.fields || {});
            if (hintsBox) hintsBox.querySelectorAll('.pl-book-hint').forEach((l) => { if (changed.includes(l.dataset.hintField) || e.detail.source === 'book') l.remove(); });
            dedupeSuggestions(row);
        });
        dedupeSuggestions(row);
    });

    if (!panel) return;
    // The consignment table folds away behind a green bar when everything checks out (the first time the list is
    // worked out); with something to check it starts open. Its button opens or closes it any time.
    const table = document.getElementById('pl-con-table');
    const bar = document.getElementById('pl-con-bar');
    const toggle = document.getElementById('pl-con-toggle');
    const stateText = document.getElementById('pl-con-state');
    let decided = false;
    let progress = { left: 0, total: 0 };
    let syncFloating = () => {};
    function showTable(open) {
        if (!table) return;
        table.hidden = !open;
        toggle.textContent = open ? '▲ Hide consignments' : '▼ Show consignments';
        toggle.setAttribute('aria-expanded', String(open));
        syncFloating();
    }
    if (toggle) toggle.addEventListener('click', () => showTable(table.hidden));
    // The same Hide on the floating Open Excel buttons (section 26), while the table is open and its own button has
    // scrolled out of sight: one click folds the table and brings its bar back into view
    const fx = document.getElementById('fx-open');
    if (fx && table && toggle) {
        const hide = document.createElement('button');
        hide.type = 'button';
        hide.className = 'fx-hide-btn';
        hide.textContent = '▲ Hide consignments';
        hide.title = 'Fold the consignment table away';
        hide.hidden = true;
        hide.addEventListener('click', () => { showTable(false); bar.scrollIntoView({ block: 'center' }); });
        fx.append(hide);
        syncFloating = () => {
            const past = toggle.getBoundingClientRect().bottom < 0;
            const show = !table.hidden && past && table.getBoundingClientRect().bottom > 80;
            if (hide.hidden === show) {
                hide.hidden = !show;
                window.dispatchEvent(new Event('resize'));  // the floating box re-checks what it might cover
            }
        };
        window.addEventListener('scroll', syncFloating, { passive: true });
    }
    // The same count on the floating Open Excel box, in sight wherever the page is scrolled; a click goes to the To
    // check list (or to the table's bar when there's nothing left in it)
    let fxState = null;
    if (fx && bar) {
        fxState = document.createElement('button');
        fxState.type = 'button';
        fxState.className = 'fx-con-state';
        fxState.title = 'Go to the To check list';
        fxState.addEventListener('click', () => (panel.hidden ? bar : panel).scrollIntoView({ block: 'start', behavior: 'smooth' }));
        fx.insertBefore(fxState, fx.querySelector('.fx-hide-btn'));  // before Hide consignments, else at the end
    }
    function tableState(clear) {
        if (!bar) return;
        const n = rows.length, s = n === 1 ? '' : 's';
        const done = clear || (progress.total > 0 && !progress.left);
        bar.classList.toggle('pl-con-clear', done);
        stateText.textContent = clear ? `✓ All ${n} consignment${s} check out` + (progress.total ? ` (${progress.total} fixed)` : '')
            : progress.left ? `${n} consignment${s}: ${progress.left} of ${progress.total} still to check (see To check above)`
            : `✓ All ${n} consignment${s} check out · the file still has something in To check above`;
        if (fxState) {
            const text = clear ? `✓ All ${n} consignment${s} check out`
                : progress.left ? `${n} consignment${s}: ${progress.left} of ${progress.total} still to check`
                : `✓ All ${n} consignment${s} check out · the file has something to check`;
            fxState.classList.toggle('fx-con-clear', done);
            if (fxState.textContent !== text) {
                fxState.textContent = text;
                window.dispatchEvent(new Event('resize'));  // the floating box re-checks what it might cover
            }
        }
        if (!decided) { decided = true; showTable(!clear); }
    }
    const list = panel.querySelector('.pl-issues-list');
    const summary = panel.querySelector('.pl-issues-summary');
    let review = {};
    try { review = JSON.parse(panel.dataset.review || '{}'); } catch (e) { review = {}; }
    const numberOf = (row) => row.querySelector('.pl-courier-num').textContent.replace(/\D/g, '');  // the number, not the ✏️ button
    const nameOf = (row) => row.querySelector('.pl-show-receiver').textContent;
    const refOf = (row) => ({ level: row.dataset.refLevel, soft: row.dataset.refSoft === 'true', message: row.dataset.refMessage || '' });
    const SHOWN = 6;  // consignments listed under a problem before "Show all"
    // Progress: every consignment that has had something to fix or check since the page opened, and the one just
    // acted on (Use, Keep as is, Save): once its problems are gone the page moves on to the next one still to check
    const ever = new Set();
    let acted = null, actedAt = 0;
    const markActed = (e) => {
        const row = e.target.closest && e.target.closest('.pl-con-row');
        if (row) { acted = row; actedAt = Date.now(); }
    };
    ['pl-ref-use', 'pl-ref-keep', 'pl-book-keep'].forEach((t) => document.addEventListener(t, markActed, true));
    document.addEventListener('click', (e) => { if (e.target.closest('.pl-edit-save')) markActed(e); }, true);
    const expanded = new Set();  // problems whose consignments are all shown, kept when the list is rebuilt after a fix
    let infoOpen = false;

    function goTo(row) {
        showTable(true);
        if (row.hidden) {  // hidden by the state or text filter: show everything again
            const state = document.getElementById('courier-state-filter'), text = document.getElementById('courier-text-filter');
            if (state) { state.value = ''; state.dispatchEvent(new Event('change')); }
            if (text) { text.value = ''; text.dispatchEvent(new Event('input')); }
        }
        row.scrollIntoView({ block: 'center', behavior: 'smooth' });
        row.classList.remove('pl-flash');
        void row.offsetWidth;  // restart the flash
        row.classList.add('pl-flash');
    }
    // A consignment under a problem: "#12 Macpac Perth" and what's wrong with it; a click goes to its row
    function entry(row, text, edit) {
        const b = document.createElement('button');
        b.type = 'button';
        b.className = 'pl-issue-go';
        const who = document.createElement('strong');
        who.textContent = `#${numberOf(row)} ${nameOf(row)}`;
        b.append(who);
        if (text) {
            const span = document.createElement('span');
            span.className = 'pl-issue-text';
            span.textContent = text;
            b.append(span);
        }
        b.addEventListener('click', () => {
            goTo(row);
            const pencil = row.querySelector('.pl-edit-btn');
            if (edit && pencil && !row.classList.contains('pl-editing')) pencil.click();
        });
        b.title = edit ? 'Go to it and open it for editing' : 'Go to it';
        return b;
    }
    // One problem: the link to fix it (a packing spec's row), what's wrong and how many, why it matters, where in the
    // file, and the consignments it's on
    function group(g) {
        const li = document.createElement('li');
        li.className = 'pl-issue-group';
        const head = document.createElement('div');
        head.className = 'pl-issue-title';
        if (g.link) {
            const link = document.createElement('a');
            link.href = g.link.href;
            link.target = '_blank';
            link.rel = 'noopener';
            link.className = 'pl-issue-link';
            link.textContent = g.link.text;
            head.append(link, ' ');
        }
        head.append(g.title);
        const entries = g.entries || [];
        if (entries.length > 1) {
            const n = document.createElement('span');
            n.className = 'pl-issue-count';
            n.textContent = entries.length;
            head.append(' ', n);
        }
        li.append(head);
        if (g.detail) {
            const d = document.createElement('div');
            d.className = 'pl-issue-detail';
            d.textContent = g.detail;
            li.append(d);
        }
        if (g.where && g.where.length) {
            const w = document.createElement('div');
            w.className = 'pl-issue-where';
            w.textContent = g.where.join(' · ');
            li.append(w);
        }
        if (entries.length) {
            const box = document.createElement('div');
            box.className = 'pl-issue-entries';
            const all = expanded.has(g.key);
            entries.forEach(([row, text], i) => {
                const b = entry(row, text, g.edit);
                b.hidden = !all && i >= SHOWN;
                box.append(b);
            });
            if (!all && entries.length > SHOWN) {
                const more = document.createElement('button');
                more.type = 'button';
                more.className = 'pl-issue-more';
                more.textContent = `Show all ${entries.length}`;
                more.addEventListener('click', () => {
                    expanded.add(g.key);
                    box.querySelectorAll('.pl-issue-go[hidden]').forEach((x) => { x.hidden = false; });
                    more.remove();
                });
                box.append(more);
            }
            li.append(box);
        }
        return li;
    }
    // Every problem, in the order it needs fixing: first what makes labels or consignments come out wrong or go
    // missing (the file's own problems, packing specs with no size, no address, the official data saying a part of
    // the address doesn't exist there), then what's worth a look, then what's only for information
    function problems() {
        const tiers = { fix: [], check: [], info: [] };
        const tabs = new Set((review.file || []).map((g) => g.tab));
        (review.file || []).forEach((g) => tiers[g.major ? 'fix' : 'check'].push({
            key: `file:${g.tab}:${g.text}`, detail: g.detail, where: g.where,
            title: `${g.text}${g.count > 1 ? ` × ${g.count}` : ''}${tabs.size > 1 ? ` (tab ${g.tab})` : ''}` }));
        (review.specs || []).forEach((sp) => tiers.fix.push({
            key: `spec:${sp.spec}`,
            title: sp.unread ? `Couldn't read the size from packing spec ${sp.spec}` : `No size saved for packing spec ${sp.spec}`,
            detail: sp.unread
                ? `${sp.cartons} carton${sp.cartons === 1 ? '' : 's'}: its digits don't split into three sizes in mm (2-4 digits each, none starting with 0, no side of 1000 cm or more), so dimensions are left blank in the courier CSV.`
                  + `${sp.meant ? ` Most likely ${sp.meant}, with a 0 typed twice.` : ''} Fix the code in the Excel file, then press 🔄 Update Previews.`
                : `${sp.cartons} carton${sp.cartons === 1 ? '' : 's'}: dimensions are left blank in the courier CSV. Fill it in, then press 🔄 Update Previews.`,
            link: { href: sp.url, text: sp.unread ? '📐 Check the code in Packing Specs' : '📐 Add its size in Packing Specs' },
            entries: sp.consignments.map((n) => rows.find((r) => numberOf(r) === String(n))).filter(Boolean).map((r) => [r, '']) }));
        (review.courier || []).forEach((w) => tiers[w.major ? 'fix' : 'check'].push({ key: `courier:${w.text}`, title: w.text, edit: true,
            entries: (w.numbers || []).map((n) => rows.find((r) => numberOf(r) === String(n))).filter(Boolean).map((r) => [r, '']) }));
        // Packs collected rather than sent ('Kyle to collect: 6 labels', 'Pack 1: 2 · Pack 2: 4'): no courier label by
        // design, so counted here for information and never as 'No address'
        (review.collect || []).forEach((w) => tiers.info.push({ key: `collect:${w.text}`, title: w.text,
            detail: `${w.detail ? `${w.detail}. ` : ''}No courier label by design: each label says who collects where the courier label would go.`,
            entries: (w.numbers || []).map((n) => rows.find((r) => numberOf(r) === String(n))).filter(Boolean).map((r) => [r, '']) }));
        // edit: a click on one of its consignments also opens it for editing (the fix is typing, not a suggestion)
        const rowGroup = (tier, key, title, detail, found, edit) => { if (found.length) tiers[tier].push({ key, title, detail, entries: found, edit }); };
        const withRef = (test) => rows.filter((r) => test(refOf(r))).map((r) => [r, refOf(r).message]);
        rowGroup('fix', 'missing', 'No address', "Left out of the courier CSV until it has one: pick a saved address or type one in with ✏️.",
            rows.filter((r) => r.classList.contains('pl-no-address')).map((r) => [r, '']), true);
        rowGroup('fix', 'incomplete', 'Suburb, state or postcode missing',
            "The courier needs all of them. Use the suggestion under the address, or fill them in here.",
            rows.filter((r) => r.dataset.missing && !r.classList.contains('pl-no-address'))
                .map((r) => [r, r.dataset.missing.split(',').map((k) => `no ${k}`).join(', ')]), true);
        rowGroup('fix', 'official', "Doesn't match the official data (G-NAF / LINZ)",
            "A suburb, state, postcode or country that isn't on record, or isn't where the address puts it. Use the suggestion under the address, or Keep as is if you know it's right.",
            withRef((x) => x.level === 'mismatch' && !x.soft).filter(([r]) => !r.dataset.missing));
        rowGroup('fix', 'street', 'Street not on record',
            "The official data has no street of that name in the suburb or next to it. Check the street and suburb; Keep as is if it's a new street.",
            withRef((x) => x.level === 'note' && !x.soft));
        rowGroup('check', 'official-soft', 'Saved address differs from the official data', '', withRef((x) => x.level === 'mismatch' && x.soft));
        rowGroup('check', 'street-soft', 'Saved address: street not on record', '', withRef((x) => x.level === 'note' && x.soft));
        rowGroup('check', 'option', 'The address book has another address for these', 'Use it, or Keep as is.',
            rows.filter((r) => r.querySelector('.pl-book-opt')).map((r) => [r, [...r.querySelectorAll('.pl-book-opt strong')].map((x) => x.textContent).join(' | ')]));
        rowGroup('check', 'hint', 'Saved addresses can fill in a part of these', '',
            rows.filter((r) => r.querySelector('.pl-book-hint')).map((r) => [r, [...r.querySelectorAll('.pl-book-hint strong')].map((x) => x.textContent).join(', ')]));
        rowGroup('info', 'book', 'Not in the address book yet', "Nothing to fix: they're saved to the address book when you generate.",
            rows.filter((r) => r.classList.contains('pl-needs-check')).map((r) => [r, '']));
        return tiers;
    }
    function refresh() {
        const tiers = problems();
        list.replaceChildren();
        [['fix', 'Needs attention first', 'These make labels or consignments come out wrong or go missing.'],
         ['check', 'Worth a check', ''], ['info', 'For information', '']].forEach(([tier, name, why]) => {
            if (!tiers[tier].length) return;
            const section = document.createElement(tier === 'info' ? 'details' : 'section');
            section.className = `pl-tier pl-tier-${tier}`;
            const head = document.createElement(tier === 'info' ? 'summary' : 'div');
            head.className = 'pl-tier-head';
            head.textContent = name;
            if (why) {
                const w = document.createElement('span');
                w.className = 'pl-tier-why';
                w.textContent = why;
                head.append(' ', w);
            }
            const ol = document.createElement('ol');
            ol.className = 'pl-issue-groups';
            tiers[tier].forEach((g) => ol.append(group(g)));
            section.append(head, ol);
            if (tier === 'info') {
                section.open = infoOpen;
                section.addEventListener('toggle', () => { infoOpen = section.open; });
            }
            list.append(section);
        });
        // The rows with a problem that needs attention first are marked red in the table too (only when that changes,
        // so the table watcher below isn't set off again)
        const urgent = new Set(tiers.fix.flatMap((g) => (g.entries || []).map(([r]) => r)));
        const toCheck = new Set([...tiers.fix, ...tiers.check].flatMap((g) => (g.entries || []).map(([r]) => r)));
        toCheck.forEach((r) => ever.add(r));
        rows.forEach((r) => { if (r.classList.contains('pl-row-major') !== urgent.has(r)) r.classList.toggle('pl-row-major', urgent.has(r)); });
        const newToBook = tiers.info.length ? tiers.info[0].entries.length : 0;
        summary.textContent = [[tiers.fix.length, tiers.fix.length === 1 ? 'problem needs attention first' : 'problems need attention first'],
                               [tiers.check.length, 'to check'], [newToBook, 'new to the address book']]
            .filter(([n]) => n).map(([n, w]) => `${n} ${w}`).join(' · ');
        if (ever.size) {  // how far along: consignments still to check out of all that needed it
            const done = document.createElement('span');
            done.className = 'pl-issues-progress';
            done.textContent = toCheck.size ? `${toCheck.size} of ${ever.size} consignment${ever.size === 1 ? '' : 's'} still to check`
                : `✓ all ${ever.size} consignment${ever.size === 1 ? '' : 's'} checked`;
            summary.append(' · ', done);
        }
        progress = { left: toCheck.size, total: ever.size };
        // The consignment just fixed has nothing left: on to the next one still to check (after it, else the first),
        // or to the table's bar when it was the last. Not while something is being typed in a row.
        if (acted && !toCheck.has(acted) && Date.now() - actedAt < 6000 && !document.querySelector('.pl-con-row.pl-editing')) {
            const order = rows.filter((r) => toCheck.has(r) && !r.hidden);
            const next = order.find((r) => acted.compareDocumentPosition(r) & Node.DOCUMENT_POSITION_FOLLOWING) || order[0];
            acted = null;
            if (next) goTo(next);
            else if (bar) bar.scrollIntoView({ block: 'center', behavior: 'smooth' });
        }
        panel.classList.toggle('pl-issues-urgent', tiers.fix.length > 0);
        panel.hidden = !list.children.length;
        tableState(!tiers.fix.length && !tiers.check.length);  // 'not in the address book yet' isn't a problem
    }
    document.addEventListener('pl-row-removed', (e) => {
        const i = rows.indexOf(e.detail.row);
        if (i >= 0) rows.splice(i, 1);
        ever.delete(e.detail.row);
    });
    document.addEventListener('pl-kept-separate', (e) => {
        const norm = (t) => String(t || '').toLowerCase().replace(/[^0-9a-z]+/g, ' ').trim().replace(/\s+/g, ' ');
        const name = norm(e.detail.receiver);
        review.file = (review.file || []).filter((g) => !(norm(g.text).startsWith(`${name} store address`) ||
                                                          norm(g.text).startsWith(`${name} installer addresses`)));
        soon();
    });
    let timer = null;
    const soon = () => { clearTimeout(timer); timer = setTimeout(refresh, 60); };
    // Every change to a row (a fix, Keep as is, a Save, an option used) updates the list
    new MutationObserver(soon).observe(document.querySelector('.pl-courier-table tbody'),
        { subtree: true, childList: true, attributes: true, attributeFilter: ['class', 'data-ref-level', 'data-ref-soft', 'data-missing'] });
    refresh();
});


// =========================================
// 27. LABEL ORDER (horizontal labels): group by allocation or not
// =========================================
// Grouped by allocation (code A, B, C …), the divider choice follows; not grouped, there are no code groups and so no
// dividers to choose, and the hint says what the PDF will hold.
document.addEventListener('DOMContentLoaded', function () {
    const box = document.getElementById('pl-grouping');
    if (!box) return;
    const dividers = box.querySelector('.pl-divider-options');
    const hint = box.querySelector('.pl-csv-hint');
    const update = () => {
        const grouped = box.querySelector('input[name="label_grouping"]:checked').value !== 'store';
        dividers.hidden = !grouped;
        hint.textContent = grouped ? hint.dataset.code : hint.dataset.store;
    };
    box.querySelectorAll('input[name="label_grouping"]').forEach((r) => r.addEventListener('change', update));
    update();
});

// =========================================
// 26. THE FLOATING "OPEN EXCEL" BUTTON (_open_excel.html)
// =========================================
// One per page that previews an uploaded file. It sits at the bottom right, or where it was last dragged to (kept in
// this browser), and when it would cover a button, box, link or checkbox it slides along that edge to the nearest
// free spot, so it never blocks anything. Checked when the page scrolls, resizes or something is clicked; no timer.
document.addEventListener('DOMContentLoaded', function () {
    const box = document.getElementById('fx-open');
    if (!box) return;
    const grip = box.querySelector('.fx-open-grip');
    const KEY = 'fx-open-spot';
    const GAP = 16;
    const CONTROLS = 'button, a, input, select, textarea, label, summary, [role="button"], [contenteditable="true"], .dropzone';
    let home = null;  // where it was dragged to, as fractions of the window, so it stays sensible after a resize
    try { home = JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { home = null; }
    let dragging = false;
    box.hidden = false;

    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
    function wanted(r) {
        const maxX = window.innerWidth - r.width - GAP, maxY = window.innerHeight - r.height - GAP;
        return home ? { x: clamp(home.x * window.innerWidth, GAP, maxX), y: clamp(home.y * window.innerHeight, GAP, maxY) }
                    : { x: maxX, y: maxY };
    }
    // A control under this spot? Six points across it; the button itself is left out while looking
    function covers(x, y, w, h) {
        for (const fx of [0.1, 0.5, 0.9]) {
            for (const fy of [0.2, 0.8]) {
                const hit = document.elementsFromPoint(x + w * fx, y + h * fy).find((el) => !box.contains(el));
                if (hit && hit.closest(CONTROLS)) return true;
            }
        }
        return false;
    }
    function place() {
        if (dragging) return;
        const r = box.getBoundingClientRect();
        const { x, y } = wanted(r);
        box.style.pointerEvents = 'none';
        let top = y;
        if (covers(x, y, r.width, r.height)) {
            const highest = GAP, lowest = window.innerHeight - r.height - GAP;
            top = null;
            for (let d = 10; d <= window.innerHeight && top === null; d += 10) {
                if (y - d >= highest && !covers(x, y - d, r.width, r.height)) top = y - d;
                else if (y + d <= lowest && !covers(x, y + d, r.width, r.height)) top = y + d;
            }
            if (top === null) top = y;  // nothing free along this edge: it stays, and can be dragged aside
        }
        box.style.pointerEvents = '';
        box.style.left = `${x}px`;
        box.style.top = `${top}px`;
    }
    let frame = 0;
    const soon = () => { cancelAnimationFrame(frame); frame = requestAnimationFrame(place); };
    window.addEventListener('scroll', soon, { passive: true });
    window.addEventListener('resize', soon);
    document.addEventListener('click', () => setTimeout(soon, 50));  // after a click has opened or closed something
    place();

    // Dragging by the handle; double-click puts it back in the corner
    grip.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        dragging = true;
        box.classList.add('fx-dragging');
        grip.setPointerCapture(e.pointerId);
        const r = box.getBoundingClientRect();
        const dx = e.clientX - r.left, dy = e.clientY - r.top;
        const move = (ev) => {
            box.style.left = `${clamp(ev.clientX - dx, 0, window.innerWidth - r.width)}px`;
            box.style.top = `${clamp(ev.clientY - dy, 0, window.innerHeight - r.height)}px`;
        };
        const up = () => {
            grip.removeEventListener('pointermove', move);
            grip.removeEventListener('pointerup', up);
            grip.removeEventListener('pointercancel', up);
            dragging = false;
            box.classList.remove('fx-dragging');
            const now = box.getBoundingClientRect();
            home = { x: now.left / window.innerWidth, y: now.top / window.innerHeight };
            try { localStorage.setItem(KEY, JSON.stringify(home)); } catch (err) { /* kept for this page only */ }
            place();
        };
        grip.addEventListener('pointermove', move);
        grip.addEventListener('pointerup', up);
        grip.addEventListener('pointercancel', up);
    });
    grip.addEventListener('dblclick', () => {
        home = null;
        try { localStorage.removeItem(KEY); } catch (err) { /* nothing kept */ }
        place();
    });
});

// =========================================
// 28. COURIER SERVICES (Address Book page) and "+ Add service" (consignment preview)
// =========================================
// The service list lives in data/courier_services.json. Every change goes to /api/courier-services… and the page is
// redrawn from the list the server sends back, so it always shows what is saved. Built with textContent only.
const CourierServices = {
    post: async (url, body, isForm) => {
        const r = await fetch(url, isForm ? { method: 'POST', body } :
            { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) });
        const data = await r.json().catch(() => ({ error: 'The app did not answer. Is it still running?' }));
        if (!r.ok || data.error) throw new Error(data.error || 'Not saved.');
        return data;
    },
    el: (tag, cls, text) => {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (text != null) e.textContent = text;
        return e;
    },
};

document.addEventListener('DOMContentLoaded', () => {
    const section = document.getElementById('courier-services');
    if (!section) return;
    const box = document.getElementById('cs-groups');
    const status = document.getElementById('cs-status');
    const { el, post } = CourierServices;
    let groups = [];
    const say = (text, bad) => { status.textContent = text; status.classList.toggle('cs-bad', !!bad); };
    const run = async (fn, done) => {
        try {
            const data = await fn();
            groups = data.groups;
            draw();
            if (done) say(done);
            return true;
        } catch (e) { say(e.message, true); return false; }
    };

    // The groups a service can be in; 'On its own' takes it out of its group (it then has no logo until one is dropped)
    function groupSelect(current) {
        const s = el('select', 'cs-input');
        groups.forEach((g) => { const o = el('option', null, g.name); o.value = g.id; o.selected = g.id === current; s.append(o); });
        const own = el('option', null, 'On its own (a group of one)');
        own.value = '__own__';
        s.append(own);
        return s;
    }

    function serviceRow(g, s) {
        const tr = el('tr', 'cs-row');
        tr.append(el('td', 'cs-code', s.code), el('td', null, s.name));
        const act = el('td', 'cs-actions');
        const edit = el('button', 'ab-icon-btn', '✏️'); edit.type = 'button'; edit.title = `Edit ${s.code}`;
        const del = el('button', 'ab-icon-btn', '🗑️'); del.type = 'button'; del.title = `Delete ${s.code}`;
        del.addEventListener('click', () => {
            if (!confirm(`Delete ${s.code} · ${s.name}? It won't be offered in the preview any more.`)) return;
            run(() => post(`/api/courier-services/${encodeURIComponent(s.code)}/delete`), `${s.code} deleted.`);
        });
        edit.addEventListener('click', () => {
            const form = el('tr', 'cs-edit');
            const td = el('td'); td.colSpan = 3;
            const code = el('input', 'cs-input cs-input-code'); code.value = s.code; code.maxLength = 20; code.setAttribute('aria-label', 'Service code');
            const name = el('input', 'cs-input'); name.value = s.name; name.maxLength = 80; name.setAttribute('aria-label', 'Service name');
            const grp = groupSelect(g.id); grp.setAttribute('aria-label', 'Move to group');
            const grpLabel = el('label', 'cs-move', 'Move to group ');
            grpLabel.append(grp);
            const save = el('button', 'ab-btn cs-small', 'Save'); save.type = 'button';
            const cancel = el('button', 'ab-btn ab-btn-plain cs-small', 'Cancel'); cancel.type = 'button';
            save.addEventListener('click', () => {
                const moved = grp.value !== g.id;
                const to = grp.value === '__own__' ? 'on its own' : `in ${grp.selectedOptions[0].textContent}`;
                run(() => post(`/api/courier-services/${encodeURIComponent(s.code)}/update`,
                    { code: code.value, name: name.value, group: grp.value }),
                    moved ? `${code.value.toUpperCase()} is now ${to}: its labels get that group's logo.` : `${code.value.toUpperCase()} saved.`);
            });
            cancel.addEventListener('click', () => { form.replaceWith(tr); });
            td.append(code, name, grpLabel, save, cancel);
            form.append(td);
            tr.replaceWith(form);
            name.focus();
        });
        act.append(edit, del);
        tr.append(act);
        return tr;
    }

    function logoArea(g) {
        const area = el('label', 'cs-logo' + (g.logo_url ? ' cs-logo-has' : ''));
        area.title = 'Drop a PNG logo strip here (about 6.3 : 1, e.g. 1000 × 160 px), or click to choose one';
        if (g.logo_url) {
            const img = el('img'); img.src = g.logo_url; img.alt = `${g.name} logo`;
            area.append(img);
        } else {
            area.append(el('span', 'cs-logo-empty', 'Drop logo here'));
        }
        const input = el('input'); input.type = 'file'; input.accept = '.png,.jpg,.jpeg,.webp'; input.hidden = true;
        const send = (file) => {
            if (!file) return;
            const fd = new FormData(); fd.append('logo', file);
            run(() => post(`/api/courier-groups/${encodeURIComponent(g.id)}/logo`, fd, true), `Logo saved for ${g.name}.`);
        };
        input.addEventListener('change', () => send(input.files[0]));
        area.addEventListener('dragover', (e) => { e.preventDefault(); area.classList.add('cs-logo-over'); });
        area.addEventListener('dragleave', () => area.classList.remove('cs-logo-over'));
        area.addEventListener('drop', (e) => { e.preventDefault(); area.classList.remove('cs-logo-over'); send(e.dataTransfer.files[0]); });
        area.append(input);
        return area;
    }

    function groupCard(g) {
        const card = el('article', 'cs-group');
        const head = el('div', 'cs-group-head');
        const title = el('h3', 'cs-group-name', g.name);
        const rename = el('button', 'ab-icon-btn', '✏️'); rename.type = 'button'; rename.title = `Rename ${g.name}`;
        rename.addEventListener('click', () => {
            const name = prompt('Group name', g.name);
            if (name && name.trim() !== g.name) run(() => post(`/api/courier-groups/${encodeURIComponent(g.id)}/rename`, { name }), 'Group renamed.');
        });
        const del = el('button', 'ab-icon-btn', '🗑️'); del.type = 'button'; del.title = `Delete the group ${g.name}`;
        del.addEventListener('click', () => {
            if (!confirm(`Delete the group ${g.name} and its logo? Its services stay in the list, each on its own.`)) return;
            run(() => post(`/api/courier-groups/${encodeURIComponent(g.id)}/delete`), `${g.name} deleted.`);
        });
        head.append(title, rename, del);
        const tools = el('div', 'cs-logo-tools');
        if (g.logo_url) {
            const rm = el('button', 'ab-btn ab-btn-plain cs-small', 'Remove logo'); rm.type = 'button';
            rm.addEventListener('click', () => run(() => post(`/api/courier-groups/${encodeURIComponent(g.id)}/logo/delete`), 'Logo removed.'));
            tools.append(rm);
        }
        const table = el('table', 'cs-table');
        const body = el('tbody');
        g.services.forEach((s) => body.append(serviceRow(g, s)));
        table.append(body);
        const add = el('div', 'cs-add');
        const code = el('input', 'cs-input cs-input-code'); code.placeholder = 'Code'; code.maxLength = 20; code.setAttribute('aria-label', 'New service code');
        const name = el('input', 'cs-input'); name.placeholder = 'Service name'; name.maxLength = 80; name.setAttribute('aria-label', 'New service name');
        const btn = el('button', 'ab-btn ab-btn-plain cs-small', '+ Add'); btn.type = 'button';
        // A code already in another group moves here (its name kept unless one is typed)
        btn.addEventListener('click', async () => {
            let result = null;
            await run(async () => { const d = await post('/api/courier-services', { code: code.value, name: name.value, group: g.id, move: true }); result = d.result; return d; });
            if (result) say(result.moved_from ? `${result.code} moved from ${result.moved_from} to ${g.name}.` : `${result.code} added to ${g.name}.`);
        });
        add.append(code, name, btn);
        card.append(head, logoArea(g), tools, table, add);
        return card;
    }

    function draw() {
        box.replaceChildren(...groups.map(groupCard));
    }

    document.getElementById('cs-add-group').addEventListener('click', async () => {
        const input = document.getElementById('cs-new-group');
        const name = input.value.trim();
        if (await run(() => post('/api/courier-groups', { name }), `Group ${name} added: add its services in its card (at the end).`)) input.value = '';
    });
    run(async () => (await fetch(section.dataset.url)).json());
});

// "+ Add service" beside the consignment preview's service list: saves a new service code into the list and offers
// it in every service dropdown on the page straight away (the main one picks it).
document.addEventListener('DOMContentLoaded', () => {
    const open = document.getElementById('pl-add-service');
    const form = document.getElementById('pl-add-service-form');
    if (!open || !form) return;
    const { post } = CourierServices;
    const msg = form.querySelector('.pl-add-service-msg');
    open.addEventListener('click', () => {
        form.hidden = !form.hidden;
        if (!form.hidden) form.querySelector('[name=new_service_code]').focus();
    });
    form.querySelector('.pl-add-service-save').addEventListener('click', async () => {
        const code = form.querySelector('[name=new_service_code]').value;
        const name = form.querySelector('[name=new_service_name]').value;
        const group = form.querySelector('[name=new_service_group]').value;
        try {
            const data = await post('/api/courier-services', { code, name, group: group || null, new_group: group ? '' : name });
            const s = data.result;
            document.querySelectorAll('#service_code, select.pl-service-input').forEach((sel) => {
                const o = document.createElement('option');
                o.value = s.code;
                o.textContent = sel.id === 'service_code' ? `${s.code} · ${s.name}` : `${s.code} · ${s.name}`;
                if (sel.id === 'service_code') { sel.append(o); sel.value = s.code; sel.dispatchEvent(new Event('change', { bubbles: true })); }
                else sel.append(o);
            });
            msg.textContent = `${s.code} added to the list (group ${s.group_name}). It's picked as the service; Apply to shown sets it on the consignments.`;
            msg.classList.remove('cs-bad');
            form.querySelectorAll('input').forEach((i) => { i.value = ''; });
        } catch (e) {
            msg.textContent = e.message;
            msg.classList.add('cs-bad');
        }
    });
});

// =========================================
// 29. GENERATE LABELS FROM A PACKING SHEETS PROJECT (horizontal labels, step 1)
// =========================================
// Picking a project lists its Final files (the Final file and any Stage files Sub-Group made: one is picked) and says
// which Signature Links file comes with it. Built with textContent only.
document.addEventListener('DOMContentLoaded', () => {
    const pick = document.getElementById('pl-project-pick');
    if (!pick) return;
    const box = document.getElementById('pl-project-finals');
    const sig = document.getElementById('pl-project-sig');
    const load = document.getElementById('pl-project-load');
    pick.addEventListener('change', () => {
        box.replaceChildren();
        const opt = pick.selectedOptions[0];
        let finals = [];
        try { finals = JSON.parse(opt.dataset.finals || '[]'); } catch (e) { finals = []; }
        finals.forEach((f, i) => {
            const label = document.createElement('label');
            label.className = 'pl-csv-option';
            const input = document.createElement('input');
            input.type = 'radio'; input.name = 'final_file'; input.value = f; input.checked = i === 0; input.required = true;
            const btn = document.createElement('span');
            btn.className = 'pl-csv-btn';
            const tick = document.createElement('span');
            tick.className = 'pl-csv-tick'; tick.setAttribute('aria-hidden', 'true'); tick.textContent = '✓';
            const name = document.createElement('strong');
            name.textContent = f;
            btn.append(tick, name);
            label.append(input, btn);
            box.append(label);
        });
        sig.textContent = !pick.value ? '' : opt.dataset.signature
            ? `With ${opt.dataset.signature}: its codes group the labels and its Label Plugin fills in the tabs.`
            : 'No Signature Links file in this project: the codes are worked out from the quantities.';
        load.disabled = !pick.value;
    });
});


// =========================================
// LOGS: what the app is doing and has done (log_view.py), newest first, live
// =========================================
document.addEventListener('DOMContentLoaded', function () {
    const root = document.getElementById('logs-page');
    if (!root) return;
    const $ = id => document.getElementById(id);
    const list = $('lg-list'), more = $('lg-more'), empty = $('lg-empty'), live = $('lg-live');
    const search = $('lg-search'), area = $('lg-area'), show = $('lg-show'), day = $('lg-day');
    const open = new Set();       // runs opened by the user stay open when the list refreshes
    let version = '', shown = 0, timer = null, busy = false, lastData = null;
    // Sent here by a page that failed (/logs?at=…&find=…&back=…): that run opened and marked, live paused so it stays put
    const sent = new URLSearchParams(window.location.search);
    let focusAt = sent.get('at'), focusFind = sent.get('find') || '';
    if (focusAt) {
        $('lg-focus-note').hidden = false;
        const back = sent.get('back') || '';
        if (back.startsWith('/') && !back.startsWith('//')) {
            $('lg-focus-back').href = back;
            $('lg-focus-back').hidden = false;
        }
        day.value = 'all';
        live.setAttribute('aria-pressed', 'false');
        live.textContent = '▶ Live';
        root.classList.add('lg-paused');
    }

    const el = (tag, cls, text) => {
        const n = document.createElement(tag);
        if (cls) n.className = cls;
        if (text !== undefined) n.textContent = text;
        return n;
    };
    const ago = since => {
        const s = Math.max(0, Math.round((Date.now() - new Date(since.replace(' ', 'T'))) / 1000));
        return s < 60 ? `${s}s` : s < 3600 ? `${Math.floor(s / 60)} min` : `${Math.floor(s / 3600)} h ${Math.floor(s % 3600 / 60)} min`;
    };
    const today = () => new Date().toLocaleDateString('en-CA');  // YYYY-MM-DD, as the log writes it
    const t12 = hms => {  // '16:23:44' (as the log writes it) -> '4:23:44 PM'
        const [h, m, sec] = hms.split(':');
        const hour = Number(h);
        return `${hour % 12 || 12}:${m}:${sec} ${hour < 12 ? 'AM' : 'PM'}`;
    };

    function params(extra) {
        const p = new URLSearchParams({ area: area.value, show: show.value, day: day.value, q: search.value.trim() });
        if (focusAt) { p.set('at', focusAt); p.set('find', focusFind); }
        Object.entries(extra || {}).forEach(([k, v]) => p.set(k, v));
        return p;
    }

    function tech(lines) {
        // A traceback, for whoever fixes the app: one click away, so the page stays readable
        const box = el('details', 'lg-tech');
        box.append(el('summary', '', 'Technical detail (the full trace)'), el('pre', 'lg-trace', lines.join('\n')));
        return box;
    }

    function line(x) {
        const row = el('div', `lg-line lv-${x.level}`);
        row.style.paddingLeft = `${Math.min(x.depth, 6) * 14 + 8}px`;
        row.append(el('span', 'lg-line-time', t12(x.time)));
        if (x.level === 'WARNING' || x.level === 'ERROR' || x.level === 'CRITICAL') row.append(el('span', 'lg-tag', x.level_label));
        row.append(el('span', 'lg-line-text', x.text));
        if (x.more && x.more.length) row.append(x.more.some(t => t.startsWith('Traceback')) ? tech(x.more) : el('pre', 'lg-trace', x.more.join('\n')));
        return row;
    }

    function section(title, cls) {
        const s = el('section', `lg-sec${cls ? ' ' + cls : ''}`);
        s.append(el('h4', 'lg-sec-title', title));
        return s;
    }

    const longDay = day => new Date(`${day}T00:00:00`).toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

    function item(a) {
        const li = el('li', `lg-item lv-${a.level}`);
        li.dataset.id = a.id;
        const head = el('button', 'lg-head');
        head.type = 'button';
        head.setAttribute('aria-expanded', open.has(a.id) ? 'true' : 'false');
        const when = el('span', 'lg-time', a.day === today() ? t12(a.clock) : `${a.day.slice(8)}/${a.day.slice(5, 7)} ${t12(a.clock)}`);
        const body = el('span', 'lg-body');
        const top = el('span', 'lg-top');
        top.append(el('span', `lg-area lg-area-${a.area}`, a.area_label));
        if (a.background) top.append(el('span', 'lg-bg', 'in the background'));
        if (a.level !== 'INFO' && a.level !== 'DEBUG') top.append(el('span', 'lg-tag', a.level_label + (a.issues > 1 ? ` (${a.issues})` : '')));
        body.append(top, el('span', 'lg-title', a.title));
        if (a.result) body.append(el('span', 'lg-result', a.result));
        const steps = a.details.filter(x => x.level !== 'DEBUG').length;
        const fine = a.details.length - steps;
        const count = steps || a.details.length;
        head.append(when, body, el('span', 'lg-count', count ? `${count} ${count === 1 ? 'step' : 'steps'} ▾` : 'Details ▾'));
        li.append(head);

        const box = el('div', 'lg-details');
        box.hidden = !open.has(a.id);

        // When, where and how it ran
        const facts = el('dl', 'lg-facts');
        const fact = (k, v) => { facts.append(el('dt', '', k), el('dd', '', v)); };
        fact('When', `${longDay(a.day)}, ${t12(a.clock)}${a.end && a.end !== a.clock ? ` to ${t12(a.end)}` : ''}`);
        fact('Part of the app', a.area_label);
        fact('How it ran', a.background ? 'In the background: the page carried on while it worked' : 'While the page waited for it');
        fact('Outcome', a.failure ? 'It stopped with an error (below)'
            : a.level === 'WARNING' ? `Done, with ${a.issues > 1 ? a.issues + ' things' : 'something'} to look at (marked below)`
            : a.result || 'Done');
        box.append(facts);

        // A failure: the error in its own words, and where in the app it stopped
        if (a.failure) {
            const f = section('Where it stopped', 'lg-sec-fail');
            const err = el('p', 'lg-fail-error');
            err.append(el('strong', '', 'The error: '), document.createTextNode(a.failure.error || 'not written in the log'));
            f.append(err);
            if (a.failure.where) {
                const at = el('p', 'lg-fail-where');
                at.append(el('strong', '', 'In the app: '), document.createTextNode(a.failure.where));
                f.append(at);
                if (a.failure.code) f.append(el('pre', 'lg-fail-code', a.failure.code));
            }
            if (a.failure.inside) {
                const inside = el('p', 'lg-fail-where');
                inside.append(el('strong', '', 'Inside: '), document.createTextNode(a.failure.inside));
                f.append(inside);
            }
            if (a.title_more.length) f.append(tech(a.title_more));
            box.append(f);
        } else if (a.title_more.length) {
            box.append(el('pre', 'lg-trace', a.title_more.join('\n')));
        }

        // Every step, the fine detail too (each label, each address lookup), which can be hidden
        if (a.details.length) {
            const s = section('What it did, step by step');
            if (fine && steps) {
                const label = `Hide the fine detail (${fine} line${fine === 1 ? '' : 's'})`;
                const toggle = el('button', 'ab-btn ab-btn-plain ab-btn-small lg-detail-toggle', label);
                toggle.type = 'button';
                toggle.addEventListener('click', () => {
                    const hidden = s.classList.toggle('lg-hide-detail');
                    toggle.textContent = hidden ? `Show the fine detail (${fine} line${fine === 1 ? '' : 's'})` : label;
                });
                s.append(toggle);
            }
            a.details.forEach(x => s.append(line(x)));
            box.append(s);
        }

        // What the app was doing just before: the lines that led up to this
        if (a.before.length) {
            const s = section('Just before, the app was', 'lg-sec-before');
            a.before.forEach(x => s.append(line(x)));
            box.append(s);
        }

        head.addEventListener('click', () => {
            box.hidden = !box.hidden;
            head.setAttribute('aria-expanded', String(!box.hidden));
            if (box.hidden) open.delete(a.id); else open.add(a.id);
        });
        li.append(box);
        return li;
    }

    function status(d) {
        const text = $('lg-now-text');
        $('lg-now').classList.toggle('lg-running', d.running.length > 0);
        if (d.running.length) {
            const r = d.running[0];
            text.textContent = `Running: ${r.what.trim()} (${r.area_label}${r.background ? ', in the background' : ''}), for ${ago(r.since)}` +
                (d.running.length > 1 ? ` · and ${d.running.length - 1} more` : '');
        } else {
            text.textContent = d.last ? `Nothing running. Last activity at ${t12(d.last.slice(11))} (${ago(d.last)} ago).` : 'Nothing in the log yet.';
        }
        $('lg-count').textContent = d.today.activities;
        $('lg-count-issues').querySelector('strong').textContent = d.today.issues;
        $('lg-count-problems').querySelector('strong').textContent = d.today.problems;
        $('lg-kept').textContent = d.kept_from ? `Kept since ${d.kept_from.slice(0, 10)}` : '';
        const chosen = area.value;
        if (area.options.length - 1 !== d.areas.length) {
            area.replaceChildren(new Option('Every part of the app', ''));
            d.areas.forEach(k => area.append(new Option(d.area_labels[k], k)));
            area.value = d.areas.includes(chosen) ? chosen : '';
        }
    }

    async function load({ older = false, quiet = false } = {}) {
        if (busy) return;
        busy = true;
        try {
            const extra = older ? { before: more.dataset.before } : { limit: Math.max(150, Math.min(shown, 500)) };
            if (quiet && version) extra.since = version;
            const res = await fetch(`${root.dataset.url}?${params(extra)}`);
            if (!res.ok) throw new Error(res.status);
            const d = await res.json();
            if (d.unchanged) { if (lastData) status(lastData); return; }  // still updates "for 12s" / "3 min ago"
            lastData = d;
            version = d.version;
            status(d);
            const focused = focusAt ? d.focus : null;
            if (focused !== null && focused !== undefined) open.add(focused);
            if (!older) list.replaceChildren();
            d.activities.forEach(a => list.append(item(a)));
            if (focusAt) {  // once: the failed run marked and in view; the page is an ordinary log page after that
                const target = focused !== null && focused !== undefined && list.querySelector(`[data-id="${focused}"]`);
                if (target) {
                    target.classList.add('lg-focus');
                    target.scrollIntoView({ block: 'start' });
                } else {
                    $('lg-focus-text').textContent = "What failed isn't in the log any more (it may have been cleared). The newest runs are below.";
                }
                focusAt = null;
            }
            shown = list.children.length;
            more.hidden = !d.more;
            if (d.oldest_shown !== null) more.dataset.before = d.oldest_shown;
            empty.hidden = shown > 0;
            $('lg-summary').textContent = shown ? `${shown} run${shown === 1 ? '' : 's'} shown, newest first` : '';
        } catch (e) {
            $('lg-now-text').textContent = "Couldn't read the log just now. It tries again in a few seconds.";
        } finally {
            busy = false;
        }
    }

    function tick() {
        clearTimeout(timer);
        if (live.getAttribute('aria-pressed') !== 'true') return;
        timer = setTimeout(async () => {
            if (!document.hidden) await load({ quiet: true });
            tick();
        }, 3000);
    }

    const reload = () => { version = ''; focusAt = null; load(); };
    let typing;
    search.addEventListener('input', () => { clearTimeout(typing); typing = setTimeout(reload, 250); });
    [area, show, day].forEach(s => s.addEventListener('change', reload));
    $('lg-count-issues').addEventListener('click', () => { show.value = 'issues'; day.value = 'today'; reload(); });
    $('lg-count-problems').addEventListener('click', () => { show.value = 'problems'; day.value = 'today'; reload(); });
    more.addEventListener('click', () => load({ older: true }));
    live.addEventListener('click', () => {
        const on = live.getAttribute('aria-pressed') !== 'true';
        live.setAttribute('aria-pressed', String(on));
        live.textContent = on ? '⏸ Pause' : '▶ Live';
        root.classList.toggle('lg-paused', !on);
        if (on) load({ quiet: true });
        tick();
    });
    load().then(tick);
});
