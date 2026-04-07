// --- Global State & Configuration ---
let fileHandle = null;
let activeId = null;
let activeType = null;
let expandedMappingId = null;
let expandedMappingType = null;
let filters = { zone: '', state: '', region: '', cluster: '', branch: '' };
let mappingFocus = 'both';
let isEditMode = false;
let history = [];
let isAppInitialized = false; // Guard for double-init
let isolatedColumn = null; // Track which column is currently isolated (full-screen)
let editingRoleEntityId = null; // Track entity open for role editing
let viewingUserMappingId = null; // Track user open in the mapping details drawer

const PREDEFINED_ROLES = {
    zone: ['Zonal Head', 'National Head'],
    state: [],
    region: ['Region Head', 'ROE', 'ROM'],
    cluster: ['Cluster Head'],
    branch: ['Branch Head', 'Assistant Branch Head']
};

// Data is now loaded from data.js


// --- Initialization ---
function init() {
    console.log('MMFSL UI: Initializing Demo Data...');
    renderLists();
}

// --- State Management ---
function saveState() {
    history.push(JSON.parse(JSON.stringify(data)));
    if (history.length > 20) history.shift();
}

// --- DOM Elements ---
let lists, svg, deleteContainer;
function initDOMElements() {
    lists = {
        zone: document.getElementById('zones-list'),
        state: document.getElementById('states-list'),
        region: document.getElementById('regions-list'),
        cluster: document.getElementById('clusters-list'),
        branch: document.getElementById('branches-list')
    };
    svg = document.getElementById('mapping-svg');
    deleteContainer = document.getElementById('delete-btn-container');

    Object.entries(lists).forEach(([k, v]) => { if (!v) console.warn(`Missing list element: ${k}-list`); });
}

// --- Icons ---
const getIconSVG = (type) => {
    const icons = {
        zone: '<path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"></path>',
        state: '<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle>',
        region: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path>',
        cluster: '<rect x="3" y="3" width="7" height="7"></rect><rect x="14" y="3" width="7" height="7"></rect><rect x="14" y="14" width="7" height="7"></rect><rect x="3" y="14" width="7" height="7"></rect>',
        branch: '<path d="M3 21v-4a4 4 0 1 1 4 4H3z"></path><path d="M3 7V3a4 4 0 1 1 4 4H3z"></path><path d="M17 21v-4a4 4 0 1 1 4 4h-4z"></path><path d="M17 7V3a4 4 0 1 1 4 4h-4z"></path>'
    };
    return `<svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="#e41837" stroke-width="2">${icons[type]}</svg>`;
};

const formatCount = (count, singular, plural = singular + 's') => {
    return `${count} ${count === 1 ? singular : plural}`;
};

function getZoneHierarchyCounts(zoneId) {
    const states = data.mappings.zoneToState.filter(m => m.from === zoneId).map(m => m.to);
    const regions = data.mappings.stateToRegion.filter(m => states.includes(m.from)).map(m => m.to);
    const clusters = data.mappings.regionToCluster.filter(m => regions.includes(m.from)).map(m => m.to);
    const branches = data.mappings.clusterToBranch.filter(m => clusters.includes(m.from)).map(m => m.to);
    return { states: states.length, regions: regions.length, clusters: clusters.length, branches: branches.length };
}

function getStateHierarchyCounts(stateId) {
    const regions = data.mappings.stateToRegion.filter(m => m.from === stateId).map(m => m.to);
    const clusters = data.mappings.regionToCluster.filter(m => regions.includes(m.from)).map(m => m.to);
    const branches = data.mappings.clusterToBranch.filter(m => clusters.includes(m.from)).map(m => m.to);
    return { regions: regions.length, clusters: clusters.length, branches: branches.length };
}

function getRegionHierarchyCounts(regionId) {
    const clusters = data.mappings.regionToCluster.filter(m => m.from === regionId).map(m => m.to);
    const branches = data.mappings.clusterToBranch.filter(m => clusters.includes(m.from)).map(m => m.to);
    return { clusters: clusters.length, branches: branches.length };
}

function getClusterHierarchyCounts(clusterId) {
    const branches = data.mappings.clusterToBranch.filter(m => m.from === clusterId).map(m => m.to);
    return { branches: branches.length };
}

function getBranchHierarchy(branchId) {
    const parents = data.mappings.clusterToBranch.filter(m => m.to === branchId).map(m => m.from);
    if (parents.length === 0) return 'Individual Branch';
    const clusterNames = data.clusters.filter(c => parents.includes(c.id)).map(c => c.name);
    const count = clusterNames.length;
    return `Part of: ${clusterNames[0]}${count > 1 ? ` (+${count - 1})` : ''}`;
}

// --- Rendering ---
function renderLists() {
    if (!lists) return;
    // Show/hide Global Unselect button
    const unselectBtn = document.querySelector('.unselect-all-btn');
    if (unselectBtn) unselectBtn.style.display = activeId ? 'flex' : 'none';

    const mappingContainer = document.querySelector('.mapping-container');
    const viewModeToggle = document.querySelector('.view-mode-toggle');
    if (mappingContainer) {
        if (activeId || expandedMappingId) mappingContainer.classList.add('focus-mode');
        else mappingContainer.classList.remove('focus-mode');

        // Handle Isolation UI
        if (isolatedColumn) {
            mappingContainer.classList.add('isolation-active');
            document.querySelectorAll('.mapping-column').forEach(col => {
                col.classList.toggle('isolated', col.dataset.type === isolatedColumn);
            });
        } else {
            mappingContainer.classList.remove('isolation-active');
            document.querySelectorAll('.mapping-column').forEach(col => col.classList.remove('isolated'));
        }
    }

    // Hide View/Edit mode toggle when in user mapping mode or when a cell is expanded
    if (viewModeToggle) {
        const hideToggle = isolatedColumn || expandedMappingId;
        viewModeToggle.style.display = hideToggle ? 'none' : 'flex';
    }

    Object.keys(lists).forEach(type => {
        const listEl = lists[type];
        if (!listEl) return;

        // Conditional Scroll Locking for column with an expanded card
        const isColumnMapping = (expandedMappingId && expandedMappingType === type);
        listEl.classList.toggle('mapping-active', isColumnMapping);

        listEl.innerHTML = '';

        // --- RESTORE COLLAPSED HEADER STATE ---
        const colHeaderTitle = document.querySelector(`.mapping-column[data-type="${type}"] .column-header span`);
        const colHeaderActions = document.querySelector(`.mapping-column[data-type="${type}"] .column-header-actions`);

        if (colHeaderTitle) {
            if (isolatedColumn === type && editingRoleEntityId) {
                const dataKey = type === 'branch' ? 'branches' : type + 's';
                const entity = data[dataKey].find(e => e.id === editingRoleEntityId);
                colHeaderTitle.innerHTML = `Manage User Mapping: ${entity ? entity.name : ''}`;
            } else {
                colHeaderTitle.innerHTML = type.charAt(0).toUpperCase() + type.slice(1) + (type === 'branch' ? 'es' : 's');
            }
        }
        if (colHeaderActions) {
            if (isolatedColumn === type) {
                // When isolated, only show the Back button
                colHeaderActions.innerHTML = `
                    <button class="header-back-btn" title="Back" style="display:flex; align-items:center; gap:4px; font-size:11.5px; font-weight:600; padding:5px 10px; border-radius:6px; background:rgba(255,255,255,0.15); color:white; border:none; cursor:pointer; transition:background 0.2s;">
                        <svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"></polyline></svg> Back
                    </button>
                `;
                colHeaderActions.querySelector('.header-back-btn').onclick = (e) => {
                    e.stopPropagation();
                    if (editingRoleEntityId) {
                        editingRoleEntityId = null;
                        viewingUserMappingId = null;
                        renderLists();
                    } else {
                        toggleIsolation(type);
                    }
                };
            } else {
                colHeaderActions.innerHTML = `
                    <button class="user-mapping-btn" title="User Access Mapping" data-column="${type}">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                            <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path>
                            <circle cx="8.5" cy="7" r="4"></circle>
                            <polyline points="17 11 19 13 23 9"></polyline>
                        </svg>
                    </button>
                `;
                colHeaderActions.querySelector('.user-mapping-btn').onclick = (e) => {
                    e.stopPropagation();
                    toggleIsolation(type);
                };
            }
        }

        // --- DIRECTION TOGGLES VISIBILITY ---
        const dirToggle = document.querySelector(`.mapping-column[data-type="${type}"] .direction-toggle`);
        if (dirToggle) {
            dirToggle.style.display = isolatedColumn ? 'none' : 'flex';
        }

        // --- ROLE EDITOR OVERRIDE ---
        const searchEl = document.querySelector(`.mapping-column[data-type="${type}"] .column-search`);

        if (isolatedColumn === type && editingRoleEntityId) {
            if (searchEl) searchEl.style.display = 'none'; // Hide generic branch search

            renderRoleEditor(listEl, type, editingRoleEntityId);
            return; // Skip normal rendering mapping cards for this column
        } else {
            if (searchEl) searchEl.style.display = 'block'; // Restore search
        }

        const dataKey = type === 'branch' ? 'branches' : type + 's';
        const filteredData = (data[dataKey] || []).filter(item =>
            item.name.toLowerCase().includes(filters[type].toLowerCase())
        );

        const col = document.querySelector(`.mapping-column[data-type="${type}"]`);
        const searchInput = col.querySelector('.column-search-input');
        const hadFocus = (document.activeElement === searchInput);
        const selectionStart = searchInput?.selectionStart;
        const selectionEnd = searchInput?.selectionEnd;

        const { ids } = getHighlightedMappings();
        const mappingContainer = document.querySelector('.mapping-container');
        if (mappingContainer) {
            if (activeId) mappingContainer.classList.add('is-dimmed');
            else mappingContainer.classList.remove('is-dimmed');
        }

        // Update footer count
        const footerCountEl = document.querySelector(`#${type}-footer .footer-count`);
        if (footerCountEl) {
            let visibleCount = 0;
            if (activeId) {
                if (activeType === type) {
                    // Since there is an active item matching this column Type, 
                    // exactly one item is "highlighted" basically.
                    // But maybe we just count how many items in this column are in `ids`.
                    visibleCount = filteredData.filter(item => ids.has(item.id)).length;
                } else {
                    visibleCount = filteredData.filter(item => ids.has(item.id)).length;
                }
            } else {
                visibleCount = filteredData.length;
            }
            if (activeId && activeType === type) {
                visibleCount = 1;
            }
            footerCountEl.textContent = visibleCount;
        }

        // Dynamic Sorting: 
        // 1. Bring expanded mapping card to the very top.
        // 2. Bring active/highlighted (mapped) cards next.
        // 3. Keep rest in original order.
        const sortedData = [...filteredData].sort((a, b) => {
            if (a.id === expandedMappingId) return -1;
            if (b.id === expandedMappingId) return 1;

            const aHighlighted = ids.has(a.id);
            const bHighlighted = ids.has(b.id);
            if (aHighlighted && !bHighlighted) return -1;
            if (!aHighlighted && bHighlighted) return 1;

            return 0;
        });

        // Use DocumentFragment for performance
        const fragment = document.createDocumentFragment();

        sortedData.forEach(item => {
            const isActive = activeId === item.id;
            const isExpanded = expandedMappingId === item.id;

            const card = document.createElement('div');
            card.className = `mapping-card ${isActive ? 'active' : ''} ${ids.has(item.id) && !isActive ? 'highlighted' : ''} ${isExpanded ? 'expanded' : ''}`;
            card.dataset.id = item.id;
            card.dataset.type = type;

            if (isExpanded) {
                // Render Expanded Mapping UI inside the card
                const targetConfigs = {
                    zone: [{ type: 'state', label: 'States (Child)', dir: 'forward' }],
                    state: [{ type: 'zone', label: 'Zones (Parent)', dir: 'backward' }, { type: 'region', label: 'Regions (Child)', dir: 'forward' }],
                    region: [{ type: 'state', label: 'States (Parent)', dir: 'backward' }, { type: 'cluster', label: 'Clusters (Child)', dir: 'forward' }],
                    cluster: [{ type: 'region', label: 'Regions (Parent)', dir: 'backward' }, { type: 'branch', label: 'Branches (Child)', dir: 'forward' }],
                    branch: [{ type: 'cluster', label: 'Clusters (Parent)', dir: 'backward' }]
                };
                const configs = targetConfigs[type];
                let currentConfig = configs[0];
                const getPlural = (t) => {
                    const map = {
                        'zone': 'Zones',
                        'state': 'States',
                        'region': 'Regions',
                        'cluster': 'Clusters',
                        'branch': 'Branches'
                    };
                    return map[t] || (t.charAt(0).toUpperCase() + t.slice(1) + 's');
                };

                card.innerHTML = `
                    <div class="inline-mapping-header">
                        <div class="selector-title">Map <strong>${item.name}</strong> to:</div>
                        <button class="close-selector-btn" title="Close Mapping">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                                <line x1="18" y1="6" x2="6" y2="18"></line>
                                <line x1="6" y1="6" x2="18" y2="18"></line>
                            </svg>
                        </button>
                    </div>
                    <div class="selector-header">
                        ${configs.length > 1 ? `<div class="selector-tabs">${configs.map((c, i) => `<div class="selector-tab ${i === 0 ? 'active' : ''}" data-index="${i}">${getPlural(c.type)}</div>`).join('')}</div>` : ''}
                        <input type="text" class="selector-search" placeholder="Search ${getPlural(currentConfig.type)}..." autofocus>
                    </div>
                    <div class="selector-list"></div>
                `;

                const searchInput = card.querySelector('.selector-search');
                const listContainer = card.querySelector('.selector-list');
                const tabs = card.querySelectorAll('.selector-tab');
                const closeBtn = card.querySelector('.close-selector-btn');

                const renderInlineItems = (filter = '') => {
                    listContainer.innerHTML = '';
                    const dataKey = currentConfig.type === 'branch' ? 'branches' : currentConfig.type + 's';
                    const targets = data[dataKey] || [];

                    const checkMapping = (targetId) => {
                        let mappingKey = '';
                        let fromId = item.id, toId = targetId, fType = type, tType = currentConfig.type;
                    if (fType === 'zone' && tType === 'state') mappingKey = 'zoneToState';
                    else if (fType === 'state' && tType === 'zone') { mappingKey = 'zoneToState';[fromId, toId] = [toId, fromId]; }
                    else if (fType === 'state' && tType === 'region') mappingKey = 'stateToRegion';
                    else if (fType === 'region' && tType === 'state') { mappingKey = 'stateToRegion';[fromId, toId] = [toId, fromId]; }
                    else if (fType === 'region' && tType === 'cluster') mappingKey = 'regionToCluster';
                    else if (fType === 'cluster' && tType === 'region') { mappingKey = 'regionToCluster';[fromId, toId] = [toId, fromId]; }
                    else if (fType === 'cluster' && tType === 'branch') mappingKey = 'clusterToBranch';
                    else if (fType === 'branch' && tType === 'cluster') { mappingKey = 'clusterToBranch';[fromId, toId] = [toId, fromId]; }
                    return mappingKey && data.mappings[mappingKey].some(m => m.from === fromId && m.to === toId);
                    };

                    targets.filter(t => t.name.toLowerCase().includes(filter.toLowerCase())).forEach(target => {
                        const isMapped = checkMapping(target.id);
                        const addIcon = `<svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>`;
                        const checkIcon = `<svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>`;
                        const row = document.createElement('div');
                        row.className = 'selector-item';
                        row.innerHTML = `<span>${target.name}</span><button class="selector-toggle-btn ${isMapped ? 'toggle-remove' : 'toggle-add'}" title="${isMapped ? 'Remove' : 'Add'}">${isMapped ? checkIcon : addIcon}</button>`;

                        const btn = row.querySelector('.selector-toggle-btn');
                        btn.onclick = (e) => {
                            e.stopPropagation();
                            if (isMapped) {
                                let fId = item.id, tId = target.id, fType = type, tType = currentConfig.type;
                                if (fType === 'state' && tType === 'zone') { [fId, tId] = [tId, fId];[fType, tType] = [tType, fType]; }
                                if (fType === 'region' && tType === 'state') { [fId, tId] = [tId, fId];[fType, tType] = [tType, fType]; }
                                if (fType === 'cluster' && tType === 'region') { [fId, tId] = [tId, fId];[fType, tType] = [tType, fType]; }
                                if (fType === 'branch' && tType === 'cluster') { [fId, tId] = [tId, fId];[fType, tType] = [tType, fType]; }
                                deleteMapping(fType, tType, fId, tId);
                            } else {
                                createMapping(item.id, type, target.id, currentConfig.type);
                            }
                            renderInlineItems(searchInput.value);
                        };
                        listContainer.appendChild(row);
                    });
                };

                const titleEl = card.querySelector('.selector-title');
                const updateTitle = () => {
                    const plural = getPlural(currentConfig.type);
                    const itemType = type.charAt(0).toUpperCase() + type.slice(1);
                    if (currentConfig.dir === 'forward') {
                        titleEl.innerHTML = `Map <strong>${plural}</strong> to <strong>${item.name}</strong> ${itemType}`;
                    } else {
                        titleEl.innerHTML = `Map <strong>${item.name}</strong> ${itemType} to <strong>${plural}</strong>`;
                    }
                };

                tabs.forEach(tab => {
                    tab.onclick = (e) => {
                        e.stopPropagation();
                        tabs.forEach(t => t.classList.remove('active'));
                        tab.classList.add('active');
                        currentConfig = configs[parseInt(tab.dataset.index)];
                        searchInput.placeholder = `Search ${getPlural(currentConfig.type)}...`;
                        updateTitle();
                        renderInlineItems(searchInput.value);
                    };
                });

                updateTitle();

                searchInput.onclick = (e) => e.stopPropagation();
                searchInput.oninput = (e) => renderInlineItems(e.target.value);
                closeBtn.onclick = (e) => { e.stopPropagation(); closeMappingSelector(); };

                renderInlineItems();
            } else {
                // Render Normal Card Content
                let subtitle = '';
                if (type === 'zone') {
                    const counts = getZoneHierarchyCounts(item.id);
                    subtitle = `${formatCount(counts.states, 'State')}, ${formatCount(counts.regions, 'Region')}, ${formatCount(counts.clusters, 'Cluster')}, ${formatCount(counts.branches, 'Branch', 'Branches')}`;
                } else if (type === 'state') {
                    const counts = getStateHierarchyCounts(item.id);
                    subtitle = `${formatCount(counts.regions, 'Region')}, ${formatCount(counts.clusters, 'Cluster')}, ${formatCount(counts.branches, 'Branch', 'Branches')}`;
                } else if (type === 'region') {
                    const counts = getRegionHierarchyCounts(item.id);
                    subtitle = `${formatCount(counts.clusters, 'Cluster')}, ${formatCount(counts.branches, 'Branch', 'Branches')}`;
                } else if (type === 'cluster') {
                    const counts = getClusterHierarchyCounts(item.id);
                    subtitle = formatCount(counts.branches, 'Branch', 'Branches');
                } else {
                    subtitle = getBranchHierarchy(item.id);
                }

                const userCount = data.mappings.userRoles.filter(m => m.entityId === item.id).length;
                const userBadgeHTML = userCount > 0 ? `<div class="card-user-badge" title="${userCount} Users Assigned"><svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="8.5" cy="7" r="4"></circle></svg> ${userCount}</div>` : '';

                card.innerHTML = `
                    ${isActive ? `<button class="unselect-card-btn" title="Unselect">&times;</button>` : ''}
                    <div class="card-icon">${getIconSVG(type)}</div>
                    <div class="card-content">
                        <div class="card-title">${item.name}</div>
                        <div class="card-subtitle-row">
                            <span class="card-subtitle">${subtitle}</span>
                            ${userBadgeHTML}
                        </div>
                    </div>
                    <div class="card-actions">
                        ${(isEditMode && type !== 'branch') ? `
                        <button class="add-mapping-btn" title="Add Mapping">
                            <svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <line x1="12" y1="5" x2="12" y2="19"></line>
                                <line x1="5" y1="12" x2="19" y2="12"></line>
                            </svg>
                        </button>` : ''}
                    </div>
                `;

                if (isEditMode) {
                    const btn = card.querySelector('.add-mapping-btn');
                    if (btn) {
                        btn.addEventListener('click', (e) => { e.stopPropagation(); showMappingSelector(item.id, type, e); });
                    }
                }
            }

            card.addEventListener('click', (e) => {
                if (e.target.classList.contains('unselect-card-btn')) {
                    handleCardClick(item.id, type);
                    return;
                }
                handleCardClick(item.id, type);
            });
            card.draggable = isEditMode;
            card.addEventListener('dragstart', (e) => {
                if (!isEditMode) { e.preventDefault(); return; }
                e.dataTransfer.setData('text/plain', JSON.stringify({ id: item.id, type: type }));
                card.classList.add('dragging');
                activeId = item.id;
                activeType = type;
                renderLists();
            });
            card.addEventListener('dragend', () => card.classList.remove('dragging'));
            card.addEventListener('dragover', (e) => { if (isEditMode) e.preventDefault(); });
            card.addEventListener('drop', (e) => {
                if (!isEditMode) return;
                e.preventDefault(); e.stopPropagation();
                const source = JSON.parse(e.dataTransfer.getData('text/plain'));
                createMapping(source.id, source.type, item.id, type);
            });
            fragment.appendChild(card);
        });

        listEl.innerHTML = '';
        listEl.appendChild(fragment);

        // Restore focus if needed
        if (hadFocus && searchInput) {
            searchInput.focus();
            if (selectionStart !== null) {
                searchInput.setSelectionRange(selectionStart, selectionEnd);
            }
        }
    });
        // requestAnimationFrame(drawConnections);
}

function showMappingSelector(id, type, event) {
    if (expandedMappingId === id) {
        expandedMappingId = null;
        expandedMappingType = null;
    } else {
        activeId = id;
        activeType = type;
        expandedMappingId = id;
        expandedMappingType = type;
    }
    renderLists();
}

function closeMappingSelector() {
    const card = document.querySelector('.mapping-card.expanded');
    if (card) {
        card.classList.add('closing');
        // Wait for the exit animation (300ms) before re-rendering the list
        setTimeout(() => {
            expandedMappingId = null;
            expandedMappingType = null;
            renderLists();
        }, 300);
    } else {
        expandedMappingId = null;
        expandedMappingType = null;
        renderLists();
    }
}

function handleCardClick(id, type) {
    // Intercept clicks in Isolation Mode to open the Role Editor
    if (isolatedColumn === type) {
        if (editingRoleEntityId === id) {
            editingRoleEntityId = null;
        } else {
            editingRoleEntityId = id;
        }
        renderLists();
        return;
    }

    const prevActiveId = activeId;
    if (activeId === id) {
        activeId = null;
        activeType = null;
    } else {
        activeId = id;
        activeType = type;
    }

    // Scroll locking removed to ensure full navigability during mapping

    renderLists();

    // Smooth scroll into view if card is not fully visible after re-render
    if (activeId && activeId !== prevActiveId) {
        setTimeout(() => {
            const activeCard = document.querySelector(`.mapping-card.active[data-id="${activeId}"]`);
            if (activeCard) {
                const list = activeCard.closest('.column-list');
                const isVisible = (
                    activeCard.offsetTop >= list.scrollTop &&
                    (activeCard.offsetTop + activeCard.offsetHeight) <= (list.scrollTop + list.offsetHeight)
                );

                if (!isVisible) {
                    activeCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }

            // Also scroll highlighted connections into view smoothly
            const { ids } = getHighlightedMappings();
            ids.forEach(highlightId => {
                const el = document.querySelector(`.mapping-card.highlighted[data-id="${highlightId}"]`);
                if (el) {
                    const list = el.closest('.column-list');
                    const isVisible = (
                        el.offsetTop >= list.scrollTop &&
                        (el.offsetTop + el.offsetHeight) <= (list.scrollTop + list.offsetHeight)
                    );
                    if (!isVisible) {
                        el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                    }
                }
            });
        }, 100);
    }
}

function createMapping(fromId, fromType, toId, toType) {
    saveState();
    let mappingKey = '';
    if (fromType === 'zone' && toType === 'state') mappingKey = 'zoneToState';
    if (fromType === 'state' && toType === 'zone') { mappingKey = 'zoneToState';[fromId, toId] = [toId, fromId]; }
    if (fromType === 'state' && toType === 'region') mappingKey = 'stateToRegion';
    if (fromType === 'region' && toType === 'state') { mappingKey = 'stateToRegion';[fromId, toId] = [toId, fromId]; }
    if (fromType === 'region' && toType === 'cluster') mappingKey = 'regionToCluster';
    if (fromType === 'cluster' && toType === 'region') { mappingKey = 'regionToCluster';[fromId, toId] = [toId, fromId]; }
    if (fromType === 'cluster' && toType === 'branch') mappingKey = 'clusterToBranch';
    if (fromType === 'branch' && toType === 'cluster') { mappingKey = 'clusterToBranch';[fromId, toId] = [toId, fromId]; }
    if (mappingKey) {
        const exists = data.mappings[mappingKey].some(m => m.from === fromId && m.to === toId);
        if (!exists) { data.mappings[mappingKey].push({ from: fromId, to: toId }); renderLists(); }
    }
}

function deleteMapping(fromType, toType, fromId, toId) {
    saveState();
    let mappingKey = '';
    if (fromType === 'zone' && toType === 'state') mappingKey = 'zoneToState';
    if (fromType === 'state' && toType === 'region') mappingKey = 'stateToRegion';
    if (fromType === 'region' && toType === 'cluster') mappingKey = 'regionToCluster';
    if (fromType === 'cluster' && toType === 'branch') mappingKey = 'clusterToBranch';
    if (mappingKey) {
        data.mappings[mappingKey] = data.mappings[mappingKey].filter(m => !(m.from === fromId && m.to === toId));
        renderLists();
    }
}

function getHighlightedMappings() {
    if (!activeId) return { ids: new Set(), paths: new Set() };
    const ids = new Set([activeId]);
    const paths = new Set();
    const showForward = (mappingFocus === 'both' || mappingFocus === 'forward');
    const showBackward = (mappingFocus === 'both' || mappingFocus === 'backward');

    // Separate backward and forward passes to prevent sibling highlighting
    if (showBackward) {
        let currentIds = new Set([activeId]);
        let changed = true;
        while (changed) {
            let startSize = currentIds.size;
            // Upward: Branch -> Cluster -> Region -> State -> Zone
            data.mappings.clusterToBranch.forEach(m => { if (currentIds.has(m.to)) { currentIds.add(m.from); paths.add(`${m.from}-${m.to}`); ids.add(m.from); } });
            data.mappings.regionToCluster.forEach(m => { if (currentIds.has(m.to)) { currentIds.add(m.from); paths.add(`${m.from}-${m.to}`); ids.add(m.from); } });
            data.mappings.stateToRegion.forEach(m => { if (currentIds.has(m.to)) { currentIds.add(m.from); paths.add(`${m.from}-${m.to}`); ids.add(m.from); } });
            data.mappings.zoneToState.forEach(m => { if (currentIds.has(m.to)) { currentIds.add(m.from); paths.add(`${m.from}-${m.to}`); ids.add(m.from); } });
            if (currentIds.size === startSize) changed = false;
        }
    }

    if (showForward) {
        let currentIds = new Set([activeId]);
        let changed = true;
        while (changed) {
            let startSize = currentIds.size;
            // Downward: Zone -> State -> Region -> Cluster -> Branch
            data.mappings.zoneToState.forEach(m => { if (currentIds.has(m.from)) { currentIds.add(m.to); paths.add(`${m.from}-${m.to}`); ids.add(m.to); } });
            data.mappings.stateToRegion.forEach(m => { if (currentIds.has(m.from)) { currentIds.add(m.to); paths.add(`${m.from}-${m.to}`); ids.add(m.to); } });
            data.mappings.regionToCluster.forEach(m => { if (currentIds.has(m.from)) { currentIds.add(m.to); paths.add(`${m.from}-${m.to}`); ids.add(m.to); } });
            data.mappings.clusterToBranch.forEach(m => { if (currentIds.has(m.from)) { currentIds.add(m.to); paths.add(`${m.from}-${m.to}`); ids.add(m.to); } });
            if (currentIds.size === startSize) changed = false;
        }
    }

    return { ids, paths };
}

function drawConnections() {
    return; // Lines removed as per user request
    if (!svg || !deleteContainer) return;
    svg.innerHTML = '';
    deleteContainer.innerHTML = '';

    const container = document.querySelector('.mapping-container');
    if (!container) return;

    // Reset canvas to container size
    svg.setAttribute('height', container.offsetHeight);
    svg.setAttribute('width', container.offsetWidth);

    // Hide all mappings unless a card is active/selected OR if in isolation mode
    if (!activeId || isolatedColumn) return;

    const highlighted = getHighlightedMappings();

    // Performance optimization: Instead of iterating thousands of mappings,
    // only iterate over the highlighted (active) paths.
    highlighted.paths.forEach(pathKey => {
        const [fromId, toId] = pathKey.split('-');

        // Find cards for this path across all potential type pairs
        const typePairs = [
            ['zone', 'state'],
            ['state', 'region'],
            ['region', 'cluster'],
            ['cluster', 'branch']
        ];

        for (const [fType, tType] of typePairs) {
            const fromCard = document.querySelector(`[data-id="${fromId}"][data-type="${fType}"]`);
            const toCard = document.querySelector(`[data-id="${toId}"][data-type="${tType}"]`);

            if (fromCard && toCard) {
                drawCurve(fromCard, toCard, true, fType, tType, fromId, toId);
                break; // Found the connection cards, move to next path
            }
        }
    });
}

function drawCurve(fromEl, toEl, isActive, fromType, toType, fromId, toId) {
    const fromList = fromEl.closest('.column-list');
    const toList = toEl.closest('.column-list');
    if (!fromList || !toList) return;

    const fromListRect = fromList.getBoundingClientRect();
    const toListRect = toList.getBoundingClientRect();
    const fromRect = fromEl.getBoundingClientRect();
    const toRect = toEl.getBoundingClientRect();

    const getAnchorY = (el) => {
        const rect = el.getBoundingClientRect();
        const header = el.querySelector('.inline-mapping-header');
        if (header) {
            const hRect = header.getBoundingClientRect();
            return hRect.top + hRect.height / 2;
        }
        const icon = el.querySelector('.card-icon');
        if (icon) {
            const iRect = icon.getBoundingClientRect();
            return iRect.top + iRect.height / 2;
        }
        return rect.top + 28;
    };

    const y1_raw = getAnchorY(fromEl);
    const y2_raw = getAnchorY(toEl);

    // Check visibility in vertical columns
    const isFromVisible = (y1_raw >= fromListRect.top - 5 && y1_raw <= fromListRect.bottom + 5);
    const isToVisible = (y2_raw >= toListRect.top - 5 && y2_raw <= toListRect.bottom + 5);

    if (!isFromVisible || !isToVisible) return;

    const svgRect = svg.getBoundingClientRect();
    const x1 = fromRect.right - svgRect.left;
    const y1 = y1_raw - svgRect.top;
    const x2 = toRect.left - svgRect.left;
    const y2 = y2_raw - svgRect.top;

    const midX = x1 + (x2 - x1) / 2;
    const d = `M ${x1} ${y1} C ${midX} ${y1} ${midX} ${y2} ${x2} ${y2}`;

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    path.setAttribute('class', `connection-path ${isActive ? 'active' : ''}`);
    svg.appendChild(path);

    if (isEditMode && isActive && deleteContainer) {
        const midY = (y1 + y2) / 2;
        const btn = document.createElement('button');
        btn.className = 'delete-mapping-btn';
        btn.innerHTML = '&times;';
        btn.style.left = `${midX}px`;
        btn.style.top = `${midY}px`;
        btn.onclick = (e) => { e.stopPropagation(); deleteMapping(fromType, toType, fromId, toId); };
        deleteContainer.appendChild(btn);
    }
}

// --- Toggles & Search ---
function bindEvents() {
    document.querySelectorAll('.dir-btn').forEach(btn => {
        btn.onclick = (e) => {
            e.stopPropagation();
            mappingFocus = btn.dataset.dir;
            document.querySelectorAll('.dir-btn').forEach(b => b.classList.toggle('active', b === btn));
            renderLists();
        };
    });

    document.querySelectorAll('.toggle-btn').forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            isEditMode = (btn.dataset.mode === 'edit');
            document.querySelectorAll('.mapping-column').forEach(col => col.classList.toggle('edit-active', isEditMode));
            renderLists();
        };
    });

    document.querySelectorAll('.column-search-input').forEach(input => {
        const type = input.closest('.mapping-column').dataset.type;
        input.oninput = (e) => { filters[type] = e.target.value; renderLists(); };
    });

    // --- Sync Mappings on Row Scroll & Global Vertical Scroll ---
    let rafId;
    document.querySelectorAll('.column-list').forEach(list => {
        list.addEventListener('scroll', () => {
            if (rafId) cancelAnimationFrame(rafId);
            rafId = requestAnimationFrame(drawConnections);
        });
    });

    const mappingContainer = document.querySelector('.mapping-container');
    if (mappingContainer) {
        mappingContainer.addEventListener('scroll', () => {
            if (rafId) cancelAnimationFrame(rafId);
            rafId = requestAnimationFrame(drawConnections);
        });
    }

    const unselectBtn = document.querySelector('.unselect-all-btn');
    if (unselectBtn) {
        unselectBtn.onclick = () => {
            activeId = null;
            activeType = null;
            document.querySelectorAll('.column-list').forEach(list => list.classList.remove('scroll-locked'));
            renderLists();
        };
    }

    document.querySelector('.export-btn').onclick = () => { alert('Exporting dashboard data...'); };
    document.querySelector('.save-btn').onclick = () => { alert('Changes saved to local session!'); };

    // --- Column Isolation ---
    document.querySelectorAll('.user-mapping-btn').forEach(btn => {
        btn.onclick = (e) => {
            e.stopPropagation();
            const type = btn.dataset.column;
            toggleIsolation(type);
        };
    });

    window.addEventListener('resize', () => {
        if (rafId) cancelAnimationFrame(rafId);
        rafId = requestAnimationFrame(drawConnections);
    });
}

function toggleIsolation(type) {
    if (isolatedColumn === type) {
        isolatedColumn = null; // Exit isolation
    } else {
        isolatedColumn = type; // Enter isolation for specific column
        // When isolating, we should clear active mappings for visual clarity
        activeId = null;
        activeType = null;
        expandedMappingId = null;
        expandedMappingType = null;
        editingRoleEntityId = null; // Clear role editor state on toggle
    }
    renderLists();
}

// --- User Role Editor ---
function renderRoleEditor(container, type, entityId) {
    if (type === 'state') {
        container.innerHTML = `
            <div style="display:flex; flex-direction:column; height:100%; min-height:400px; align-items:center; justify-content:center; background: white; color: #374151; text-align:center; border-radius: 8px; margin: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <svg viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2" style="width:48px; height:48px; margin-bottom:12px;">
                    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
                    <line x1="12" y1="9" x2="12" y2="13"></line>
                    <line x1="12" y1="17" x2="12.01" y2="17"></line>
                </svg>
                <div style="font-weight:600; font-size:16px; color: #111;">No Org structure available for the State</div>
                <div style="font-size:13px; margin-top:6px; color: #6b7280; max-width: 200px;"></div>
            </div>
        `;
        return;
    }

    const dataKey = type === 'branch' ? 'branches' : type + 's';
    const entity = data[dataKey].find(e => e.id === entityId);
    if (!entity) return;

    const roles = PREDEFINED_ROLES[type] || [];
    let activeRole = roles[0];
    let searchQuery = '';

    const renderEditorContent = () => {
        container.innerHTML = `
            <div class="role-editor-wrapper">
                <div class="role-tabs">
                    ${roles.map(role => {
            const count = data.mappings.userRoles.filter(m => m.role === role && m.entityId === entityId).length;
            return `<button class="role-tab ${role === activeRole ? 'active' : ''}" data-role="${role}">
                            ${role} ${count > 0 ? `<span class="role-count-badge">${count}</span>` : ''}
                        </button>`;
        }).join('')}
                </div>

                <div class="role-editor-body">
                    <!-- Left: User Selection -->
                    <div class="role-users-section">
                        <div class="search-wrapper">
                            <input type="text" placeholder="Search System Users..." class="user-search-input" value="${searchQuery}">
                        </div>
                        <div class="user-list"></div>
                    </div>

                    <!-- Right: Context Panel (Overwritable profile vs nearby list) -->
                    <div class="role-context-section" style="flex: 1.2; min-width: 320px; overflow-y: auto; background: #f9fafb; border-left: 1px solid #e5e7eb;"></div>
                </div>
            </div>
        `;

        // Bind Tab Events
        container.querySelectorAll('.role-tab').forEach(tab => {
            tab.onclick = () => {
                activeRole = tab.dataset.role;
                renderEditorContent();
            };
        });

        // Bind Search
        const searchInput = container.querySelector('.user-search-input');
        searchInput.oninput = (e) => {
            searchQuery = e.target.value.toLowerCase();
            renderUsers(); // only trigger user re-render so we don't lose focus
        };
        // Ensure focus returns to end of input if re-rendered while typing
        if (searchQuery) {
            searchInput.focus();
            searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
        }

        const renderUsers = () => {
            const userListEl = container.querySelector('.user-list');
            userListEl.innerHTML = '';

            const roleDeptMap = {
                'Zonal Head': 'Zonal Head',
                'National Head': 'National Head',
                'Region Head': 'Region Head',
                'ROE': 'ROE',
                'ROM': 'ROM',
                'Cluster Head': 'Cluster Head',
                'Branch Head': 'Branch Head',
                'Assistant Branch Head': 'Assistant Branch Head'
            };
            const targetDept = roleDeptMap[activeRole];

            let filteredUsers = systemUsers.filter(u => {
                // Must match the role's department if mapping exists
                if (targetDept && u.department !== targetDept) return false;

                // Search filter
                const q = searchQuery.toLowerCase();
                return u.name.toLowerCase().includes(q) || u.empId.toLowerCase().includes(q);
            });

            // Sort to bring assigned users to the very top
            filteredUsers.sort((a, b) => {
                const aAssigned = data.mappings.userRoles.some(m => m.userId === a.id && m.role === activeRole && m.entityId === entityId);
                const bAssigned = data.mappings.userRoles.some(m => m.userId === b.id && m.role === activeRole && m.entityId === entityId);
                if (aAssigned && !bAssigned) return -1;
                if (!aAssigned && bAssigned) return 1;
                return 0;
            });

            filteredUsers.forEach(user => {
                const isAssigned = data.mappings.userRoles.some(m => m.userId === user.id && m.role === activeRole && m.entityId === entityId);
                const allUserAssignments = data.mappings.userRoles.filter(m => m.userId === user.id);

                const viewDetailsBtnHTML = `
                    <button class="view-mapping-btn">
                        View Mapping <span class="view-mapping-count">${allUserAssignments.length}</span>
                    </button>
                `;



                const userCard = document.createElement('div');
                userCard.className = `user-card ${isAssigned ? 'assigned' : ''} ${viewingUserMappingId === user.id ? 'viewing-active' : ''}`;
                userCard.innerHTML = `
                    <div class="user-info">
                        <div class="user-avatar">${user.name.charAt(0)}</div>
                        <div class="user-details">
                            <div class="user-name-row">
                                <div class="user-name">${user.name}</div>
                            </div>
                            <div class="user-emp-id">${user.empId} ${isAssigned ? `<span class="active-role-tag">• ${activeRole}</span>` : `• ${user.department}`}</div>
                            <div class="user-card-actions" style="margin-top: 6px;">
                                ${viewDetailsBtnHTML}
                            </div>
                        </div>
                    </div>
                    <button class="assign-user-btn ${isAssigned ? 'revoke' : 'assign'}">
                        ${isAssigned ? 'Revoke' : 'Assign'}
                    </button>
                `;

                // Bind View Mapping Button
                const viewBtn = userCard.querySelector('.view-mapping-btn');
                if (viewBtn) {
                    viewBtn.onclick = () => {
                        viewingUserMappingId = user.id;
                        renderEditorContent();
                    };
                }

                // Bind Assign/Revoke
                userCard.querySelector('.assign-user-btn').onclick = () => {
                    if (isAssigned) {
                        data.mappings.userRoles = data.mappings.userRoles.filter(m => !(m.userId === user.id && m.role === activeRole && m.entityId === entityId));
                    } else {
                        data.mappings.userRoles.push({ userId: user.id, role: activeRole, entityId: entityId });
                    }
                    renderEditorContent();
                };
                userListEl.appendChild(userCard);
            });
        };

        const renderUserDetails = () => {
            const contextSection = container.querySelector('.role-context-section');
            if (!viewingUserMappingId) return;

            const user = systemUsers.find(u => u.id === viewingUserMappingId);
            if (!user) return;

            const allAssignments = data.mappings.userRoles.filter(m => m.userId === user.id);

            let HTML = `
                <div style="padding: 24px;">
                    <div class="ticket-wrapper" style="background: white; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.06); position: relative; overflow: hidden; border: 1px solid #f0f0f0;">
                        <!-- Elevated Floating Close Button over Cover -->
                        <button class="ticket-close-btn close-details-btn" title="Close Profile" style="position: absolute; top: 12px; right: 12px; width: 32px; height: 32px; padding: 0; background: rgba(0,0,0,0.25); backdrop-filter: blur(4px); border: none; border-radius: 50%; color: white; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: background 0.2s, transform 0.2s; z-index: 10;" onmouseover="this.style.background='rgba(0,0,0,0.5)'; this.style.transform='scale(1.05)';" onmouseout="this.style.background='rgba(0,0,0,0.25)'; this.style.transform='scale(1)';">
                            <svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="margin: 0; display: block;"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                        </button>
                        
                        <!-- Top Half of Ticket -->
                        <div class="user-profile-cover" style="height: 100px; background: linear-gradient(135deg, var(--primary-theme), #b91c1c); width: 100%;"></div>
                        <div class="details-profile" style="padding: 0 20px 24px; display: flex; flex-direction: column; align-items: center; text-align: center; margin-top: -46px;">
                            <div class="profile-avatar-wrapper" style="background: white; padding: 6px; border-radius: 50%; margin-bottom: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.06);">
                                <div class="user-avatar xl" style="width: 76px; height: 76px; font-size: 32px; background: #111827; border: 1px solid #f3f4f6;">${user.name.charAt(0)}</div>
                            </div>
                            <div class="details-name" style="font-size: 20px; font-weight: 700; color: #111; margin: 0 0 12px 0; letter-spacing: -0.01em;">${user.name}</div>
                            <div class="details-badges" style="display: flex; gap: 8px; align-items: center; justify-content: center;">
                                <span class="user-meta-badge"><svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg> ${user.empId}</span>
                                <span class="user-meta-badge"><svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path></svg> ${user.department}</span>
                            </div>
                        </div>

                        <!-- Ticket Cutout Separator -->
                        <div class="ticket-separator" style="position: relative; border-top: 2px dashed #e5e7eb; margin: 0;">
                            <div style="position: absolute; top: -14px; left: -14px; width: 28px; height: 28px; background: #f9fafb; border-radius: 50%; box-shadow: inset -2px 0px 4px rgba(0,0,0,0.03); z-index: 2; border-right: 1px solid #f0f0f0;"></div>
                            <div style="position: absolute; top: -14px; right: -14px; width: 28px; height: 28px; background: #f9fafb; border-radius: 50%; box-shadow: inset 2px 0px 4px rgba(0,0,0,0.03); z-index: 2; border-left: 1px solid #f0f0f0;"></div>
                        </div>

                        <!-- Bottom Half of Ticket -->
                        <div class="details-mapping-list" style="padding: 24px 20px; background: #fffcfc;">
                            <h4 class="mapping-section-title" style="margin: 0 0 16px 0; font-size: 13px; font-weight: 700; color: #111; text-transform: uppercase; letter-spacing: 0.05em; display: flex; align-items: center; justify-content: space-between;">
                                Active Mapping 
                                <span style="background: var(--primary-theme); color: white; padding: 2px 8px; border-radius: 12px; font-size: 11px;">${allAssignments.length}</span>
                            </h4>
            `;

            if (allAssignments.length === 0) {
                HTML += `<div class="no-mappings">No active mapping found.</div>`;
            } else {
                allAssignments.forEach(m => {
                    let entityName = 'Unknown Entity';
                    let entityGroupName = 'Geography';
                    for (let typeKey of ['zones', 'states', 'regions', 'clusters', 'branches']) {
                        const e = data[typeKey].find(item => item.id === m.entityId);
                        if (e) {
                            entityName = e.name;
                            entityGroupName = typeKey.charAt(0).toUpperCase() + typeKey.slice(1, -1);
                            break;
                        }
                    }
                    HTML += `
                        <div class="mapping-record" style="background: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 14px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); transition: transform 0.2s, box-shadow 0.2s;">
                            <div class="mapping-role"><span class="badge role">${m.role}</span></div>
                            <div class="mapping-entity">
                                <span class="badge type">${entityGroupName}</span> ${entityName}
                            </div>
                        </div>
                    `;
                });
            }

            HTML += `
                        </div>
                    </div>
                </div>
            `;
            contextSection.innerHTML = HTML;

            // Bind Close Event
            contextSection.querySelector('.close-details-btn').onclick = () => {
                viewingUserMappingId = null;
                renderEditorContent();
            };
        };

        const renderNearby = () => {
            const contextSection = container.querySelector('.role-context-section');

            // Logic to find nearby siblings and parent name
            let siblings = [];
            let parentName = '';
            let parentType = '';
            if (type === 'branch') {
                const parentClusterId = data.mappings.clusterToBranch.find(m => m.to === entityId)?.from;
                if (parentClusterId) {
                    const parentEntity = data.clusters.find(c => c.id === parentClusterId);
                    parentName = parentEntity ? parentEntity.name : '';
                    parentType = 'Cluster';
                    const siblingBranchIds = data.mappings.clusterToBranch.filter(m => m.from === parentClusterId && m.to !== entityId).map(m => m.to);
                    siblings = data.branches.filter(b => siblingBranchIds.includes(b.id));
                }
            } else if (type === 'cluster') {
                const parentId = data.mappings.regionToCluster.find(m => m.to === entityId)?.from;
                if (parentId) {
                    const parentEntity = data.regions.find(c => c.id === parentId);
                    parentName = parentEntity ? parentEntity.name : '';
                    parentType = 'Region';
                    const siblingIds = data.mappings.regionToCluster.filter(m => m.from === parentId && m.to !== entityId).map(m => m.to);
                    siblings = data.clusters.filter(b => siblingIds.includes(b.id));
                }
            } else if (type === 'region') {
                const parentId = data.mappings.stateToRegion.find(m => m.to === entityId)?.from;
                if (parentId) {
                    const parentEntity = data.states.find(r => r.id === parentId);
                    parentName = parentEntity ? parentEntity.name : '';
                    parentType = 'State';
                    const siblingIds = data.mappings.stateToRegion.filter(m => m.from === parentId && m.to !== entityId).map(m => m.to);
                    siblings = data.regions.filter(b => siblingIds.includes(b.id));
                }
            } else if (type === 'state') {
                const parentId = data.mappings.zoneToState.find(m => m.to === entityId)?.from;
                if (parentId) {
                    const parentEntity = data.zones.find(z => z.id === parentId);
                    parentName = parentEntity ? parentEntity.name : '';
                    parentType = 'Zone';
                    const siblingIds = data.mappings.zoneToState.filter(m => m.from === parentId && m.to !== entityId).map(m => m.to);
                    siblings = data.states.filter(b => siblingIds.includes(b.id));
                }
            }

            const pluralLabel = type === 'branch' ? 'Branches' : (type.charAt(0).toUpperCase() + type.slice(1) + 's');
            const singularLabel = type.charAt(0).toUpperCase() + type.slice(1);
            const typeLabel = siblings.length === 1 ? singularLabel : pluralLabel;
            const subtitle = parentName
                ? `${typeLabel} in ${parentType}: <strong>${parentName}</strong>`
                : `Nearby ${typeLabel}`;

            contextSection.innerHTML = `
                <div style="padding: 20px;">
                    <h3 style="margin: 0 0 4px 0; font-size: 16px;">${typeLabel} in same ${parentType || 'Group'}</h3>
                    <p class="section-desc" style="margin: 0 0 16px 0; font-size: 12px; color: #6b7280;">${subtitle}</p>
                    <div class="nearby-list"></div>
                </div>
            `;
            const nearbyListEl = contextSection.querySelector('.nearby-list');

            if (siblings.length === 0) {
                nearbyListEl.innerHTML = '<div class="no-nearby">No other entities found in this group.</div>';
                return;
            }

            siblings.forEach(sibling => {
                const siblingUsers = data.mappings.userRoles.filter(m => m.entityId === sibling.id && m.role === activeRole);
                const assignedCount = siblingUsers.length;

                const siblingItem = document.createElement('div');
                siblingItem.className = 'nearby-item';
                siblingItem.innerHTML = `
                    <div class="nearby-info">
                        <div class="nearby-name">${sibling.name}</div>
                        <div class="nearby-status">${assignedCount} User(s) Assigned to this role</div>
                    </div>
                `;

                const editSiblingBtn = document.createElement('button');
                editSiblingBtn.className = 'edit-sibling-btn';
                editSiblingBtn.innerHTML = 'Switch <svg class="icon-xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>';
                editSiblingBtn.onclick = () => {
                    editingRoleEntityId = sibling.id;
                    renderLists();
                };
                siblingItem.appendChild(editSiblingBtn);
                nearbyListEl.appendChild(siblingItem);
            });
        };

        renderUsers();
        if (viewingUserMappingId) {
            renderUserDetails();
        } else {
            renderNearby();
        }
    };

    renderEditorContent();
}

// --- Final Setup ---
function setup() {
    if (isAppInitialized) return;
    isAppInitialized = true;
    initDOMElements();
    bindEvents();
    init();
}

document.addEventListener('DOMContentLoaded', setup);
if (document.readyState === 'complete' || document.readyState === 'interactive') {
    setup();
}
