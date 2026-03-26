// --- Utility: Format date as dd/mmm/yyyy (e.g., 21/Mar/2026) ---
function formatDateDDMMMYYYY(dateInput) {
    if (!dateInput) return '';
    try {
        let d = (dateInput instanceof Date) ? dateInput : _toDateOnly_(dateInput);
        if (d && !isNaN(d.getTime())) {
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const day = String(d.getDate()).padStart(2, '0');
            const mon = months[d.getMonth()];
            const year = d.getFullYear();
            return `${day}/${mon}/${year}`;
        }
    } catch (e) {}
    return String(dateInput);
}
// ============================================================
// main.js - PO Manager - Shared Logic
// Replace 'YOUR_GOOGLE_SCRIPT_URL' with your deployed Google Apps Script Web App URL
// ============================================================
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzigYT3lF6uQnnVO6boHFr54vKkmEEccKx_gCuc4b7wtF2zmcRcoynFszc40lkLLRxH/exec';

// ============================================================
// AUTH & SESSION MANAGEMENT
// ============================================================

// Fetch users from API, fallback to users.json, then hardcoded default
function fetchAppUsers(basePath) {
    var prefix = basePath || '';
    return fetch(GOOGLE_SCRIPT_URL + '?action=read_users', { redirect: 'follow' })
        .then(function(res) { return res.json(); })
        .then(function(data) {
            var users = Array.isArray(data) ? data : (data.users || []);
            if (users.length > 0) return users;
            return fetch(prefix + 'users.json').then(function(r) { return r.json(); }).catch(function() {
                return [{ username: 'ranga', password: 'Ranga@7677', role: 'admin' }];
            });
        })
        .catch(function() {
            return fetch(prefix + 'users.json').then(function(r) { return r.json(); }).catch(function() {
                return [{ username: 'ranga', password: 'Ranga@7677', role: 'admin' }];
            });
        });
}

// Save users to API (persists to Google Sheets)
function saveAppUsers(users) {
    return fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST', redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'save_users', users: users })
    }).then(function(res) { return res.json(); }).catch(function() { return { success: false }; });
}

function getSession() {
    return JSON.parse(localStorage.getItem('po_session') || 'null');
}

function isLoggedIn() {
    var s = getSession();
    return !!(s && s.username);
}

function isAdmin() {
    var s = getSession();
    return s && s.role === 'admin';
}

function getSessionRole() {
    var s = getSession();
    return (s && s.role ? String(s.role).toLowerCase() : 'viewer');
}

function canManagePO() {
    var role = getSessionRole();
    return role === 'admin' || role === 'editor';
}

function canManageMasters() {
    var role = getSessionRole();
    return role === 'admin' || role === 'editor';
}

function canBulkUploadPO() {
    return canManagePO();
}

function canBulkUploadMasters() {
    return canManageMasters();
}

function getSessionUser() {
    var s = getSession();
    return s ? s.username : '';
}

// Redirect to login if not authenticated. Call on every page load.
function requireAuth() {
    if (!isLoggedIn()) {
        // Determine correct login path based on current location
        var path = window.location.pathname;
        if (path.indexOf('/pages/') !== -1) {
            window.location.href = '../login.html';
        } else {
            window.location.href = 'login.html';
        }
        return false;
    }
    return true;
}

// Apply role-based UI restrictions after page renders
function applyRoleRestrictions() {
    var isAdminRole = isAdmin();
    var canPO = canManagePO();
    var canMaster = canManageMasters();

    if (!isAdminRole) {
        // Hide all elements marked as admin-only
        document.querySelectorAll('[data-admin-only], .admin-only').forEach(function(el) {
            el.style.display = 'none';
        });
    }

    if (!canMaster) {
        document.querySelectorAll('[data-master-edit]').forEach(function(el) {
            el.style.display = 'none';
        });
    }

    if (!canBulkUploadMasters()) {
        document.querySelectorAll('[data-master-bulk]').forEach(function(el) {
            el.style.display = 'none';
        });
    }

    if (!canPO) {
        document.querySelectorAll('[data-po-edit]').forEach(function(el) {
            el.style.display = 'none';
        });
    }

    if (!canBulkUploadPO()) {
        document.querySelectorAll('[data-po-bulk]').forEach(function(el) {
            el.style.display = 'none';
        });
    }

    // Hide "Create PO" sidebar link for users without PO edit rights
    if (!canPO) {
        document.querySelectorAll('.menu a').forEach(function(a) {
            if (a.getAttribute('href') && (a.getAttribute('href').indexOf('add-po') !== -1)) {
                var li = a.closest('li');
                if (li) li.style.display = 'none';
            }
        });
    }
}

// Inject user info + logout into the navbar
function injectNavbarAuth() {
    var s = getSession();
    if (!s) return;
    // Find the navbar's flex-1 title div and add user info + logout to the right
    var navbar = document.querySelector('.navbar');
    if (!navbar) return;

    // Check if already injected
    if (document.getElementById('navbarAuthBlock')) return;

    var profileHref = (window.location.pathname.indexOf('/pages/') !== -1) ? 'profile.html' : 'pages/profile.html';
    var div = document.createElement('div');
    div.id = 'navbarAuthBlock';
    div.className = 'flex items-center gap-2 ml-auto';
    div.innerHTML = '<a href="' + profileHref + '" class="flex items-center gap-1.5 text-sm hover:opacity-80 transition-opacity" title="My Profile">'
        + '<div class="avatar placeholder"><div class="bg-primary text-primary-content rounded-full w-7 h-7"><span class="text-xs">' + (s.username || '?').charAt(0).toUpperCase() + '</span></div></div>'
        + '<span class="font-medium hidden sm:inline">' + escHtml(s.username) + '</span>'
        + '<span class="badge badge-xs ' + (s.role === 'admin' ? 'badge-primary' : 'badge-ghost') + '">' + (s.role || 'viewer') + '</span>'
        + '</a>'
        + '<button class="btn btn-ghost btn-xs text-error" onclick="logout()" title="Sign Out">'
        + '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" /></svg>'
        + '</button>';
    navbar.appendChild(div);
}

// Inject sidebar links (Profile for all, User Management + Settings for admin)
function injectSidebarAuthLinks() {
    // Cleanup: if older logic injected sidebar links into dropdown menus, remove them.
    document.querySelectorAll('.dropdown-content .sidebar-profile-link, .dropdown-content .sidebar-users-link, .dropdown-content .sidebar-settings-link').forEach(function(el) {
        el.remove();
    });

    // Only target the actual sidebar navigation menu(s), not every .menu in the page.
    var menus = Array.from(document.querySelectorAll('.modern-sidebar > ul.menu, .drawer-side aside > ul.menu'));
    if (!menus.length) {
        menus = Array.from(document.querySelectorAll('.app-sidebar-shell ul.menu.menu-md'));
    }
    menus.forEach(function(menu) {
        if (menu.closest('.dropdown-content')) return;
        var path = window.location.pathname;
        var prefix = (path.indexOf('/pages/') !== -1) ? '' : 'pages/';

        // ── Profile link (for all users) ──
        if (!menu.querySelector('.sidebar-profile-link') && !menu.querySelector('a[href*="profile"]')) {
            var liP = document.createElement('li');
            liP.className = 'sidebar-profile-link';
            var isActiveP = path.indexOf('profile') !== -1 ? ' class="active"' : '';
            liP.innerHTML = '<a href="' + prefix + 'profile.html"' + isActiveP + '>'
                + '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" /></svg>'
                + 'My Profile'
                + '</a>';
            menu.appendChild(liP);
        }

        if (!isAdmin()) return;

        // ── User Management link ──
        if (!menu.querySelector('.sidebar-users-link') && !menu.querySelector('a[href*="user-management"]')) {
            var li = document.createElement('li');
            li.className = 'sidebar-users-link';
            var isActive = path.indexOf('user-management') !== -1 ? ' class="active"' : '';
            li.innerHTML = '<a href="' + prefix + 'user-management.html"' + isActive + '>'
                + '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197M13 7a4 4 0 11-8 0 4 4 0 018 0z" /></svg>'
                + 'User Management'
                + '</a>';
            menu.appendChild(li);
        }

        // ── Settings link ──
        if (!menu.querySelector('.sidebar-settings-link') && !menu.querySelector('a[href*="settings"]')) {
            var li2 = document.createElement('li');
            li2.className = 'sidebar-settings-link';
            var isActiveS = path.indexOf('settings') !== -1 ? ' class="active"' : '';
            li2.innerHTML = '<a href="' + prefix + 'settings.html"' + isActiveS + '>'
                + '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826-3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" /><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /></svg>'
                + 'Settings'
                + '</a>';
            menu.appendChild(li2);
        }
    });
}

function logout() {
    localStorage.removeItem('po_session');
    var path = window.location.pathname;
    if (path.indexOf('/pages/') !== -1) {
        window.location.href = '../login.html';
    } else {
        window.location.href = 'login.html';
    }
}

var SIDEBAR_COLLAPSE_KEY = 'po_sidebar_collapsed_desktop';

function _isCompactSidebarPage_() {
    var path = String((window.location && window.location.pathname) || '').toLowerCase();
    return path.indexOf('/pages/add-po.html') !== -1 || path.indexOf('/pages/po-edit.html') !== -1;
}

function _isMobileViewport_() {
    return window.matchMedia('(max-width: 767px)').matches;
}

function _isTabletViewport_() {
    return window.matchMedia('(min-width: 768px) and (max-width: 1023px)').matches;
}

function _isSidebarCollapsedDesktop_() {
    // Force compact sidebar on Create/Edit PO pages for better form workspace.
    if (_isCompactSidebarPage_()) return true;

    var stored = localStorage.getItem(SIDEBAR_COLLAPSE_KEY);
    // Default to expanded sidebar across the app when no preference exists.
    if (stored === null) return false;
    return stored === '1';
}

function _setSidebarCollapsedDesktop_(collapsed) {
    localStorage.setItem(SIDEBAR_COLLAPSE_KEY, collapsed ? '1' : '0');
}

function _isSidebarIconMode_() {
    if (_isTabletViewport_()) return true;
    if (_isMobileViewport_()) return false;
    var t = document.querySelector('.drawer-toggle');
    return !!(t && !t.checked);
}

function applySidebarModeClass() {
    if (!document.body) return;
    document.body.classList.toggle('sidebar-mobile-mode', _isMobileViewport_());
    document.body.classList.toggle('sidebar-tablet-mode', _isTabletViewport_());
    document.body.classList.toggle('sidebar-desktop-mode', !_isMobileViewport_() && !_isTabletViewport_());
    document.body.classList.toggle('sidebar-icon-mode', _isSidebarIconMode_());
}

function refreshSidebarLinkTooltips() {
    var iconMode = _isSidebarIconMode_();
    document.querySelectorAll('.drawer-side .menu a').forEach(function(a) {
        var label = a.textContent.replace(/\s+/g, ' ').trim();
        if (!label) return;
        a.setAttribute('aria-label', label);
        if (iconMode) {
            a.classList.add('tooltip', 'tooltip-right');
            a.setAttribute('data-tip', label);
            a.setAttribute('title', label);
        } else {
            a.classList.remove('tooltip', 'tooltip-right');
            a.removeAttribute('data-tip');
            a.removeAttribute('title');
        }
    });
}

function normalizeDrawerForViewport() {
    var isMobile = _isMobileViewport_();
    var isTablet = _isTabletViewport_();
    var collapsedDesktop = _isSidebarCollapsedDesktop_();
    var toggles = document.querySelectorAll('.drawer-toggle');
    toggles.forEach(function(t) {
        if (isMobile) t.checked = false;
        else if (isTablet) t.checked = false;
        else t.checked = !collapsedDesktop;
    });
    applySidebarModeClass();
    refreshSidebarLinkTooltips();
}

function forceCloseDrawerMobile() {
    if (!_isMobileViewport_()) return;
    document.querySelectorAll('.drawer-toggle').forEach(function(t) {
        t.checked = false;
    });
}

function bindSidebarInteractions() {
    if (window.__poSidebarInteractionsBound) return;
    window.__poSidebarInteractionsBound = true;

    function closeOnMobile() {
        if (!_isMobileViewport_()) return;
        document.querySelectorAll('.drawer-toggle').forEach(function(t) { t.checked = false; });
    }

    document.querySelectorAll('.drawer-toggle').forEach(function(t) {
        t.addEventListener('change', function() {
            if (_isMobileViewport_() || _isTabletViewport_()) {
                applySidebarModeClass();
                refreshSidebarLinkTooltips();
                return;
            }
            _setSidebarCollapsedDesktop_(!t.checked);
            applySidebarModeClass();
            refreshSidebarLinkTooltips();
        });
    });

    // Close drawer after selecting a sidebar link on mobile.
    document.querySelectorAll('.drawer-side .menu a').forEach(function(a) {
        a.addEventListener('click', function() {
            closeOnMobile();
        });
    });

    // Overlay click should always close drawer on mobile.
    document.querySelectorAll('.drawer-overlay').forEach(function(overlay) {
        overlay.addEventListener('click', function() {
            closeOnMobile();
        });
    });

    // Esc key closes open drawer on mobile.
    document.addEventListener('keydown', function(e) {
        if (e.key !== 'Escape') return;
        closeOnMobile();
    });

    // Re-evaluate tooltip/icon mode after sidebar links are injected.
    setTimeout(function() {
        applySidebarModeClass();
        refreshSidebarLinkTooltips();
    }, 80);
}

function bindResponsiveLayoutGuards() {
    if (window.__poResponsiveGuardsBound) return;
    window.__poResponsiveGuardsBound = true;

    // Some mobile browsers restore checkbox state on history navigation.
    window.addEventListener('pageshow', function() {
        forceCloseDrawerMobile();
    });

    var timer = null;
    window.addEventListener('resize', function() {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function() {
            forceCloseDrawerMobile();
            normalizeDrawerForViewport();
            applySidebarModeClass();
            refreshSidebarLinkTooltips();
        }, 120);
    });
}

// Master init for auth on every page — call from window.onload
function initPageAuth() {
    if (!requireAuth()) return false;
    bindSidebarInteractions();
    bindResponsiveLayoutGuards();
    forceCloseDrawerMobile();
    normalizeDrawerForViewport();
    setTimeout(forceCloseDrawerMobile, 60);
    requestAnimationFrame(forceCloseDrawerMobile);
    injectNavbarAuth();
    injectSidebarAuthLinks();
    applySidebarModeClass();
    refreshSidebarLinkTooltips();
    // Load seal/settings from backend (overrides seal.js if stored)
    loadSealFromSettings();
    // Defer role restrictions so page content renders first
    setTimeout(applyRoleRestrictions, 50);
    return true;
}

// Load company seal from Settings sheet (overrides hardcoded seal.js)
function loadSealFromSettings(callback) {
    gsFetchCached('read_settings').then(function(settings) {
        if (settings && settings.seal) {
            window.COMPANY_SEAL_IMG = settings.seal;
        }
        // Store other settings globally for use anywhere
        if (settings && typeof settings === 'object') {
            window._appSettings = settings;
        }
        if (typeof callback === 'function') callback(settings);
    }).catch(function() {
        // Fallback — keep seal.js value
        if (typeof callback === 'function') callback(null);
    });
}

// --- Utility: safe element getter ---
function $(id) { return document.getElementById(id); }

function showEmptyState(containerId, title, description, actionLabel, actionJs, iconSvg) {
    var el = $(containerId);
    if (!el) return;
    var icon = iconSvg || '<svg xmlns="http://www.w3.org/2000/svg" class="h-12 w-12" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" /></svg>';
    var button = '';
    if (actionLabel && actionJs) {
        button = '<button type="button" class="btn btn-sm btn-primary mt-2" onclick="' + actionJs + '">' + escHtml(actionLabel) + '</button>';
    }
    el.innerHTML = ''
        + '<div class="empty-state py-10 px-6 text-center">'
        + '  <div class="mx-auto mb-3 text-slate-300">' + icon + '</div>'
        + '  <h3 class="text-base font-semibold text-slate-700">' + escHtml(title || 'No data available') + '</h3>'
        + '  <p class="text-sm text-slate-500 mt-1 mb-2">' + escHtml(description || 'Try adjusting filters or adding new records.') + '</p>'
        + button
        + '</div>';
}

// --- Toast notification (replaces alert) ---
function showToast(message, type) {
    type = type || 'success';
    var c = $('toastContainer');
    if (!c) { c = document.createElement('div'); c.id = 'toastContainer'; c.className = 'toast-container'; document.body.appendChild(c); }
    var icons = {
        success: '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>',
        error: '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>',
        info: '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>',
        warning: '<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z" /></svg>'
    };
    var t = document.createElement('div');
    t.className = 'toast-item toast-' + type;
    t.innerHTML = (icons[type] || '') + '<span>' + message + '</span>';
    c.appendChild(t);
    setTimeout(function(){ t.classList.add('show'); }, 10);
    setTimeout(function(){ t.classList.remove('show'); setTimeout(function(){ t.remove(); }, 350); }, 3500);
}

// --- Set button loading state ---
function setBtnLoading(btn, loading) {
    if (!btn) return;
    if (loading) {
        btn.dataset.origText = btn.innerHTML;
        btn.innerHTML = '<span class="loading loading-spinner loading-sm"></span> Saving...';
        btn.disabled = true;
    } else {
        btn.innerHTML = btn.dataset.origText || btn.innerHTML;
        btn.disabled = false;
    }
}

// ============================================================
// PAGINATION UTILITY
// ============================================================
var _pagination = {};
var PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

function _initPagination(key, defaultSize) {
    if (!_pagination[key]) _pagination[key] = { page: 1, size: defaultSize || 10 };
    return _pagination[key];
}

function _buildPaginationBar(key, totalItems, renderFn) {
    var pg = _pagination[key];
    var totalPages = Math.max(1, Math.ceil(totalItems / pg.size));
    if (pg.page > totalPages) pg.page = totalPages;
    var renderFnName = (typeof renderFn === 'function' && renderFn.name)
        ? renderFn.name
        : String(renderFn || '');

    var start = (pg.page - 1) * pg.size + 1;
    var end = Math.min(pg.page * pg.size, totalItems);

    var html = '<div class="flex flex-wrap items-center justify-between gap-3 px-4 py-3 border-t border-base-300 bg-base-100 text-sm">';

    // Left: showing info + page size selector
    html += '<div class="flex items-center gap-3">';
    html += '<span class="text-base-content/60">Showing <b>' + start + '</b>–<b>' + end + '</b> of <b>' + totalItems + '</b></span>';
    html += '<select class="select select-bordered select-xs w-auto" onchange="_pgSize(\'' + key + '\', this.value, \'' + renderFnName + '\')">';
    PAGE_SIZE_OPTIONS.forEach(function(s) {
        html += '<option value="' + s + '"' + (pg.size === s ? ' selected' : '') + '>' + s + ' / page</option>';
    });
    html += '</select>';
    html += '</div>';

    // Right: page buttons
    html += '<div class="join">';
    html += '<button class="join-item btn btn-xs' + (pg.page <= 1 ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\', ' + 1 + ', \'' + renderFnName + '\')" title="First">&laquo;</button>';
    html += '<button class="join-item btn btn-xs' + (pg.page <= 1 ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\', ' + (pg.page - 1) + ', \'' + renderFnName + '\')" title="Previous">&lsaquo;</button>';

    // Page number buttons (show max 5)
    var startP = Math.max(1, pg.page - 2);
    var endP = Math.min(totalPages, startP + 4);
    if (endP - startP < 4) startP = Math.max(1, endP - 4);
    for (var p = startP; p <= endP; p++) {
        html += '<button class="join-item btn btn-xs' + (p === pg.page ? ' btn-active btn-primary' : '') + '" onclick="_pgGo(\'' + key + '\', ' + p + ', \'' + renderFnName + '\')">' + p + '</button>';
    }

    html += '<button class="join-item btn btn-xs' + (pg.page >= totalPages ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\', ' + (pg.page + 1) + ', \'' + renderFnName + '\')" title="Next">&rsaquo;</button>';
    html += '<button class="join-item btn btn-xs' + (pg.page >= totalPages ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\', ' + totalPages + ', \'' + renderFnName + '\')" title="Last">&raquo;</button>';
    html += '</div>';

    html += '</div>';
    return html;
}

function _resolveRenderFn_(renderFnOrName) {
    if (typeof renderFnOrName === 'function') return renderFnOrName;
    if (typeof renderFnOrName === 'string' && typeof window[renderFnOrName] === 'function') {
        return window[renderFnOrName];
    }
    return null;
}

function _pgGo(key, page, renderFnOrName) {
    _pagination[key].page = page;
    var fn = _resolveRenderFn_(renderFnOrName);
    if (fn) fn();
}

function _pgSize(key, size, renderFnOrName) {
    _pagination[key].size = parseInt(size);
    _pagination[key].page = 1;
    var fn = _resolveRenderFn_(renderFnOrName);
    if (fn) fn();
}

function _pgSlice(key, data) {
    var pg = _pagination[key];
    var start = (pg.page - 1) * pg.size;
    return data.slice(start, start + pg.size);
}

// --- Shared PDF Document Builder ---
// Takes a data object and returns a jsPDF instance
// d.includeSign: if true, embed the company seal/signature image
function buildPODocument(d) {
    if (!window.jspdf) return null;
    var items = d.items || [];
    if (typeof items === 'string') {
        try { items = JSON.parse(items); } catch(e) { items = []; }
    }
    var transport = parseFloat(d.transport) || 0;
    var itemsTotal = items.reduce(function(s, i) { return s + (parseFloat(i.total) || 0); }, 0);
    var subtotal = itemsTotal + transport;
    var cgstPct = parseFloat(d.cgst_percent) || 0;
    var sgstPct = parseFloat(d.sgst_percent) || 0;
    var cgstAmt = Math.round(subtotal * cgstPct / 100);
    var sgstAmt = Math.round(subtotal * sgstPct / 100);
    var grandTotal = Math.round(subtotal + cgstAmt + sgstAmt);
    var totalQty = items.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);

    var jsPDF = window.jspdf.jsPDF;
    var doc = new jsPDF('p', 'mm', 'a4');
    var pageW = doc.internal.pageSize.getWidth();
    var y = 12;

    // Company Header
    doc.setFontSize(16); doc.setFont(undefined, 'bold');
    doc.text('RANGA ELECTRICALS PVT LTD', pageW / 2, y, { align: 'center' }); y += 6;
    doc.setFontSize(8); doc.setFont(undefined, 'normal');
    doc.text('NO.326, RAMAKRISHNA NAGAR MAIN ROAD, PORUR, CHENNAI-600116.', pageW / 2, y, { align: 'center' }); y += 4;
    doc.text('E-mail: info@rangaelectrical.com', pageW / 2, y, { align: 'center' }); y += 4;
    doc.setFont(undefined, 'bold');
    doc.text('GST IN: 33AAHCR4037J1ZD', pageW / 2, y, { align: 'center' }); y += 6;
    doc.setDrawColor(0); doc.line(14, y, pageW - 14, y); y += 5;

    doc.setFontSize(12); doc.setFont(undefined, 'bold');
    doc.text('PURCHASE ORDER', pageW / 2, y, { align: 'center' }); y += 8;

    // ── Two card boxes side by side ──
    var margin = 14;
    var gap = 4;
    var boxW = (pageW - margin * 2 - gap) / 2;
    var leftBoxX = margin;
    var rightBoxX = margin + boxW + gap;
    var boxStartY = y;

    // --- Left box: Vendor details ---
    var ly = boxStartY + 5;
    doc.setFontSize(9); doc.setFont(undefined, 'normal');
    doc.setTextColor(120, 120, 120);
    doc.text('To,', leftBoxX + 4, ly); ly += 5;
    doc.setTextColor(0, 0, 0);
    doc.setFont(undefined, 'bold'); doc.setFontSize(10);
    doc.text(d.vendor_name || '________', leftBoxX + 4, ly); ly += 5;
    doc.setFont(undefined, 'normal'); doc.setFontSize(8);
    if (d.vendor_address) {
        var addrLines = doc.splitTextToSize(d.vendor_address, boxW - 8);
        addrLines.forEach(function(l) { doc.text(l, leftBoxX + 4, ly); ly += 3.5; });
    }
    if (d.vendor_gstin) { ly += 1; doc.text('GSTIN: ' + d.vendor_gstin, leftBoxX + 4, ly); ly += 4; }
    var leftBoxH = ly - boxStartY + 3;

    // --- Right box: PO details ---
    var ry = boxStartY + 5;
    doc.setFontSize(9); doc.setFont(undefined, 'normal');
    doc.setTextColor(100, 100, 100);
    doc.text('PO NO:', rightBoxX + 4, ry);
    doc.setTextColor(37, 99, 235);
    doc.setFont(undefined, 'bold');
    doc.text(d.po_no || '____', rightBoxX + boxW - 4, ry, { align: 'right' }); ry += 5;
    doc.setTextColor(100, 100, 100);
    doc.setFont(undefined, 'normal');
    doc.text('DATE:', rightBoxX + 4, ry);
    doc.setTextColor(0, 0, 0);
    doc.setFont(undefined, 'bold');
    doc.text(formatDateDDMMMYYYY(d.po_date) || '____', rightBoxX + boxW - 4, ry, { align: 'right' }); ry += 6;
    doc.setFont(undefined, 'normal'); doc.setFontSize(8);
    doc.setTextColor(0, 0, 0);
    if (d.client_name) {
        doc.text('Client Name:', rightBoxX + 4, ry);
        doc.setFont(undefined, 'bold');
        doc.text('M/s. ' + d.client_name, rightBoxX + 28, ry);
        doc.setFont(undefined, 'normal');
        ry += 5;
    }
    if (d.event_date) {
        doc.text('Event Date:', rightBoxX + 4, ry);
        doc.setFont(undefined, 'bold');
        doc.text(formatDateDDMMMYYYY(d.event_date), rightBoxX + 28, ry);
        doc.setFont(undefined, 'normal');
        ry += 5;
    }
    var rightBoxH = ry - boxStartY + 3;

    // Use the taller box height for both
    var boxH = Math.max(leftBoxH, rightBoxH, 28);

    // Draw rounded rect borders
    doc.setDrawColor(200, 200, 200);
    doc.setLineWidth(0.4);
    doc.roundedRect(leftBoxX, boxStartY, boxW, boxH, 2, 2, 'S');
    doc.roundedRect(rightBoxX, boxStartY, boxW, boxH, 2, 2, 'S');

    y = boxStartY + boxH + 6;

    // Instruction line
    doc.setFont(undefined, 'italic'); doc.setFontSize(8);
    doc.setTextColor(0, 0, 0);
    doc.text('We are pleased to place our order in your favour for the supply of following items', margin, y); y += 6;

    var rightX = pageW - 14;

    // Items table
    var tableRows = items.map(function(item, idx) {
        return [
            idx + 1, item.desc || '', item.qty || '', item.uom || '',
            Number(item.per_day || 0).toFixed(2), Number(item.days || 0).toFixed(2),
            Number(item.total || 0).toFixed(2)
        ];
    });
    if (transport > 0) {
        tableRows.push([items.length + 1, 'TRANSPORT CHARGES', '', '', '', '', transport.toFixed(2)]);
    }

    var pageH = doc.internal.pageSize.getHeight();

    if (tableRows.length) {
        doc.autoTable({
            startY: y,
            head: [['Sl.No', 'Item Description', 'Qty', 'UOM', 'Per Day Amt', 'Days', 'Total (INR)']],
            body: tableRows,
            foot: [['', 'Total', String(totalQty), '', '', '', subtotal.toFixed(2)]],
            showFoot: 'lastPage',
            showHead: 'everyPage',
            styles: { fontSize: 8, cellPadding: 1.5 },
            headStyles: { fillColor: [41, 128, 185], textColor: 255, fontStyle: 'bold' },
            footStyles: { fillColor: [236, 240, 241], textColor: [0, 0, 0], fontStyle: 'bold' },
            columnStyles: {
                0: { halign: 'center', cellWidth: 12 },
                2: { halign: 'right', cellWidth: 14 },
                4: { halign: 'right', cellWidth: 22 },
                5: { halign: 'right', cellWidth: 14 },
                6: { halign: 'right', cellWidth: 26 }
            },
            margin: { left: 14, right: 14 },
            didDrawPage: function(data) {
                // Repeat company header on continuation pages
                if (data.pageNumber > 1) {
                    doc.setFontSize(7); doc.setFont(undefined, 'italic');
                    doc.setTextColor(150, 150, 150);
                    doc.text('RANGA ELECTRICALS PVT LTD — PO: ' + (d.po_no || ''), pageW / 2, 8, { align: 'center' });
                    doc.setTextColor(0, 0, 0);
                }
            }
        });
        y = doc.lastAutoTable.finalY + 4;
    } else {
        doc.setFontSize(9); doc.setFont(undefined, 'normal');
        doc.text('(No items added yet)', 14, y); y += 8;
    }

    // Estimate space needed for footer block: GST lines + Grand Total + Terms + Signature
    var footerNeeded = 0;
    if (cgstPct) footerNeeded += 5;
    if (sgstPct) footerNeeded += 5;
    footerNeeded += 8;  // Grand Total
    if (d.terms) footerNeeded += 6;
    footerNeeded += 45; // divider + EVENT LOCATION + signature block

    if (y + footerNeeded > pageH - 10) {
        doc.addPage();
        y = 15;
    }

    // GST
    doc.setFontSize(9); doc.setFont(undefined, 'normal');
    if (cgstPct) { doc.text('Add : CGST @ ' + cgstPct + '%', rightX - 60, y); doc.text(cgstAmt.toFixed(2), rightX, y, { align: 'right' }); y += 5; }
    if (sgstPct) { doc.text('Add : SGST @ ' + sgstPct + '%', rightX - 60, y); doc.text(sgstAmt.toFixed(2), rightX, y, { align: 'right' }); y += 5; }
    doc.setFont(undefined, 'bold'); doc.setFontSize(10);
    doc.text('GRAND TOTAL', rightX - 60, y); doc.text(grandTotal.toFixed(2), rightX, y, { align: 'right' }); y += 8;

    // Terms
    doc.setFontSize(8); doc.setFont(undefined, 'normal');
    if (d.terms) { doc.text('TERMS OF PAYMENT: ' + d.terms, 14, y); y += 6; }

    // Ensure enough space for signature block; push to new page if needed
    if (y + 35 > pageH - 10) {
        doc.addPage();
        y = 15;
    }

    // Event location & signature
    doc.setDrawColor(0); doc.line(14, y, pageW - 14, y); y += 5;
    var signStartY = y; // remember Y for right-side signature block
    doc.setFont(undefined, 'bold'); doc.text('EVENT LOCATION', 14, y);
    doc.setFont(undefined, 'normal'); doc.text('for RANGA ELECTRICALS PVT LTD', rightX, y, { align: 'right' }); y += 5;
    if (d.event_name) { doc.setFont(undefined, 'bold'); doc.text(d.event_name, 14, y); y += 4; }
    if (d.event_location) {
        doc.setFont(undefined, 'normal');
        var locLines = doc.splitTextToSize(d.event_location, 80);
        locLines.forEach(function(l) { doc.text(l, 14, y); y += 3.5; });
    }

    // Right side: seal + signature block (independent of left side y)
    var signCenterX = rightX - 22;
    var sealY = signStartY + 4;

    // Add company seal/signature image if requested
    if (d.includeSign && typeof COMPANY_SEAL_IMG === 'string') {
        try { doc.addImage(COMPANY_SEAL_IMG, 'PNG', signCenterX - 12.5, sealY, 25, 25); } catch(e) {}
    }

    // Signature line + text below seal
    var sigLineY = sealY + (d.includeSign && typeof COMPANY_SEAL_IMG === 'string' ? 27 : 20);
    doc.setDrawColor(150);
    doc.line(signCenterX - 22, sigLineY, signCenterX + 22, sigLineY);
    doc.setFont(undefined, 'bold'); doc.setFontSize(8);
    doc.text('AUTHORISED SIGNATURE', signCenterX, sigLineY + 4, { align: 'center' });

    // Ensure y is past both left and right content
    y = Math.max(y, sigLineY + 8);

    return doc;
}

// --- Utility: CORS-safe fetch for Google Apps Script ---
function gsFetch(url, options) {
    return fetch(url, { redirect: 'follow', ...options });
}
function gsPost(data) {
    return fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(data)
    }).then(res => res.text()).then(text => {
        try {
            return JSON.parse(text);
        } catch (e) {
            console.error('Server response (not JSON):', text);
            return { success: false, error: 'Invalid response from server' };
        }
    });
}

// ============================================================
// RESPONSE CACHE — avoid redundant API calls (30s TTL)
// ============================================================
var _apiCache = {};
var _CACHE_TTL = 30000; // 30 seconds

function gsFetchCached(action) {
    var now = Date.now();
    if (_apiCache[action] && (now - _apiCache[action].ts < _CACHE_TTL)) {
        return Promise.resolve(_apiCache[action].data);
    }
    return gsFetch(GOOGLE_SCRIPT_URL + '?action=' + action)
        .then(function(res) { return res.json(); })
        .then(function(data) {
            _apiCache[action] = { data: data, ts: Date.now() };
            return data;
        });
}

function invalidateCache(action) {
    if (action) { delete _apiCache[action]; }
    else { _apiCache = {}; }
}

// Bulk fetch — single API call for POs + Vendors + Items (used by dashboard, add-po, etc.)
function fetchAll(callback) {
    gsFetchCached('read_all').then(function(data) {
        if (data && data.vendors) {
            vendorMasterList = (data.vendors || []).map(function(v) {
                return { name: v.name || '', address: v.address || '', gstin: v.gstin || '' };
            });
        }
        if (data && data.items) {
            itemMasterCache = data.items || [];
        }
        if (data && data.pos) {
            window._poList = data.pos;
        }
        if (callback) callback(data);
    }).catch(function(err) {
        console.error('fetchAll error:', err);
        if (callback) callback(null);
    });
}

// Read first worksheet from CSV/XLS/XLSX and return rows as key-value objects.
function readSpreadsheetRows(file, onSuccess) {
    if (!file) return;
    if (typeof XLSX === 'undefined') {
        showToast('Excel library is not loaded on this page.', 'error');
        return;
    }
    var reader = new FileReader();
    reader.onload = function(e) {
        try {
            var data = e.target.result;
            var wb = XLSX.read(data, { type: 'array' });
            var sheetName = wb.SheetNames[0];
            var ws = wb.Sheets[sheetName];
            var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
            onSuccess(rows);
        } catch (err) {
            showToast('Unable to read file: ' + err.message, 'error');
        }
    };
    reader.onerror = function() {
        showToast('Unable to read file.', 'error');
    };
    reader.readAsArrayBuffer(file);
}

function ensureEditPopupModal_() {
    var dlg = document.getElementById('appEditPopup');
    if (dlg) return dlg;

    dlg = document.createElement('dialog');
    dlg.id = 'appEditPopup';
    dlg.className = 'modal modal-bottom sm:modal-middle';
    dlg.innerHTML = ''
        + '<div class="modal-box max-w-xl">'
        + '  <h3 id="appEditPopupTitle" class="font-bold text-lg">Edit</h3>'
        + '  <form id="appEditPopupForm" class="mt-4 space-y-3"></form>'
        + '  <div class="modal-action">'
        + '    <button type="button" id="appEditPopupCancel" class="btn btn-ghost">Cancel</button>'
        + '    <button type="button" id="appEditPopupSave" class="btn btn-primary">Save</button>'
        + '  </div>'
        + '</div>'
        + '<form method="dialog" class="modal-backdrop"><button>close</button></form>';
    document.body.appendChild(dlg);
    return dlg;
}

function showEditPopup_(opts) {
    return new Promise(function(resolve) {
        var dlg = ensureEditPopupModal_();
        var titleEl = document.getElementById('appEditPopupTitle');
        var formEl = document.getElementById('appEditPopupForm');
        var btnSave = document.getElementById('appEditPopupSave');
        var btnCancel = document.getElementById('appEditPopupCancel');

        titleEl.textContent = opts.title || 'Edit';
        formEl.innerHTML = (opts.fields || []).map(function(f) {
            var type = f.type || 'text';
            var value = (f.value === undefined || f.value === null) ? '' : String(f.value);
            var req = f.required ? ' required' : '';
            var placeholder = f.placeholder ? (' placeholder="' + escHtml(String(f.placeholder)) + '"') : '';
            if (type === 'textarea') {
                return ''
                    + '<div class="form-control">'
                    + '  <label class="label"><span class="label-text font-medium">' + escHtml(f.label || f.name) + '</span></label>'
                    + '  <textarea class="textarea textarea-bordered" name="' + escHtml(f.name) + '"' + req + placeholder + '>' + escHtml(value) + '</textarea>'
                    + '</div>';
            }
            return ''
                + '<div class="form-control">'
                + '  <label class="label"><span class="label-text font-medium">' + escHtml(f.label || f.name) + '</span></label>'
                + '  <input class="input input-bordered" type="' + escHtml(type) + '" name="' + escHtml(f.name) + '" value="' + escHtml(value) + '"' + req + placeholder + '>'
                + '</div>';
        }).join('');

        function cleanup() {
            btnSave.onclick = null;
            btnCancel.onclick = null;
            dlg.onclose = null;
        }

        btnCancel.onclick = function() {
            cleanup();
            dlg.close();
            resolve(null);
        };

        btnSave.onclick = function() {
            var data = {};
            var valid = true;
            (opts.fields || []).forEach(function(f) {
                var el = formEl.querySelector('[name="' + f.name + '"]');
                if (!el) return;
                var val = String(el.value || '').trim();
                if (f.required && !val) valid = false;
                data[f.name] = val;
            });
            if (!valid) {
                showToast('Please fill all required fields.', 'warning');
                return;
            }
            cleanup();
            dlg.close();
            resolve(data);
        };

        dlg.onclose = function() {
            cleanup();
        };

        dlg.showModal();
    });
}

function _downloadCsvFile_(fileName, headers, rows) {
    var esc = function(v) {
        var s = String(v === undefined || v === null ? '' : v);
        if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
        return s;
    };
    var csv = [headers.map(esc).join(',')]
        .concat((rows || []).map(function(r) { return r.map(esc).join(','); }))
        .join('\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function() { URL.revokeObjectURL(a.href); }, 500);
}

async function _downloadXlsxFile_(fileName, sheetName, headers, rows) {
    if (typeof ExcelJS !== 'undefined') {
        try {
            var wbStyled = new ExcelJS.Workbook();
            var wsStyled = wbStyled.addWorksheet(sheetName || 'Sample');
            var dataRows = rows || [];

            wsStyled.columns = (headers || []).map(function(h) {
                var width = Math.max(12, String(h || '').length + 4);
                return { width: width };
            });

            wsStyled.addRow(headers || []);
            var hr = wsStyled.getRow(1);
            hr.height = 22;
            hr.eachCell(function(cell) {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                    left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                    bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                    right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
                };
            });

            dataRows.forEach(function(r) { wsStyled.addRow(r); });
            for (var i = 2; i <= wsStyled.rowCount; i++) {
                var fill = (i % 2 === 0) ? 'FFF8FAFC' : 'FFEFF6FF';
                wsStyled.getRow(i).eachCell(function(cell) {
                    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                    };
                });
            }

            wsStyled.views = [{ state: 'frozen', ySplit: 1 }];
            var buffer = await wbStyled.xlsx.writeBuffer();
            _downloadBlobFile_(fileName, new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
            return;
        } catch (e) {
            console.error('Styled sample export failed, fallback to SheetJS:', e);
        }
    }

    if (typeof XLSX === 'undefined') {
        showToast('Excel library is not loaded on this page.', 'error');
        return;
    }
    var aoa = [headers].concat(rows || []);
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Sample');
    XLSX.writeFile(wb, fileName);
}

function downloadBulkSample(type, format) {
    var kind = String(type || '').toLowerCase();
    var fmt = String(format || 'csv').toLowerCase();
    var headers = [];
    var rows = [];
    var baseName = 'bulk_sample';

    if (kind === 'vendor') {
        headers = ['name', 'address', 'gstin'];
        rows = [
            ['ABC Power Systems', 'No.12, Industrial Estate, Chennai', '33ABCDE1234F1Z5'],
            ['Prime Rentals', '45, Market Road, Coimbatore', '33PQRSX6789L1Z2']
        ];
        baseName = 'vendor_bulk_sample';
    } else if (kind === 'item') {
        headers = ['desc', 'uom', 'rate', 'default_qty'];
        rows = [
            ['DG Cable 50m', 'NOS', 1500, 1],
            ['AMF Panel', 'NOS', 2500, 1]
        ];
        baseName = 'item_bulk_sample';
    } else if (kind === 'po') {
        headers = ['po_no', 'po_date', 'vendor_name', 'vendor_address', 'vendor_gstin', 'client_name', 'event_name', 'event_location', 'event_date', 'transport', 'cgst_percent', 'sgst_percent', 'terms', 'item_desc', 'qty', 'uom', 'rate', 'days'];
        rows = [
            ['RE/EO/101', '2026-03-14', 'ABC Power Systems', 'No.12, Industrial Estate, Chennai', '33ABCDE1234F1Z5', 'RC Electricals', 'Expo Setup', 'Chennai Trade Center', '15-03-2026 TO 18-03-2026', 5000, 9, 9, '60 DAYS FROM THE DATE OF SUPPLY', 'DG Cable 50m', 2, 'NOS', 1500, 4],
            ['RE/EO/102', '2026-03-15', 'Prime Rentals', '45, Market Road, Coimbatore', '33PQRSX6789L1Z2', 'Skyline Events', 'Concert', 'Madurai', '20-03-2026 TO 22-03-2026', 3000, 9, 9, '50% ADVANCE', 'AMF Panel', 1, 'NOS', 2500, 3]
        ];
        baseName = 'po_bulk_sample';
    } else {
        showToast('Unknown sample type.', 'error');
        return;
    }

    if (fmt === 'xlsx') {
        _downloadXlsxFile_(baseName + '.xlsx', 'Sample', headers, rows);
    } else {
        _downloadCsvFile_(baseName + '.csv', headers, rows);
    }
}

// ============================================================
// VENDOR MASTER
// ============================================================
let vendorMasterList = []; // Array of {name, address, gstin}

function fetchVendors() {
    gsFetchCached('read_vendors')
        .then(data => {
            vendorMasterList = (data || []).map(v => ({
                name: v.name || '',
                address: v.address || '',
                gstin: v.gstin || ''
            }));
            if ($('vendorList')) renderVendorList(data);
            // Init searchable vendor dropdown on PO form
            if ($('vendorDropdownWrap')) initVendorDropdown('vendorDropdownWrap', 'vendor_name', 'vendorDDList');
        })
        .catch(() => {
            if ($('vendorList')) $('vendorList').innerHTML = '<div class="text-error p-4">Failed to load vendors</div>';
        });
}

// ---- Searchable Vendor Dropdown ----
function initVendorDropdown(wrapId, inputId, listId) {
    const wrap = $(wrapId);
    const input = $(inputId);
    const list = $(listId);
    if (!wrap || !input || !list) return;

    let activeIdx = -1;

    function renderList(filter) {
        const q = (filter || '').toLowerCase().trim();
        const filtered = q
            ? vendorMasterList.filter(v => v.name.toLowerCase().includes(q))
            : vendorMasterList;

        // Check if typed name exactly matches any vendor
        const exactMatch = q && vendorMasterList.some(v => v.name.toLowerCase() === q);

        if (!filtered.length) {
            list.innerHTML = '<div class="vendor-dd-empty">No vendors found</div>'
                + (q ? '<div class="vendor-dd-add" data-action="quick-add" data-value="' + escHtml(input.value.trim()) + '"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" /></svg> Add "' + escHtml(input.value.trim()) + '" as new vendor</div>' : '');
        } else {
            let html = filtered.map((v, i) => {
                const highlighted = q ? v.name.replace(new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi'), '<span class="match-highlight">$1</span>') : v.name;
                const subInfo = [v.address, v.gstin].filter(Boolean).join(' | ');
                return '<div class="vendor-dd-item" data-idx="' + i + '" data-value="' + escHtml(v.name) + '">'
                    + '<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" /></svg>'
                    + '<div><span>' + highlighted + '</span>'
                    + (subInfo ? '<div class="text-xs opacity-50" style="line-height:1.2">' + escHtml(subInfo) + '</div>' : '')
                    + '</div></div>';
            }).join('');
            // Show quick-add option if typed text is not an exact match
            if (q && !exactMatch) {
                html += '<div class="vendor-dd-add" data-action="quick-add" data-value="' + escHtml(input.value.trim()) + '"><svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" /></svg> Add "' + escHtml(input.value.trim()) + '" as new vendor</div>';
            }
            list.innerHTML = html;
        }
        activeIdx = -1;
    }

    function open() { wrap.classList.add('open'); renderList(input.value); }
    function close() { wrap.classList.remove('open'); activeIdx = -1; }
    function select(val) {
        input.value = val;
        // Auto-fill vendor address & GSTIN from master
        autoFillVendorDetails(val);
        close();
        input.focus();
        // Trigger live preview if available
        if (typeof generateLivePreview === 'function') generateLivePreview();
    }

    input.addEventListener('focus', open);
    input.addEventListener('input', () => { open(); });
    input.addEventListener('keydown', (e) => {
        const items = list.querySelectorAll('.vendor-dd-item');
        if (!items.length) return;
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            activeIdx = Math.min(activeIdx + 1, items.length - 1);
            items.forEach((el, i) => el.classList.toggle('active', i === activeIdx));
            items[activeIdx].scrollIntoView({ block: 'nearest' });
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            activeIdx = Math.max(activeIdx - 1, 0);
            items.forEach((el, i) => el.classList.toggle('active', i === activeIdx));
            items[activeIdx].scrollIntoView({ block: 'nearest' });
        } else if (e.key === 'Enter' && activeIdx >= 0) {
            e.preventDefault();
            select(items[activeIdx].getAttribute('data-value'));
        } else if (e.key === 'Escape') {
            close();
        }
    });

    list.addEventListener('mousedown', (e) => {
        // Quick-add button in dropdown
        const addBtn = e.target.closest('.vendor-dd-add');
        if (addBtn) {
            e.preventDefault();
            const val = addBtn.getAttribute('data-value');
            if (val) quickAddVendor(val, input, wrap);
            return;
        }
        const item = e.target.closest('.vendor-dd-item');
        if (item) { e.preventDefault(); select(item.getAttribute('data-value')); }
    });

    document.addEventListener('click', (e) => {
        if (!wrap.contains(e.target)) close();
    });
}

// Auto-fill vendor address & GSTIN from master data
function autoFillVendorDetails(vendorName) {
    const vendor = vendorMasterList.find(v => v.name.toLowerCase() === (vendorName || '').toLowerCase());
    const addrEl = $('vendor_address');
    const gstEl = $('vendor_gstin');
    if (vendor) {
        if (addrEl && vendor.address) addrEl.value = vendor.address;
        if (gstEl && vendor.gstin) gstEl.value = vendor.gstin;
    }
}

function escHtml(s) {
    return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// Quick add vendor to master and refresh dropdown
function quickAddVendor(name, inputEl, wrapEl) {
    const trimmed = (name || '').trim();
    if (!trimmed) return;
    // Also grab address/gstin from form if available
    const address = ($('vendor_address') || {}).value || '';
    const gstin = ($('vendor_gstin') || {}).value || '';
    gsPost({ action: 'add_vendor', name: trimmed, address: address.trim(), gstin: gstin.trim() })
    .then(resp => {
        if (resp.success) {
            invalidateCache('read_vendors'); invalidateCache('read_all');
            vendorMasterList.push({ name: trimmed, address: address.trim(), gstin: gstin.trim() });
            if (inputEl) inputEl.value = trimmed;
            if (wrapEl) wrapEl.classList.remove('open');
            // Refresh vendor list on Vendor Master page if visible
            if ($('vendorList')) fetchVendors();
        } else {
            alert('Could not add vendor: ' + (resp.error || 'Unknown error'));
        }
    })
    .catch(err => alert('Error adding vendor: ' + err.message));
}

// Quick add from the + button next to vendor input
function quickAddVendorFromInput() {
    const input = $('vendor_name');
    const wrap = $('vendorDropdownWrap');
    if (!input) return;
    const name = input.value.trim();
    if (!name) { alert('Type a vendor name first'); input.focus(); return; }
    if (vendorMasterList.some(v => v.name.toLowerCase() === name.toLowerCase())) {
        alert('Vendor "' + name + '" already exists in master');
        return;
    }
    quickAddVendor(name, input, wrap);
}

var _vendorData = [];
var _vendorSearchTerm = '';
function filterVendorList(term) {
    _vendorSearchTerm = (term || '').toLowerCase().trim();
    _pagination['vendor'] = { page: 1, size: (_pagination['vendor'] || {}).size || 10 };
    renderVendorList();
}
function renderVendorList(data) {
    if (data) _vendorData = data;
    var allData = _vendorData || [];
    // Apply search filter
    var filtered = allData;
    if (_vendorSearchTerm) {
        filtered = allData.filter(function(v) {
            return (v.name || '').toLowerCase().includes(_vendorSearchTerm)
                || (v.address || '').toLowerCase().includes(_vendorSearchTerm)
                || (v.gstin || '').toLowerCase().includes(_vendorSearchTerm);
        });
    }
    if (!Array.isArray(allData) || allData.length === 0) {
        $('vendorList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z" /></svg><p class="text-sm">No vendors found. Add your first vendor above.</p></div>';
        return;
    }
    if (filtered.length === 0 && _vendorSearchTerm) {
        $('vendorList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg><p class="text-sm">No vendors match "<strong>' + escHtml(_vendorSearchTerm) + '</strong>"</p></div>';
        return;
    }
    // Hero stats
    if ($('vendorHeroStats')) {
        var statsHtml = '<div class="hero-stat"><div class="stat-label">Total Vendors</div><div class="stat-value">' + allData.length + '</div></div>';
        if (_vendorSearchTerm) statsHtml += '<div class="hero-stat"><div class="stat-label">Matched</div><div class="stat-value">' + filtered.length + '</div></div>';
        $('vendorHeroStats').innerHTML = statsHtml;
    }
    var pg = _initPagination('vendor', 10);
    var total = filtered.length;
    var startIdx = (pg.page - 1) * pg.size;
    var pageData = _pgSlice('vendor', filtered);

    var canEditMaster = canManageMasters();
    let html = '<table class="table table-zebra w-full mobile-card-table"><thead><tr><th>#</th><th>Vendor Name</th><th>Address</th><th>GSTIN</th>' + (canEditMaster ? '<th>Action</th>' : '') + '</tr></thead><tbody>';
    pageData.forEach((vendor, i) => {
        var displayIdx = startIdx + i;
        var origIdx = allData.indexOf(vendor);
        html += '<tr><td data-label="#" class="font-medium text-slate-500">' + (displayIdx + 1) + '</td><td data-label="Vendor Name" class="font-semibold text-slate-800">' + (vendor.name || '') + '</td><td data-label="Address" class="max-w-xs truncate text-slate-600">' + (vendor.address || '-') + '</td><td data-label="GSTIN"><code class="text-xs bg-slate-100 px-1.5 py-0.5 rounded">' + (vendor.gstin || '-') + '</code></td>' + (canEditMaster ? '<td data-label="Actions" class="flex gap-1"><button class="btn btn-xs btn-warning btn-outline btn-modern" onclick="editVendor(' + origIdx + ')">Edit</button><button class="btn btn-xs btn-error btn-outline btn-modern" onclick="deleteVendor(' + origIdx + ')">Delete</button></td>' : '') + '</tr>';
    });
    html += '</tbody></table>';
    if (total > pg.size) html += _buildPaginationBar('vendor', total, renderVendorList);
    $('vendorList').innerHTML = html;
}

function addVendor() {
    if (!canManageMasters()) return;
    const name = $('vendorNameInput').value.trim();
    const address = ($('vendorAddressInput') || {}).value || '';
    const gstin = ($('vendorGstinInput') || {}).value || '';
    if (!name) return;
    gsPost({ action: 'add_vendor', name, address: address.trim(), gstin: gstin.trim() })
    .then(resp => {
        if (resp.success) {
            invalidateCache('read_vendors'); invalidateCache('read_all');
            fetchVendors();
            $('vendorNameInput').value = '';
            if ($('vendorAddressInput')) $('vendorAddressInput').value = '';
            if ($('vendorGstinInput')) $('vendorGstinInput').value = '';
        }
        else alert('Failed to add vendor: ' + (resp.error || 'Unknown error'));
    })
    .catch(err => alert('Error adding vendor: ' + err.message));
}

function editVendor(idx) {
    if (!canManageMasters()) return;
    var v = (_vendorData || [])[idx];
    if (!v) return;

    showEditPopup_({
        title: 'Edit Vendor',
        fields: [
            { name: 'name', label: 'Vendor Name', value: v.name || '', required: true },
            { name: 'address', label: 'Address', type: 'textarea', value: v.address || '' },
            { name: 'gstin', label: 'GSTIN', value: v.gstin || '' }
        ]
    }).then(function(data) {
        if (!data) return;
        gsPost({ action: 'update_vendor', idx: idx, name: data.name, address: data.address, gstin: data.gstin })
        .then(function(resp) {
            if (resp.success) {
                invalidateCache('read_vendors');
                invalidateCache('read_all');
                fetchVendors();
                showToast('Vendor updated.', 'success');
            } else {
                showToast('Failed to update vendor: ' + (resp.error || 'Unknown error'), 'error');
            }
        })
        .catch(function(err) {
            showToast('Error updating vendor: ' + err.message, 'error');
        });
    });
}

function deleteVendor(idx) {
    if (!canManageMasters()) return;
    if (!confirm('Delete this vendor?')) return;
    gsPost({ action: 'delete_vendor', idx })
    .then(resp => { if (resp.success) { invalidateCache('read_vendors'); invalidateCache('read_all'); fetchVendors(); } else alert('Failed to delete vendor: ' + (resp.error || 'Unknown error')); })
    .catch(err => alert('Error deleting vendor: ' + err.message));
}

// ============================================================
// BULK UPLOAD — VERIFICATION PREVIEW MODAL
// ============================================================
function _ensureBulkPreviewModal_() {
    var dlg = document.getElementById('appBulkPreviewModal');
    if (dlg) return dlg;
    dlg = document.createElement('dialog');
    dlg.id = 'appBulkPreviewModal';
    dlg.className = 'modal modal-bottom sm:modal-middle';
    dlg.innerHTML =
        '<div class="modal-box bulk-preview-box">' +
        '  <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2" id="bulkPreviewCloseX">✕</button>' +
        '  <h3 id="bulkPreviewTitle" class="font-bold text-lg pr-8">Verify Upload Data</h3>' +
        '  <p id="bulkPreviewSubtitle" class="text-xs text-base-content/55 mt-1 mb-3"></p>' +
        '  <div id="bulkPreviewStats" class="bulk-preview-stats"></div>' +
        '  <div class="bulk-preview-table-wrap">' +
        '    <table id="bulkPreviewTable" class="table table-xs w-full">' +
        '      <thead id="bulkPreviewThead"></thead>' +
        '      <tbody id="bulkPreviewTbody"></tbody>' +
        '    </table>' +
        '  </div>' +
        '  <div class="modal-action pt-3 border-t border-base-200">' +
        '    <button type="button" id="bulkPreviewCancel" class="btn btn-ghost btn-sm">Cancel</button>' +
        '    <button type="button" id="bulkPreviewConfirm" class="btn btn-primary btn-sm gap-1.5">' +
        '      <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 0115.9 6L16 6a5 5 0 011 9.9M15 13l-3 3m0 0l-3-3m3 3V8"/></svg>' +
        '      Confirm &amp; Upload' +
        '    </button>' +
        '  </div>' +
        '</div>' +
        '<form method="dialog" class="modal-backdrop"><button>close</button></form>';
    document.body.appendChild(dlg);
    return dlg;
}

/**
 * showBulkPreview_(config)
 * config = {
 *   title, subtitle,
 *   headers: [{key, label, required}],
 *   rows: allMapped,
 *   validRows: filteredValid,
 *   onConfirm(validRows), onCancel()
 * }
 */
function showBulkPreview_(config) {
    var dlg = _ensureBulkPreviewModal_();
    var titleEl  = document.getElementById('bulkPreviewTitle');
    var subEl    = document.getElementById('bulkPreviewSubtitle');
    var statsEl  = document.getElementById('bulkPreviewStats');
    var thead    = document.getElementById('bulkPreviewThead');
    var tbody    = document.getElementById('bulkPreviewTbody');
    var btnOk    = document.getElementById('bulkPreviewConfirm');
    var btnX     = document.getElementById('bulkPreviewCloseX');
    var btnCancel= document.getElementById('bulkPreviewCancel');

    var allRows   = config.rows || [];
    var validRows = config.validRows || allRows;
    var skipped   = allRows.length - validRows.length;
    var validSet  = new Set(validRows);

    titleEl.textContent = config.title || 'Verify Upload Data';
    subEl.textContent   = config.subtitle || '';

    // Stats bar
    statsEl.innerHTML =
        '<div class="bulk-stat bulk-stat-total"><span class="bulk-stat-num">' + allRows.length + '</span><span class="bulk-stat-lbl">Total Rows</span></div>' +
        '<div class="bulk-stat bulk-stat-valid"><span class="bulk-stat-num">' + validRows.length + '</span><span class="bulk-stat-lbl">Valid</span></div>' +
        (skipped > 0 ? '<div class="bulk-stat bulk-stat-skip"><span class="bulk-stat-num">' + skipped + '</span><span class="bulk-stat-lbl">Skipped</span></div>' : '');

    // Table header
    thead.innerHTML = '<tr class="bulk-preview-thead-row">' +
        '<th class="bulk-col-no">#</th>' +
        (config.headers || []).map(function(h) {
            return '<th>' + escHtml(h.label) + (h.required ? '<span class="bulk-req-star">*</span>' : '') + '</th>';
        }).join('') +
        '<th class="bulk-col-status">Status</th>' +
        '</tr>';

    // Table body
    tbody.innerHTML = allRows.map(function(row, i) {
        var isValid = validSet.has(row);
        var cellsHtml = (config.headers || []).map(function(h) {
            var val = row[h.key];
            if (val === undefined || val === null) val = '';
            var sval = String(val).trim();
            var missing = h.required && !sval;
            return '<td class="' + (missing ? 'bulk-cell-missing' : '') + '" data-label="' + escHtml(h.label) + '">' +
                (sval ? escHtml(sval) : '<span class="bulk-empty-cell">—</span>') + '</td>';
        }).join('');
        var badge = isValid
            ? '<span class="bulk-badge-valid">✓ Valid</span>'
            : '<span class="bulk-badge-skip">✗ Skipped</span>';
        return '<tr class="' + (isValid ? '' : 'bulk-row-invalid') + '">' +
            '<td class="bulk-col-no" data-label="#">' + (i + 1) + '</td>' +
            cellsHtml +
            '<td class="bulk-col-status" data-label="Status">' + badge + '</td>' +
            '</tr>';
    }).join('');

    // Confirm button disability
    btnOk.disabled = validRows.length === 0;
    btnOk.classList.toggle('btn-disabled', validRows.length === 0);

    function _close_() {
        btnOk.onclick = null;
        btnCancel.onclick = null;
        btnX.onclick = null;
        dlg.onclose = null;
    }
    function _cancel_() { _close_(); dlg.close(); if (config.onCancel) config.onCancel(); }

    btnCancel.onclick = _cancel_;
    btnX.onclick = _cancel_;
    dlg.onclose = function() { _close_(); };

    btnOk.onclick = function() {
        if (validRows.length === 0) return;
        _close_();
        dlg.close();
        if (config.onConfirm) config.onConfirm(validRows);
    };

    dlg.showModal();
}

// ============================================================
// VENDOR MASTER — BULK UPLOAD
// ============================================================
function triggerVendorBulkUpload() {
    if (!canBulkUploadMasters()) return;
    var inp = $('vendorBulkFile');
    if (inp) inp.click();
}

function handleVendorBulkFile(input) {
    if (!input || !input.files || !input.files[0]) return;
    var file = input.files[0];
    readSpreadsheetRows(file, function(rows) {
        if (!rows || !rows.length) {
            showToast('No rows found in file.', 'warning');
            input.value = '';
            return;
        }

        var allMapped = rows.map(function(r) {
            var name = r.name || r.vendor_name || r.vendor || r['Vendor Name'] || r['Name'] || '';
            var address = r.address || r.vendor_address || r['Vendor Address'] || '';
            var gstin = r.gstin || r.gst || r.vendor_gstin || r['GSTIN'] || '';
            return {
                name: String(name || '').trim(),
                address: String(address || '').trim(),
                gstin: String(gstin || '').trim()
            };
        });
        var validRows = allMapped.filter(function(r) { return r.name; });

        showBulkPreview_({
            title: 'Verify Vendor Upload',
            subtitle: 'File: ' + file.name + '  •  ' + rows.length + ' row(s) read — review below before uploading.',
            headers: [
                { key: 'name',    label: 'Vendor Name', required: true  },
                { key: 'address', label: 'Address',     required: false },
                { key: 'gstin',   label: 'GSTIN',       required: false }
            ],
            rows: allMapped,
            validRows: validRows,
            onCancel: function() { input.value = ''; },
            onConfirm: function(confirmed) {
                gsPost({ action: 'bulk_upsert_vendors', rows: confirmed })
                .then(function(resp) {
                    if (resp.success) {
                        invalidateCache('read_vendors');
                        invalidateCache('read_all');
                        fetchVendors();
                        showToast('Vendor bulk upload complete. Added: ' + (resp.added || 0) + ', Updated: ' + (resp.updated || 0), 'success');
                    } else {
                        showToast('Vendor bulk upload failed: ' + (resp.error || 'Unknown error'), 'error');
                    }
                    input.value = '';
                })
                .catch(function(err) {
                    showToast('Vendor bulk upload failed: ' + err.message, 'error');
                    input.value = '';
                });
            }
        });
    });
}

// ============================================================
// ITEM MASTER
// ============================================================
let itemMasterCache = [];

function fetchItemMaster() {
    gsFetchCached('read_items')
        .then(data => {
            itemMasterCache = data || [];
            if ($('itemMasterList')) renderItemMasterList(data);
            if ($('poItemsTable')) renderPOItemsTable();
        })
        .catch(() => {
            if ($('itemMasterList')) $('itemMasterList').innerHTML = '<div class="text-error p-4">Failed to load items</div>';
        });
}

var _itemMasterData = [];
var _itemSearchTerm = '';
function filterItemList(term) {
    _itemSearchTerm = (term || '').toLowerCase().trim();
    _pagination['item'] = { page: 1, size: (_pagination['item'] || {}).size || 10 };
    renderItemMasterList();
}
function renderItemMasterList(data) {
    if (data) _itemMasterData = data;
    var allData = _itemMasterData || [];
    // Apply search filter
    var filtered = allData;
    if (_itemSearchTerm) {
        filtered = allData.filter(function(it) {
            return (it.desc || '').toLowerCase().includes(_itemSearchTerm)
                || (it.uom || '').toLowerCase().includes(_itemSearchTerm)
                || String(it.rate || '').includes(_itemSearchTerm);
        });
    }
    if (!Array.isArray(allData) || allData.length === 0) {
        $('itemMasterList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4" /></svg><p class="text-sm">No items found. Add your first item above.</p></div>';
        return;
    }
    if (filtered.length === 0 && _itemSearchTerm) {
        $('itemMasterList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg><p class="text-sm">No items match "<strong>' + escHtml(_itemSearchTerm) + '</strong>"</p></div>';
        return;
    }
    // Hero stats
    if ($('itemHeroStats')) {
        var statsHtml = '<div class="hero-stat"><div class="stat-label">Total Items</div><div class="stat-value">' + allData.length + '</div></div>';
        if (_itemSearchTerm) statsHtml += '<div class="hero-stat"><div class="stat-label">Matched</div><div class="stat-value">' + filtered.length + '</div></div>';
        $('itemHeroStats').innerHTML = statsHtml;
    }
    var pg = _initPagination('item', 10);
    var total = filtered.length;
    var startIdx = (pg.page - 1) * pg.size;
    var pageData = _pgSlice('item', filtered);

    var canEditMaster = canManageMasters();
    let html = '<table class="table table-zebra w-full mobile-card-table"><thead><tr><th>#</th><th>Description</th><th>UOM</th><th>Rate/Day</th><th>Default Qty</th>' + (canEditMaster ? '<th>Action</th>' : '') + '</tr></thead><tbody>';
    pageData.forEach((item, i) => {
        var displayIdx = startIdx + i;
        var origIdx = allData.indexOf(item);
        html += '<tr><td data-label="#" class="font-medium text-slate-500">' + (displayIdx + 1) + '</td><td data-label="Description" class="font-semibold text-slate-800">' + item.desc + '</td><td data-label="UOM"><span class="badge badge-ghost badge-sm">' + item.uom + '</span></td><td data-label="Rate/Day" class="font-mono text-slate-700">' + item.rate + '</td><td data-label="Default Qty" class="text-slate-600">' + (item.default_qty || '') + '</td>' + (canEditMaster ? '<td data-label="Actions" class="flex gap-1"><button class="btn btn-xs btn-warning btn-outline btn-modern" onclick="editItemMaster(' + origIdx + ')">Edit</button><button class="btn btn-xs btn-error btn-outline btn-modern" onclick="deleteItemMaster(' + origIdx + ')">Delete</button></td>' : '') + '</tr>';
    });
    html += '</tbody></table>';
    if (total > pg.size) html += _buildPaginationBar('item', total, renderItemMasterList);
    $('itemMasterList').innerHTML = html;
}

function addItemMaster() {
    if (!canManageMasters()) return;
    const desc = $('itemDescInput').value.trim();
    const uom = $('itemUomInput').value.trim();
    const rate = $('itemRateInput').value.trim();
    const default_qty = $('itemDefaultQtyInput').value.trim();
    if (!desc || !uom || !rate) return;
    gsPost({ action: 'add_item', desc, uom, rate, default_qty })
    .then(resp => {
        if (resp.success) {
            invalidateCache('read_items'); invalidateCache('read_all');
            fetchItemMaster();
            $('itemDescInput').value = '';
            $('itemUomInput').value = '';
            $('itemRateInput').value = '';
            $('itemDefaultQtyInput').value = '';
        } else alert('Failed to add item: ' + (resp.error || 'Unknown error'));
    })
    .catch(err => alert('Error adding item: ' + err.message));
}

function editItemMaster(idx) {
    if (!canManageMasters()) return;
    var item = (_itemMasterData || [])[idx];
    if (!item) return;

    showEditPopup_({
        title: 'Edit Item',
        fields: [
            { name: 'desc', label: 'Description', value: item.desc || '', required: true },
            { name: 'uom', label: 'UOM', value: item.uom || 'NOS', required: true },
            { name: 'rate', label: 'Rate / Day', type: 'number', value: item.rate || 0, required: true },
            { name: 'default_qty', label: 'Default Qty', type: 'number', value: item.default_qty || 1, required: true }
        ]
    }).then(function(data) {
        if (!data) return;
        gsPost({ action: 'update_item', idx: idx, desc: data.desc, uom: data.uom, rate: data.rate, default_qty: data.default_qty })
        .then(function(resp) {
            if (resp.success) {
                invalidateCache('read_items');
                invalidateCache('read_all');
                fetchItemMaster();
                showToast('Item updated.', 'success');
            } else {
                showToast('Failed to update item: ' + (resp.error || 'Unknown error'), 'error');
            }
        })
        .catch(function(err) {
            showToast('Error updating item: ' + err.message, 'error');
        });
    });
}

function deleteItemMaster(idx) {
    if (!canManageMasters()) return;
    if (!confirm('Delete this item?')) return;
    gsPost({ action: 'delete_item', idx })
    .then(resp => { if (resp.success) { invalidateCache('read_items'); invalidateCache('read_all'); fetchItemMaster(); } else alert('Failed to delete item: ' + (resp.error || 'Unknown error')); })
    .catch(err => alert('Error deleting item: ' + err.message));
}

function triggerItemBulkUpload() {
    if (!canBulkUploadMasters()) return;
    var inp = $('itemBulkFile');
    if (inp) inp.click();
}

function handleItemBulkFile(input) {
    if (!input || !input.files || !input.files[0]) return;
    var file = input.files[0];
    readSpreadsheetRows(file, function(rows) {
        if (!rows || !rows.length) {
            showToast('No rows found in file.', 'warning');
            input.value = '';
            return;
        }

        var allMapped = rows.map(function(r) {
            var desc = r.desc || r.description || r.item || r['Item Description'] || r['Description'] || '';
            var uom = r.uom || r.unit || r['UOM'] || 'NOS';
            var rate = r.rate || r.per_day || r['Rate/Day'] || 0;
            var defaultQty = r.default_qty || r.defaultqty || r.qty || r['Default Qty'] || 1;
            return {
                desc: String(desc || '').trim(),
                uom: String(uom || 'NOS').trim() || 'NOS',
                rate: Number(rate) || 0,
                default_qty: parseInt(defaultQty, 10) || 1
            };
        });
        var validRows = allMapped.filter(function(r) { return r.desc; });

        showBulkPreview_({
            title: 'Verify Item Upload',
            subtitle: 'File: ' + file.name + '  •  ' + rows.length + ' row(s) read — review below before uploading.',
            headers: [
                { key: 'desc',        label: 'Description', required: true  },
                { key: 'uom',         label: 'UOM',         required: false },
                { key: 'rate',        label: 'Rate/Day',    required: false },
                { key: 'default_qty', label: 'Default Qty', required: false }
            ],
            rows: allMapped,
            validRows: validRows,
            onCancel: function() { input.value = ''; },
            onConfirm: function(confirmed) {
                gsPost({ action: 'bulk_upsert_items', rows: confirmed })
                .then(function(resp) {
                    if (resp.success) {
                        invalidateCache('read_items');
                        invalidateCache('read_all');
                        fetchItemMaster();
                        showToast('Item bulk upload complete. Added: ' + (resp.added || 0) + ', Updated: ' + (resp.updated || 0), 'success');
                    } else {
                        showToast('Item bulk upload failed: ' + (resp.error || 'Unknown error'), 'error');
                    }
                    input.value = '';
                })
                .catch(function(err) {
                    showToast('Item bulk upload failed: ' + err.message, 'error');
                    input.value = '';
                });
            }
        });
    });
}

// ============================================================
// PO LIST
// ============================================================
function fetchPOs() {
    gsFetchCached('read')
        .then(data => {
            window._poList = data;
            if ($('poList')) renderPOList(data);
        })
        .catch(() => {
            if ($('poList')) $('poList').innerHTML = '<div class="text-error p-4">Failed to load POs</div>';
        });
}

var _poSearchTerm = '';
var _poVendorFilter = '';
var _poDateFrom = '';
var _poDateTo = '';
var _poFilteredList = [];
var _poQuickRange = '';
var _poMinAmount = null;
var _poMaxAmount = null;
var _poHasItemsFilter = '';
var _poViewMode = 'table';
var _poSearchDebounceTimer = null;
var _poDesktopPreferredView = 'table';
var _poResponsiveViewBound = false;
var _poVisibleColumns = {
    po_no: true,
    date: true,
    vendor: true,
    event: true,
    location: true,
    amount: true,
    actions: true
};

function _toDateOnly_(value) {
    if (!value) return null;
    var d = new Date(value);
    if (!isNaN(d.getTime())) {
        d.setHours(0, 0, 0, 0);
        return d;
    }
    var parts = String(value).split(/[\/-]/);
    if (parts.length === 3) {
        var dd = parseInt(parts[0], 10);
        var mm = parseInt(parts[1], 10);
        var yy = parseInt(parts[2], 10);
        var d2 = new Date(yy, mm - 1, dd);
        if (!isNaN(d2.getTime())) {
            d2.setHours(0, 0, 0, 0);
            return d2;
        }
    }
    return null;
}

function _getPOTotal_(po) {
    var t = parseFloat(po.total);
    if (!isNaN(t) && t > 0) return t;
    var items = po.items || [];
    if (typeof items === 'string') {
        try { items = JSON.parse(items); } catch (e) { items = []; }
    }
    if (!Array.isArray(items)) items = [];
    var itemsTotal = items.reduce(function(sum, it) { return sum + (parseFloat(it.total) || 0); }, 0);
    var transport = parseFloat(po.transport) || 0;
    var sub = itemsTotal + transport;
    var cgst = sub * (parseFloat(po.cgst_percent) || 0) / 100;
    var sgst = sub * (parseFloat(po.sgst_percent) || 0) / 100;
    return Math.round(sub + cgst + sgst);
}

function _applyPOFilters_(allData) {
    var fromDate = _poDateFrom ? _toDateOnly_(_poDateFrom) : null;
    var toDate = _poDateTo ? _toDateOnly_(_poDateTo) : null;
    return allData.filter(function(po) {
        var matchesSearch = !_poSearchTerm || (
            (po.po_no || '').toLowerCase().includes(_poSearchTerm)
            || (po.vendor_name || '').toLowerCase().includes(_poSearchTerm)
            || (po.event_name || '').toLowerCase().includes(_poSearchTerm)
            || (po.event_location || '').toLowerCase().includes(_poSearchTerm)
            || (po.po_date || '').toLowerCase().includes(_poSearchTerm)
            || String(po.total || '').includes(_poSearchTerm)
        );
        if (!matchesSearch) return false;

        if (_poVendorFilter) {
            var vn = String(po.vendor_name || '').toLowerCase().trim();
            if (vn !== _poVendorFilter) return false;
        }

        if (fromDate || toDate) {
            var pd = _toDateOnly_(po.po_date);
            if (!pd) return false;
            if (fromDate && pd < fromDate) return false;
            if (toDate && pd > toDate) return false;
        }

        var poTotal = _getPOTotal_(po);
        if (_poMinAmount !== null && poTotal < _poMinAmount) return false;
        if (_poMaxAmount !== null && poTotal > _poMaxAmount) return false;

        if (_poHasItemsFilter) {
            var items = po.items || [];
            if (typeof items === 'string') {
                try { items = JSON.parse(items); } catch (e) { items = []; }
            }
            var hasItems = Array.isArray(items) && items.length > 0;
            if (_poHasItemsFilter === 'yes' && !hasItems) return false;
            if (_poHasItemsFilter === 'no' && hasItems) return false;
        }

        return true;
    });
}

function _extractPOSequence_(poNo) {
    var text = String(poNo || '').trim();
    var m = text.match(/\/(\d+)$/);
    if (!m) return 0;
    var n = parseInt(m[1], 10);
    return isNaN(n) ? 0 : n;
}

// Keep PO list newest-first: latest date first, then higher PO sequence.
function _sortPOListLatestFirst_(list) {
    return (Array.isArray(list) ? list.slice() : []).sort(function(a, b) {
        var da = _toDateOnly_(a && a.po_date);
        var db = _toDateOnly_(b && b.po_date);
        var ta = da ? da.getTime() : 0;
        var tb = db ? db.getTime() : 0;
        if (tb !== ta) return tb - ta;

        var sa = _extractPOSequence_(a && a.po_no);
        var sb = _extractPOSequence_(b && b.po_no);
        if (sb !== sa) return sb - sa;

        return String((b && b.po_no) || '').localeCompare(String((a && a.po_no) || ''));
    });
}

function _populatePOVendorFilter_(allData) {
    var sel = $('poVendorFilter');
    if (!sel) return;
    var current = _poVendorFilter;
    var seen = {};
    var vendors = [];
    allData.forEach(function(po) {
        var name = String(po.vendor_name || '').trim();
        var key = name.toLowerCase();
        if (!name || seen[key]) return;
        seen[key] = true;
        vendors.push({ key: key, label: name });
    });
    vendors.sort(function(a, b) { return a.label.localeCompare(b.label); });

    var html = '<option value="">All Vendors</option>';
    vendors.forEach(function(v) {
        html += '<option value="' + escHtml(v.key) + '">' + escHtml(v.label) + '</option>';
    });
    sel.innerHTML = html;
    sel.value = current;
}

function _updatePOFilteredSummary_(filtered) {
    if ($('poFilteredCount')) {
        var totalAll = (window._poList || []).length;
        $('poFilteredCount').textContent = filtered.length + ' of ' + totalAll;
    }
    if ($('poFilteredValue')) {
        var total = filtered.reduce(function(sum, po) { return sum + _getPOTotal_(po); }, 0);
        $('poFilteredValue').textContent = 'Rs ' + total.toLocaleString('en-IN');
    }
}

function handleGlobalSearch(query) {
    if (_poSearchDebounceTimer) clearTimeout(_poSearchDebounceTimer);
    _poSearchDebounceTimer = setTimeout(function() {
        filterPOList(query);
    }, 180);
}

function toggleAdvancedFilters() {
    var panel = $('advancedFiltersContent');
    var icon = $('advancedFilterIcon');
    if (!panel) return;
    panel.classList.toggle('hidden');
    if (icon) {
        icon.style.transform = panel.classList.contains('hidden') ? 'rotate(0deg)' : 'rotate(180deg)';
    }
}

function _updateAdvancedFilterCount_() {
    var count = 0;
    if (_poMinAmount !== null) count += 1;
    if (_poMaxAmount !== null) count += 1;
    if (_poHasItemsFilter) count += 1;
    if ($('advancedFilterCount')) $('advancedFilterCount').textContent = count + ' active';
}

function applyAdvancedFilters() {
    var minV = $('filterMinAmount') ? $('filterMinAmount').value.trim() : '';
    var maxV = $('filterMaxAmount') ? $('filterMaxAmount').value.trim() : '';
    _poMinAmount = minV === '' ? null : Math.max(0, Number(minV));
    _poMaxAmount = maxV === '' ? null : Math.max(0, Number(maxV));
    _poHasItemsFilter = $('filterHasItems') ? String($('filterHasItems').value || '') : '';
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    _updateAdvancedFilterCount_();
    renderPOList();
}

function clearAdvancedFilters() {
    _poMinAmount = null;
    _poMaxAmount = null;
    _poHasItemsFilter = '';
    if ($('filterMinAmount')) $('filterMinAmount').value = '';
    if ($('filterMaxAmount')) $('filterMaxAmount').value = '';
    if ($('filterHasItems')) $('filterHasItems').value = '';
    _updateAdvancedFilterCount_();
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function _isPOMobileViewport_() {
    return window.matchMedia('(max-width: 767px)').matches;
}

function setTableView(view, opts) {
    opts = opts || {};
    _poViewMode = (view === 'cards') ? 'cards' : 'table';

    // Keep desktop user preference so we can restore it after mobile auto-mode.
    if (!_isPOMobileViewport_() && !opts.auto) {
        _poDesktopPreferredView = _poViewMode;
    }

    var tableBtn = $('poTableViewBtn');
    var cardBtn = $('poCardViewBtn');
    if (tableBtn) {
        tableBtn.classList.toggle('active', _poViewMode === 'table');
        tableBtn.classList.toggle('btn-outline', _poViewMode !== 'table');
    }
    if (cardBtn) {
        cardBtn.classList.toggle('active', _poViewMode === 'cards');
        cardBtn.classList.toggle('btn-outline', _poViewMode !== 'cards');
    }
    var tableWrap = $('poListTableWrap');
    var cardWrap = $('poCardView');
    if (tableWrap) tableWrap.classList.toggle('hidden', _poViewMode === 'cards');
    if (cardWrap) cardWrap.classList.toggle('hidden', _poViewMode !== 'cards');
}

function _applyAutoPOViewMode_() {
    // Only run on PO list screen where toggle buttons exist.
    if (!$('poTableViewBtn') || !$('poCardViewBtn')) return;
    if (_isPOMobileViewport_()) {
        setTableView('cards', { auto: true });
    } else {
        setTableView(_poDesktopPreferredView || 'table', { auto: true });
    }
}

function _bindPOResponsiveViewMode_() {
    if (_poResponsiveViewBound) return;
    if (!$('poTableViewBtn') || !$('poCardViewBtn')) return;
    _poResponsiveViewBound = true;

    var timer = null;
    window.addEventListener('resize', function() {
        if (timer) clearTimeout(timer);
        timer = setTimeout(function() {
            _applyAutoPOViewMode_();
        }, 120);
    });
}

function togglePOColumn(columnName, isVisible) {
    if (!_poVisibleColumns.hasOwnProperty(columnName)) return;
    _poVisibleColumns[columnName] = !!isVisible;
    applyPOColumnVisibility();
}

function applyPOColumnVisibility() {
    var table = $('poTable');
    Object.keys(_poVisibleColumns).forEach(function(col) {
        var visible = _poVisibleColumns[col];
        if (table) {
            table.querySelectorAll('[data-col="' + col + '"]').forEach(function(cell) {
                cell.style.display = visible ? '' : 'none';
            });
        }
        var cardWrap = $('poCardView');
        if (cardWrap) {
            cardWrap.querySelectorAll('[data-col="' + col + '"]').forEach(function(cell) {
                cell.style.display = visible ? '' : 'none';
            });
        }
    });
}

function bindPOColumnControls() {
    document.querySelectorAll('[data-po-column]').forEach(function(chk) {
        var col = chk.getAttribute('data-po-column');
        chk.checked = _poVisibleColumns[col] !== false;
        chk.onchange = function() {
            togglePOColumn(col, chk.checked);
        };
    });
}

function _renderPOCardGrid_(pageData, allData, startIdx) {
    var html = '<div class="po-card-grid">';
    pageData.forEach(function(po, i) {
        var idx = allData.indexOf(po);
        var poId = po.po_id || po.po_no;
        var isSelected = window._poSelectedIds && window._poSelectedIds.has(poId);
        var totalNum = _getPOTotal_(po);
        var poNo = escHtml(String(po.po_no || '-'));
        var vendor = escHtml(String(po.vendor_name || '-'));
        var date = escHtml(formatDateDDMMMYYYY(po.po_date) || '-');
        var eventName = escHtml(String(po.event_name || '-'));
        var location = escHtml(String(po.event_location || '-'));
        html += ''
            + '<div class="po-mobile-card ' + (isSelected ? 'border-primary bg-primary/5' : '') + '">'
            + '  <div class="flex items-start justify-between gap-2">'
            + '    <div class="flex items-start gap-2">'
            + '      <div class="mt-1"><input type="checkbox" class="checkbox checkbox-xs" data-po-id="' + poId + '" ' + (isSelected ? 'checked' : '') + ' onchange="onPOSelectChange(\'' + poId + '\', this.checked); renderPOList();"></div>'
            + '      <div>'
            + '        <div class="po-card-title" data-col="po_no">' + poNo + '</div>'
            + '        <div class="po-card-meta" data-col="date">#' + (startIdx + i + 1) + ' • ' + date + '</div>'
            + '      </div>'
            + '    </div>'
            + '    <div data-col="amount" class="text-right font-mono text-emerald-600 text-sm font-semibold">₹' + Number(totalNum || 0).toLocaleString('en-IN') + '</div>'
            + '  </div>'
            + '  <div data-col="vendor" class="mt-2 text-sm text-slate-700"><span class="font-medium">Vendor:</span> ' + vendor + '</div>'
            + '  <div data-col="event" class="text-xs text-slate-500 mt-0.5">' + eventName + '</div>'
            + '  <div data-col="location" class="text-xs text-slate-500 mt-0.5">' + location + '</div>'
            + '  <div data-col="actions" class="mt-3 flex flex-wrap gap-1.5">'
            + '    <a href="po-view.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-primary btn-outline">View</a>'
            + (canManagePO() ? ('<a href="po-edit.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-warning btn-outline">Edit</a>') : '')
            + '    <button type="button" class="btn btn-xs btn-error btn-outline" onclick="listDownloadPDF(' + idx + ')">PDF</button>'
            + '    <button type="button" class="btn btn-xs btn-success btn-outline" onclick="listDownloadExcel(' + idx + ')">Excel</button>'
            + '    <button type="button" class="btn btn-xs btn-neutral btn-outline" onclick="showPOItemsModal(' + idx + ')">Items</button>'
            + (isAdmin() ? ('<button type="button" class="btn btn-xs btn-error" onclick="confirmDeletePO(\'' + poId + '\')"><svg xmlns="http://www.w3.org/2000/svg" class="h-3 w-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg></button>') : '')
            + '  </div>'
            + '</div>';
    });
    html += '</div>';
    return html;
}

function filterPOList(term) {
    _poSearchTerm = (term || '').toLowerCase().trim();
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function setPOVendorFilter(vendorKey) {
    _poVendorFilter = String(vendorKey || '').toLowerCase().trim();
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function setPODateFilters() {
    _poDateFrom = ($('poDateFrom') ? $('poDateFrom').value : '') || '';
    _poDateTo = ($('poDateTo') ? $('poDateTo').value : '') || '';
    _poQuickRange = '';
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function _fmtDateInput_(d) {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + day;
}

function _applyQuickRangeButtonState_() {
    var ids = {
        'today': 'poQuickToday',
        'week': 'poQuickWeek',
        'month': 'poQuickMonth',
        'last-month': 'poQuickLastMonth'
    };
    Object.keys(ids).forEach(function(k) {
        var el = $(ids[k]);
        if (!el) return;
        if (_poQuickRange === k) {
            el.classList.remove('btn-outline');
            el.classList.add('btn-primary');
        } else {
            el.classList.remove('btn-primary');
            el.classList.add('btn-outline');
        }
    });
}

function applyPOQuickDateRange(type) {
    var now = new Date();
    now.setHours(0, 0, 0, 0);
    var from = '';
    var to = '';

    if (type === 'today') {
        from = _fmtDateInput_(now);
        to = _fmtDateInput_(now);
    } else if (type === 'week') {
        var dayIdx = now.getDay();
        var mondayOffset = (dayIdx === 0 ? -6 : 1 - dayIdx);
        var monday = new Date(now);
        monday.setDate(now.getDate() + mondayOffset);
        var sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);
        from = _fmtDateInput_(monday);
        to = _fmtDateInput_(sunday);
    } else if (type === 'month') {
        var m1 = new Date(now.getFullYear(), now.getMonth(), 1);
        var m2 = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        from = _fmtDateInput_(m1);
        to = _fmtDateInput_(m2);
    } else if (type === 'last-month') {
        var lm1 = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        var lm2 = new Date(now.getFullYear(), now.getMonth(), 0);
        from = _fmtDateInput_(lm1);
        to = _fmtDateInput_(lm2);
    } else {
        type = '';
    }

    _poQuickRange = type;
    _poDateFrom = from;
    _poDateTo = to;
    if ($('poDateFrom')) $('poDateFrom').value = from;
    if ($('poDateTo')) $('poDateTo').value = to;

    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function clearPOFilters() {
    _poSearchTerm = '';
    _poVendorFilter = '';
    _poDateFrom = '';
    _poDateTo = '';
    _poQuickRange = '';
    _poMinAmount = null;
    _poMaxAmount = null;
    _poHasItemsFilter = '';

    if ($('poSearchInput')) $('poSearchInput').value = '';
    if ($('poVendorFilter')) $('poVendorFilter').value = '';
    if ($('poDateFrom')) $('poDateFrom').value = '';
    if ($('poDateTo')) $('poDateTo').value = '';
    if ($('filterMinAmount')) $('filterMinAmount').value = '';
    if ($('filterMaxAmount')) $('filterMaxAmount').value = '';
    if ($('filterHasItems')) $('filterHasItems').value = '';
    _updateAdvancedFilterCount_();

    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}

function exportCurrentPOViewExcel() {
    var list = Array.isArray(_poFilteredList) && _poFilteredList.length ? _poFilteredList : [];
    if (!list.length) {
        showToast('No filtered POs available to export.', 'warning');
        return;
    }
    exportAllPOsExcel(list);
}

function renderPOList(data) {
    if (data) window._poList = data;
    var allData = window._poList || [];
    _populatePOVendorFilter_(allData);
    var filtered = _sortPOListLatestFirst_(_applyPOFilters_(allData));
    _poFilteredList = filtered.slice();
    _updatePOFilteredSummary_(filtered);
    _applyQuickRangeButtonState_();
    _updateAdvancedFilterCount_();
    bindPOColumnControls();
    if (!Array.isArray(allData) || allData.length === 0) {
        showEmptyState('poList', 'No purchase orders yet', 'Create your first purchase order to start tracking procurement.', 'Create PO', "window.location.href='add-po.html'");
        return;
    }
    if (filtered.length === 0) {
        var hasFilter = !!(_poSearchTerm || _poVendorFilter || _poDateFrom || _poDateTo || _poMinAmount !== null || _poMaxAmount !== null || _poHasItemsFilter);
        var emptyMsg = hasFilter ? 'No POs match current filters. Try broadening your criteria.' : 'No purchase orders found.';
        showEmptyState('poList', 'No matching records', emptyMsg, 'Reset Filters', 'clearPOFilters()');
        return;
    }
    // Hero stats
    if ($('poHeroStats')) {
        var statsHtml = '<div class="hero-stat"><div class="stat-label">Total POs</div><div class="stat-value">' + allData.length + '</div></div>';
        if (_poSearchTerm || _poVendorFilter || _poDateFrom || _poDateTo) {
            statsHtml += '<div class="hero-stat"><div class="stat-label">Filtered</div><div class="stat-value">' + filtered.length + '</div></div>';
        }
        var filteredValue = filtered.reduce(function(sum, po) { return sum + _getPOTotal_(po); }, 0);
        statsHtml += '<div class="hero-stat"><div class="stat-label">Value</div><div class="stat-value">₹' + filteredValue.toLocaleString('en-IN') + '</div></div>';
        $('poHeroStats').innerHTML = statsHtml;
    }
    var pg = _initPagination('po', 10);
    var total = filtered.length;
    var startIdx = (pg.page - 1) * pg.size;
    var pageData = _pgSlice('po', filtered);

    // Track selected PO IDs for bulk actions
    if (!window._poSelectedIds) window._poSelectedIds = new Set();
    
    // Check if the select-all checkbox should be checked
    var areAllPageItemsSelected = pageData.length > 0 && pageData.every(po => window._poSelectedIds.has(po.po_id || po.po_no));

    function safeCell(v, maxLen) {
        var s = (v == null ? '' : String(v));
        s = s.replace(/\s+/g, ' ').trim();
        if (maxLen && s.length > maxLen) s = s.slice(0, maxLen - 1) + '...';
        return escHtml(s);
    }

    let html = '<div id="poListTableWrap" class="overflow-x-auto"><table id="poTable" class="table table-zebra w-full mobile-card-table"><thead><tr><th class="w-10"><label class="cursor-pointer"><input type="checkbox" class="checkbox checkbox-sm checkbox-primary" ' + (areAllPageItemsSelected ? 'checked' : '') + ' onchange="toggleAllPagePOs(this.checked)"></label></th><th>#</th><th data-col="po_no">PO No</th><th data-col="date">Date</th><th data-col="vendor">Vendor</th><th data-col="event">Event</th><th data-col="location">Location</th><th data-col="amount">Total</th><th data-col="actions">Actions</th></tr></thead><tbody id="poTableBody">';
    pageData.forEach((po, i) => {
        var displayIdx = startIdx + i;
        var idx = allData.indexOf(po);
        var poId = po.po_id || po.po_no;
        var isSelected = window._poSelectedIds.has(poId);
        var poNo = safeCell(po.po_no, 40);
        var poDate = safeCell(formatDateDDMMMYYYY(po.po_date), 24);
        var vendor = safeCell(po.vendor_name, 64);
        var eventName = safeCell(po.event_name, 56);
        var eventLocation = safeCell(po.event_location, 72);
        var totalNum = _getPOTotal_(po);
        html += '<tr>';
        html += '<td class="w-10"><label class="cursor-pointer"><input type="checkbox" class="checkbox checkbox-sm" data-po-id="' + poId + '" ' + (isSelected ? 'checked' : '') + ' onchange="onPOSelectChange(\'' + poId + '\', this.checked)"></label></td>';
        html += '<td data-label="#" class="font-medium text-slate-500">' + (displayIdx + 1) + '</td>';
        html += '<td data-label="PO No" data-col="po_no"><span class="font-semibold text-primary" title="' + poNo + '">' + poNo + '</span></td>';
        html += '<td data-label="Date" data-col="date" class="text-slate-600">' + poDate + '</td>';
        html += '<td data-label="Vendor" data-col="vendor" class="font-medium" title="' + vendor + '">' + vendor + '</td>';
        html += '<td data-label="Event" data-col="event" class="text-slate-600" title="' + eventName + '">' + eventName + '</td>';
        html += '<td data-label="Location" data-col="location" class="text-slate-500 text-xs" title="' + eventLocation + '">' + eventLocation + '</td>';
        html += '<td data-label="Total" data-col="amount"><span class="font-mono font-semibold text-emerald-600">' + (totalNum ? '₹' + Number(totalNum).toLocaleString('en-IN') : '') + '</span></td>';
        html += '<td data-label="Actions" data-col="actions">';
        html += '<div class="flex flex-wrap gap-1">';
        html += '<a href="po-view.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-primary btn-outline btn-modern" title="View"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg></a>';
        if (canManagePO()) {
        html += '<a href="po-edit.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-warning btn-outline btn-modern" title="Edit"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg></a>';
        }
        html += '<button onclick="listDownloadPDF(' + idx + ')" class="btn btn-xs btn-error btn-outline btn-modern" title="Download PDF"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg></button>';
        html += '<button onclick="listDownloadExcel(' + idx + ')" class="btn btn-xs btn-success btn-outline btn-modern" title="Download Excel"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"/></svg></button>';
        html += '<button onclick="listPrintPDF(' + idx + ')" class="btn btn-xs btn-info btn-outline btn-modern" title="Print PDF"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg></button>';
        html += '<button onclick="showPOItemsModal(' + idx + ')" class="btn btn-xs btn-neutral btn-outline btn-modern" title="Show Items"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h10"/></svg> Items</button>';
        if (isAdmin()) {
            html += '<button onclick="confirmDeletePO(\'' + poId + '\')" class="btn btn-xs btn-error btn-outline btn-modern" title="Delete PO"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg></button>';
        }
        html += '</div>';
        html += '</td>';
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    html += '<div id="poCardView" class="hidden">' + _renderPOCardGrid_(pageData, allData, startIdx) + '</div>';
    if (total > pg.size) html += _buildPaginationBar('po', total, renderPOList);
    $('poList').innerHTML = html;
    applyPOColumnVisibility();
    _bindPOResponsiveViewMode_();
    _applyAutoPOViewMode_();
}

function triggerPOBulkUpload() {
    if (!canBulkUploadPO()) return;
    var inp = $('poBulkFile');
    if (inp) inp.click();
}

function _parsePOItemsForBulk(row) {
    var itemsRaw = row.items_json || row.items || row.line_items || '';
    if (itemsRaw && typeof itemsRaw === 'string') {
        try {
            var parsed = JSON.parse(itemsRaw);
            if (Array.isArray(parsed)) return parsed;
        } catch (e) {
            // Fall through to single-item parsing.
        }
    }

    var desc = row.item_desc || row.desc || row.description || '';
    if (!String(desc || '').trim()) return [];

    var qty = parseFloat(row.qty || row.quantity || 1) || 1;
    var perDay = parseFloat(row.per_day || row.rate || 0) || 0;
    var days = parseFloat(row.days || 1) || 1;
    return [{
        desc: String(desc).trim(),
        qty: qty,
        uom: String(row.uom || 'NOS').trim() || 'NOS',
        per_day: perDay,
        days: days,
        total: qty * perDay * days,
        isMaster: false
    }];
}

function handlePOBulkFile(input) {
    if (!input || !input.files || !input.files[0]) return;
    var file = input.files[0];
    readSpreadsheetRows(file, function(rows) {
        if (!rows || !rows.length) {
            showToast('No rows found in file.', 'warning');
            input.value = '';
            return;
        }

        var allMapped = rows.map(function(r) {
            var poNo = r.po_no || r.po || r['PO No'] || r['PO Number'] || '';
            var poDate = r.po_date || r.date || r['PO Date'] || '';
            var vendorName = r.vendor_name || r.vendor || r['Vendor Name'] || '';
            var parsedItems = _parsePOItemsForBulk(r);
            return {
                po_no: String(poNo || '').trim(),
                po_date: String(poDate || '').trim(),
                vendor_name: String(vendorName || '').trim(),
                vendor_address: String(r.vendor_address || r.address || r['Vendor Address'] || '').trim(),
                vendor_gstin: String(r.vendor_gstin || r.gstin || r.gst || r['GSTIN'] || '').trim(),
                client_name: String(r.client_name || r.client || r['Client Name'] || '').trim(),
                event_name: String(r.event_name || r.event || r['Event Name'] || '').trim(),
                event_location: String(r.event_location || r.location || r['Event Location'] || '').trim(),
                event_date: String(r.event_date || r['Event Date'] || '').trim(),
                transport: parseFloat(r.transport || 0) || 0,
                cgst_percent: parseFloat(r.cgst_percent || r.cgst || 0) || 0,
                sgst_percent: parseFloat(r.sgst_percent || r.sgst || 0) || 0,
                terms: String(r.terms || r['Terms'] || '').trim(),
                items: parsedItems,
                _items_count: parsedItems.length   // display helper
            };
        });
        var validRows = allMapped.filter(function(r) {
            return r.po_no && r.po_date && r.vendor_name;
        });

        // Build display rows that show items as a simple count string
        var displayRows = allMapped.map(function(r) {
            return Object.assign({}, r, {
                items: r._items_count ? r._items_count + ' item(s)' : '—'
            });
        });
        var displayValid = validRows.map(function(r) {
            return displayRows[allMapped.indexOf(r)];
        });

        showBulkPreview_({
            title: 'Verify PO Upload',
            subtitle: 'File: ' + file.name + '  •  ' + rows.length + ' row(s) read — review below before uploading.',
            headers: [
                { key: 'po_no',         label: 'PO No',       required: true  },
                { key: 'po_date',       label: 'Date',        required: true  },
                { key: 'vendor_name',   label: 'Vendor',      required: true  },
                { key: 'event_name',    label: 'Event',       required: false },
                { key: 'event_location',label: 'Location',    required: false },
                { key: 'transport',     label: 'Transport',   required: false },
                { key: 'items',         label: 'Items',       required: false }
            ],
            rows: displayRows,
            validRows: displayValid,
            onCancel: function() { input.value = ''; },
            onConfirm: function(/* displayValid is passed, but we need original rows */) {
                // Use the original validRows (with actual items arrays, not display strings)
                gsPost({ action: 'bulk_upsert_pos', rows: validRows })
                .then(function(resp) {
                    if (resp.success) {
                        invalidateCache('read');
                        invalidateCache('read_all');
                        fetchPOs();
                        showToast('PO bulk upload complete. Added: ' + (resp.added || 0) + ', Updated: ' + (resp.updated || 0), 'success');
                    } else {
                        showToast('PO bulk upload failed: ' + (resp.error || 'Unknown error'), 'error');
                    }
                    input.value = '';
                })
                .catch(function(err) {
                    showToast('PO bulk upload failed: ' + err.message, 'error');
                    input.value = '';
                });
            }
        });
    });
}

// ── Show all items of a PO in a popup modal ──
function showPOItemsModal(idx) {
    var po = _preparePO(idx);
    if (!po) return;
    var items = po._items || [];
    var transport = parseFloat(po.transport) || 0;
    var itemsTotal = items.reduce(function(s, i) { return s + (parseFloat(i.total) || 0); }, 0);
    var subtotal = itemsTotal + transport;
    var cgstPct  = parseFloat(po.cgst_percent) || 0;
    var sgstPct  = parseFloat(po.sgst_percent) || 0;
    var cgstAmt  = Math.round(subtotal * cgstPct / 100);
    var sgstAmt  = Math.round(subtotal * sgstPct / 100);
    var grandTotal = parseFloat(po.total) || (subtotal + cgstAmt + sgstAmt);
    var totalQty = items.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);

    // Title
    var titleEl = document.getElementById('poItemsModalTitle');
    if (titleEl) titleEl.textContent = (po.po_no || 'PO') + '  —  ' + (po.vendor_name || '');

    // Build items table
    var html = '';
    var isMobile = _isPOMobileViewport_();
    if (!items.length && !transport) {
        html = '<div class="py-10 text-center text-base-content/40 text-sm">No line items recorded for this PO.</div>';
    } else {
        if (isMobile) {
            html += '<div class="po-items-mobile-grid">';
            items.forEach(function(item, i) {
                var total = parseFloat(item.total) || 0;
                html += '<div class="po-items-mobile-row">';
                html += '<div class="flex items-start justify-between gap-2">';
                html += '<div class="font-semibold text-sm text-slate-800">#' + (i + 1) + ' ' + escHtml(item.desc || '—') + '</div>';
                html += '<div class="font-mono text-sm font-bold text-emerald-600">' + (total ? '₹' + total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—') + '</div>';
                html += '</div>';
                html += '<div class="grid grid-cols-3 gap-2 mt-2">';
                html += '<div><div class="po-items-mobile-label">Qty</div><div class="font-mono text-sm">' + (item.qty || '') + '</div></div>';
                html += '<div><div class="po-items-mobile-label">UOM</div><div class="text-sm">' + escHtml(item.uom || 'NOS') + '</div></div>';
                html += '<div><div class="po-items-mobile-label">Days</div><div class="font-mono text-sm">' + (item.days || '') + '</div></div>';
                html += '</div>';
                html += '<div class="mt-2"><span class="po-items-mobile-label">Rate/Day: </span><span class="font-mono text-sm">' + (item.per_day ? Number(item.per_day).toLocaleString('en-IN') : '—') + '</span></div>';
                html += '</div>';
            });
            if (transport > 0) {
                html += '<div class="po-items-mobile-row">';
                html += '<div class="flex items-center justify-between text-sm italic text-base-content/70">';
                html += '<span>Transport Charges</span>';
                html += '<span class="font-mono">₹' + transport.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</span>';
                html += '</div></div>';
            }
            html += '<div class="po-items-mobile-row bg-base-200/70">';
            html += '<div class="flex items-center justify-between text-sm font-semibold">';
            html += '<span>Total Qty: <span class="font-mono">' + totalQty + '</span></span>';
            html += '<span>Subtotal: <span class="font-mono">₹' + subtotal.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</span></span>';
            html += '</div></div>';
            html += '</div>';
        } else {
            html += '<table class="table table-sm table-zebra w-full text-sm">';
            html += '<thead><tr class="bg-base-200">';
            html += '<th class="text-center w-10">#</th>';
            html += '<th>Item Description</th>';
            html += '<th class="text-right w-14">Qty</th>';
            html += '<th class="text-center w-16">UOM</th>';
            html += '<th class="text-right w-24">Rate/Day</th>';
            html += '<th class="text-right w-16">Days</th>';
            html += '<th class="text-right w-28">Total (INR)</th>';
            html += '</tr></thead><tbody>';

            items.forEach(function(item, i) {
                var total = parseFloat(item.total) || 0;
                html += '<tr>';
                html += '<td class="text-center text-base-content/50 font-mono text-xs">' + (i + 1) + '</td>';
                html += '<td class="font-medium">' + escHtml(item.desc || '—') + '</td>';
                html += '<td class="text-right font-mono">' + (item.qty || '') + '</td>';
                html += '<td class="text-center">';
                html += '<span class="badge badge-ghost badge-sm">' + escHtml(item.uom || 'NOS') + '</span>';
                html += '</td>';
                html += '<td class="text-right font-mono">' + (item.per_day ? Number(item.per_day).toLocaleString('en-IN') : '—') + '</td>';
                html += '<td class="text-right font-mono">' + (item.days || '') + '</td>';
                html += '<td class="text-right font-semibold font-mono">' + (total ? '\u20b9' + total.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—') + '</td>';
                html += '</tr>';
            });

            if (transport > 0) {
                html += '<tr class="italic text-base-content/60">';
                html += '<td class="text-center">+</td>';
                html += '<td colspan="5">Transport Charges</td>';
                html += '<td class="text-right font-mono">₹' + transport.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</td>';
                html += '</tr>';
            }

            html += '<tr class="font-bold bg-base-200">';
            html += '<td></td>';
            html += '<td>Total</td>';
            html += '<td class="text-right font-mono">' + totalQty + '</td>';
            html += '<td colspan="3"></td>';
            html += '<td class="text-right font-mono">₹' + subtotal.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</td>';
            html += '</tr>';
            html += '</tbody></table>';
        }

        // Tax + Grand Total summary strip
        html += '<div class="mt-3 flex flex-col items-end gap-1 pr-1 pb-1">';
        if (cgstPct) {
            html += '<div class="flex items-center gap-6 text-sm">';
            html += '<span class="text-base-content/60">Add : CGST @ ' + cgstPct + '%</span>';
            html += '<span class="font-mono font-semibold w-28 text-right">₹' + cgstAmt.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</span>';
            html += '</div>';
        }
        if (sgstPct) {
            html += '<div class="flex items-center gap-6 text-sm">';
            html += '<span class="text-base-content/60">Add : SGST @ ' + sgstPct + '%</span>';
            html += '<span class="font-mono font-semibold w-28 text-right">₹' + sgstAmt.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</span>';
            html += '</div>';
        }
        html += '<div class="flex items-center gap-6 mt-1 pt-2 border-t border-base-300">';
        html += '<span class="text-base font-bold uppercase tracking-wide">Grand Total</span>';
        html += '<span class="font-mono font-bold text-lg text-primary w-28 text-right">₹' + grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</span>';
        html += '</div></div>';
    }

    var bodyEl = document.getElementById('poItemsModalBody');
    if (bodyEl) bodyEl.innerHTML = html;

    var modal = document.getElementById('poItemsModal');
    if (modal) modal.showModal();
}

// Prepare a PO from _poList for buildPODocument / Excel export
function _preparePO(idx) {
    var po = (window._poList || [])[idx];
    if (!po) return null;
    var items = po.items || [];
    if (typeof items === 'string') { try { items = JSON.parse(items); } catch(e) { items = []; } }
    po._items = items;
    return po;
}

function _sanitizeFileNamePart_(value) {
    var text = String(value || '')
        .replace(/[\\/:*?"<>|]+/g, ' - ')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\s+/g, ' ')
        .replace(/\s*-\s*/g, '-')
        .trim();
    return text;
}

function _stripBusinessPrefix_(value) {
    var text = _sanitizeFileNamePart_(value);
    return text.replace(/^m\s*\/?\s*s\.?\s*/i, '').replace(/^ms\.?\s*/i, '').trim();
}

// Builds export names in this pattern:
// RE-EO-001-VENDOR NAME-(CLIENT NAME - LOCATION)
function getPOExportFileBaseName(po) {
    var poNo = _sanitizeFileNamePart_(po && po.po_no) || 'RE-EO-001';
    var vendorName = _stripBusinessPrefix_(po && po.vendor_name) || 'VENDOR NAME';
    var clientName = _stripBusinessPrefix_(po && po.client_name) || 'CLIENT NAME';
    var locationName = _sanitizeFileNamePart_(po && (po.event_location || po.location));
    var clientLocation = locationName ? (clientName + ' - ' + locationName) : clientName;
    return poNo + '-' + vendorName + '-(' + clientLocation + ')';
}

// Download PDF for a PO from the list page
function listDownloadPDF(idx) {
    var po = _preparePO(idx);
    if (!po) return alert('PO not found');
    var doc = buildPODocument({
        po_no: po.po_no, po_date: po.po_date,
        vendor_name: po.vendor_name, vendor_address: po.vendor_address, vendor_gstin: po.vendor_gstin,
        client_name: po.client_name, event_name: po.event_name, event_location: po.event_location, event_date: po.event_date,
        transport: po.transport, cgst_percent: po.cgst_percent, sgst_percent: po.sgst_percent,
        terms: po.terms, items: po._items,
        includeSign: true
    });
    if (doc) doc.save(getPOExportFileBaseName(po) + '.pdf');
}

// Download Excel for a PO from the list page
function _downloadBlobFile_(fileName, blob) {
    var link = document.createElement('a');
    var url = URL.createObjectURL(blob);
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    setTimeout(function() {
        URL.revokeObjectURL(url);
        link.remove();
    }, 500);
}

async function _downloadPOExcelStyled_(po, fileName) {

    if (typeof ExcelJS === 'undefined') return false;

    // --- Helper row management for worksheet ---
    var rowNo = 1;
    function add(arr) {
        ws.addRow(arr);
        return rowNo++;
    }
    function mergeRow(r) {
        ws.mergeCells('A' + r + ':G' + r);
    }
    // --- Helper: Apply border to a rectangular cell range (ExcelJS) ---
    function styleRangeBorder(startCell, endCell, borderStyle) {
        // startCell, endCell: e.g. 'A2', 'D5'
        var startCol = startCell.charCodeAt(0) - 65 + 1;
        var startRow = parseInt(startCell.slice(1), 10);
        var endCol = endCell.charCodeAt(0) - 65 + 1;
        var endRow = parseInt(endCell.slice(1), 10);
        for (var r = startRow; r <= endRow; r++) {
            for (var c = startCol; c <= endCol; c++) {
                ws.getRow(r).getCell(c).border = borderStyle;
            }
        }
    }

    var items = po._items || [];
    var transport = parseFloat(po.transport) || 0;
    var itemsTotal = items.reduce(function(s, i) { return s + (parseFloat(i.total) || 0); }, 0);
    var subtotal = itemsTotal + transport;
    var cgstPct = parseFloat(po.cgst_percent) || 0;
    var sgstPct = parseFloat(po.sgst_percent) || 0;
    var cgstAmt = Math.round(subtotal * cgstPct / 100);
    var sgstAmt = Math.round(subtotal * sgstPct / 100);
    var grandTotal = parseFloat(po.total) || (subtotal + cgstAmt + sgstAmt);
    var totalQty = items.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);

    var wb = new ExcelJS.Workbook();
    wb.creator = 'PO Manager';
    wb.created = new Date();
    var ws = wb.addWorksheet('PurchaseOrder', {
        pageSetup: {
            paperSize: 9, // A4
            orientation: 'portrait',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            horizontalCentered: true,
            margins: { left: 0.35, right: 0.35, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 }
        }
    });

    ws.views = [{ state: 'frozen', ySplit: 16 }];
    ws.columns = [
        { width: 8 },
        { width: 37 },
        { width: 11 },
        { width: 9 },
        { width: 15 },
        { width: 11 },
        { width: 17 }
    ];

    var thinGray = { style: 'thin', color: { argb: 'FFD1D5DB' } };
    var gridBorder = { top: thinGray, left: thinGray, bottom: thinGray, right: thinGray };

    var rows = [[
        po.po_no || '',
        formatDateDDMMMYYYY(po.po_date) || '',
        po.vendor_name || '',
        po.vendor_address || '',
        po.vendor_gstin || '',
        po.client_name || '',
        po.event_name || '',
        po.event_location || '',
        formatDateDDMMMYYYY(po.event_date) || '',
        _getPOTotal_(po)
    ]];

    // Header block (PDF style)
    var r1 = add(['RANGA ELECTRICALS PVT LTD']); mergeRow(r1);
    var r2 = add(['NO.326, RAMAKRISHNA NAGAR MAIN ROAD, PORUR, CHENNAI-600116.']); mergeRow(r2);
    var r3 = add(['E-mail: info@rangaelectricals.com']); mergeRow(r3);
    var r4 = add(['GST IN: 33AAHCR4037J1ZD']); mergeRow(r4);
    var r5 = add(['']); mergeRow(r5);
    var r6 = add(['PURCHASE ORDER']); mergeRow(r6);
    var r7 = add(['']); mergeRow(r7);

    ws.getCell('A' + r1).font = { bold: true, size: 17, color: { argb: 'FF1E3A8A' } };
    ws.getCell('A' + r2).font = { size: 10, color: { argb: 'FF334155' } };
    ws.getCell('A' + r3).font = { size: 9, color: { argb: 'FF64748B' } };
    ws.getCell('A' + r4).font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };
    ws.getCell('A' + r6).font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
    ws.getCell('A' + r6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
    ['A' + r1, 'A' + r2, 'A' + r3, 'A' + r4, 'A' + r6].forEach(function(addr) {
        ws.getCell(addr).alignment = { horizontal: 'center', vertical: 'middle' };
    });

    ws.getRow(r1).height = 24;
    ws.getRow(r2).height = 16;
    ws.getRow(r3).height = 14;
    ws.getRow(r4).height = 16;
    ws.getRow(r6).height = 22;

    // Two detail boxes (left vendor / right PO info)
    var boxTop = add(['To,', '', '', '', 'PO NO:', po.po_no || '', '']);
    ws.mergeCells('A' + boxTop + ':D' + boxTop);
    ws.mergeCells('F' + boxTop + ':G' + boxTop);
    ws.getCell('A' + boxTop).font = { italic: true, size: 9, color: { argb: 'FF6B7280' } };
    ws.getCell('E' + boxTop).font = { size: 9, color: { argb: 'FF6B7280' }, bold: true };
    ws.getCell('F' + boxTop).alignment = { horizontal: 'right' };
    ws.getCell('F' + boxTop).font = { bold: true, color: { argb: 'FF2563EB' } };

    var boxR2 = add([po.vendor_name || '', '', '', '', 'DATE:', po.po_date || '', '']);
    ws.mergeCells('A' + boxR2 + ':D' + boxR2);
    ws.mergeCells('F' + boxR2 + ':G' + boxR2);
    ws.getCell('A' + boxR2).font = { bold: true, size: 11 };
    ws.getCell('E' + boxR2).font = { size: 9, color: { argb: 'FF6B7280' }, bold: true };
    ws.getCell('F' + boxR2).alignment = { horizontal: 'right' };
    ws.getCell('F' + boxR2).font = { bold: true };

    var boxR3 = add([po.vendor_address || '', '', '', '', 'Client Name:', po.client_name ? ('M/s. ' + po.client_name) : '', '' ]);
    ws.mergeCells('A' + boxR3 + ':D' + boxR3);
    ws.mergeCells('F' + boxR3 + ':G' + boxR3);
    ws.getCell('A' + boxR3).alignment = { wrapText: true, vertical: 'top' };
    ws.getCell('A' + boxR3).font = { size: 9 };
    ws.getCell('E' + boxR3).font = { size: 9, color: { argb: 'FF6B7280' }, bold: true };
    ws.getCell('F' + boxR3).font = { bold: true, size: 9 };
    ws.getCell('F' + boxR3).alignment = { horizontal: 'right' };

    var boxR4 = add([po.vendor_gstin ? ('GSTIN: ' + po.vendor_gstin) : '', '', '', '', 'Event Name:', po.event_name || '', '']);
    ws.mergeCells('A' + boxR4 + ':D' + boxR4);
    ws.mergeCells('F' + boxR4 + ':G' + boxR4);
    ws.getCell('A' + boxR4).font = { size: 9 };
    ws.getCell('E' + boxR4).font = { size: 9, color: { argb: 'FF6B7280' }, bold: true };
    ws.getCell('F' + boxR4).font = { bold: true, size: 9 };
    ws.getCell('F' + boxR4).alignment = { horizontal: 'right' };

    var boxR5 = add(['', '', '', '', 'Event Date:', po.event_date || '', '']);
    ws.mergeCells('A' + boxR5 + ':D' + boxR5);
    ws.mergeCells('F' + boxR5 + ':G' + boxR5);
    ws.getCell('E' + boxR5).font = { size: 9, color: { argb: 'FF6B7280' }, bold: true };
    ws.getCell('F' + boxR5).font = { bold: true, size: 9 };
    ws.getCell('F' + boxR5).alignment = { horizontal: 'right' };

    styleRangeBorder('A' + boxTop, 'D' + boxR5, gridBorder);
    styleRangeBorder('E' + boxTop, 'G' + boxR5, gridBorder);
    ws.getRow(boxR3).height = po.vendor_address ? 36 : 18;

    var instRow = add(['We are pleased to place our order in your favour for the supply of following items']);
    ws.mergeCells('A' + instRow + ':G' + instRow);
    ws.getCell('A' + instRow).font = { italic: true, size: 9, color: { argb: 'FF475569' } };

    var headerRow = add(['Sl.No', 'Item Description', 'Qty', 'UOM', 'Per Day Amt', 'Days', 'Total (INR)']);
    ws.getRow(headerRow).height = 22;
    for (var c = 1; c <= 7; c++) {
        var hCell = ws.getRow(headerRow).getCell(c);
        hCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        hCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
        hCell.alignment = { horizontal: 'center', vertical: 'middle' };
        hCell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
        };
    }

    var dataStart = headerRow + 1;
    items.forEach(function(item, idx) {
        add([
            idx + 1,
            item.desc || '',
            parseFloat(item.qty) || 0,
            item.uom || '',
            parseFloat(item.per_day) || 0,
            parseFloat(item.days) || 0,
            parseFloat(item.total) || 0
        ]);
    });
    if (transport > 0) add([items.length + 1, 'TRANSPORT CHARGES', '', '', '', '', transport]);

    var dataEnd = rowNo - 1;
    for (var r = dataStart; r <= dataEnd; r++) {
        for (var cc = 1; cc <= 7; cc++) {
            var cell = ws.getRow(r).getCell(cc);
            cell.border = gridBorder;
            if (cc === 2) cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
            if (cc === 1 || cc === 3 || cc === 4 || cc === 6) cell.alignment = { horizontal: 'center', vertical: 'middle' };
            if (cc === 5 || cc === 7) {
                cell.alignment = { horizontal: 'right', vertical: 'middle' };
                cell.numFmt = '#,##0.00';
            }
        }
        ws.getRow(r).height = 20;
    }

    var spacer = add([]);
    ws.getRow(spacer).height = 8;
    var totalRow = add(['', 'Total', totalQty, '', '', '', subtotal]);
    var cgstRow = cgstPct ? add(['', '', '', '', '', 'Add : CGST @ ' + cgstPct + '%', cgstAmt]) : 0;
    var sgstRow = sgstPct ? add(['', '', '', '', '', 'Add : SGST @ ' + sgstPct + '%', sgstAmt]) : 0;
    var grandRow = add(['', '', '', '', '', 'GRAND TOTAL', grandTotal]);

    [totalRow, cgstRow, sgstRow, grandRow].forEach(function(rr) {
        if (!rr) return;
        for (var c2 = 6; c2 <= 7; c2++) {
            var ccell = ws.getRow(rr).getCell(c2);
            ccell.border = gridBorder;
            ccell.alignment = { horizontal: 'right', vertical: 'middle' };
            if (c2 === 7) ccell.numFmt = '#,##0.00';
        }
    });
    ws.getRow(totalRow).font = { bold: true };
    if (cgstRow) ws.getRow(cgstRow).font = { italic: true };
    if (sgstRow) ws.getRow(sgstRow).font = { italic: true };
    ws.getRow(grandRow).font = { bold: true, size: 12, color: { argb: 'FF92400E' } };
    ws.getCell('F' + grandRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
    ws.getCell('G' + grandRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

    var blankAfterTotal = add([]);
    ws.getRow(blankAfterTotal).height = 8;

    if (po.terms) {
        var termsRow = add(['TERMS OF PAYMENT: ' + po.terms]);
        ws.mergeCells('A' + termsRow + ':G' + termsRow);
        ws.getCell('A' + termsRow).font = { size: 9 };
        ws.getCell('A' + termsRow).alignment = { wrapText: true };
    }

    var dividerRow = add([]);
    for (var dc = 1; dc <= 7; dc++) {
        ws.getRow(dividerRow).getCell(dc).border = { top: { style: 'thin', color: { argb: 'FF64748B' } } };
    }

    var evRow = add(['EVENT LOCATION', po.event_name || '', '', '', '', 'for RANGA ELECTRICALS PVT LTD', '']);
    ws.mergeCells('A' + evRow + ':A' + evRow);
    ws.mergeCells('B' + evRow + ':D' + evRow);
    ws.mergeCells('F' + evRow + ':G' + evRow);
    ws.getCell('A' + evRow).font = { bold: true, size: 9 };
    ws.getCell('B' + evRow).font = { bold: true, size: 9 };
    ws.getCell('F' + evRow).font = { italic: true, size: 9 };
    ws.getCell('F' + evRow).alignment = { horizontal: 'right' };

    if (po.event_location) {
        var ev2Row = add(['', po.event_location, '', '', '', '', '']);
        ws.mergeCells('B' + ev2Row + ':D' + ev2Row);
        ws.getCell('B' + ev2Row).alignment = { wrapText: true };
        ws.getRow(ev2Row).height = 26;
    }

    var signSpacer = add([]);
    ws.getRow(signSpacer).height = 22;
    var signLine = add(['', '', '', '', '', '____________________', '']);
    ws.mergeCells('F' + signLine + ':G' + signLine);
    ws.getCell('F' + signLine).alignment = { horizontal: 'right' };
    ws.getCell('F' + signLine).font = { color: { argb: 'FF6B7280' } };
    var signText = add(['', '', '', '', '', 'Authorised Signatory', '']);
    ws.mergeCells('F' + signText + ':G' + signText);
    ws.getCell('F' + signText).alignment = { horizontal: 'right' };
    ws.getCell('F' + signText).font = { bold: true, size: 9 };

    ws.headerFooter.oddHeader = '&C&RANGA ELECTRICALS PVT LTD';
    ws.headerFooter.oddFooter = '&LPO: ' + (po.po_no || '') + '&RPage &P of &N';
    ws.pageSetup.printArea = 'A1:G' + (rowNo - 1);

    var buffer = await wb.xlsx.writeBuffer();
    _downloadBlobFile_(fileName, new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    return true;
}

async function downloadPOExcelFromObject(po) {
    if (!po) return;
    var fileName = getPOExportFileBaseName(po) + '.xlsx';
    var styledOk = await _downloadPOExcelStyled_(po, fileName);
    if (styledOk) return;

    // Fallback to legacy SheetJS export
    var items = po._items || [];
    var transport = parseFloat(po.transport) || 0;
    var itemsTotal = items.reduce(function(s, i) { return s + (parseFloat(i.total) || 0); }, 0);
    var subtotal = itemsTotal + transport;
    var cgstPct = parseFloat(po.cgst_percent) || 0;
    var sgstPct = parseFloat(po.sgst_percent) || 0;
    var cgstAmt = Math.round(subtotal * cgstPct / 100);
    var sgstAmt = Math.round(subtotal * sgstPct / 100);
    var grandTotal = parseFloat(po.total) || (subtotal + cgstAmt + sgstAmt);
    var totalQty = items.reduce(function(s, i) { return s + (parseInt(i.qty) || 0); }, 0);

    var rows = [];
    rows.push(['RANGA ELECTRICALS PVT LTD']);
    rows.push(['NO.326, RAMAKRISHNA NAGAR MAIN ROAD, PORUR, CHENNAI-600116.']);
    rows.push(['GST IN: 33AAHCR4037J1ZD']);
    rows.push([]);
    rows.push(['PURCHASE ORDER']);
    rows.push([]);
    rows.push(['To:', po.vendor_name || '', '', 'PO NO:', po.po_no || '']);
    if (po.vendor_address) rows.push(['', po.vendor_address, '', 'DATE:', po.po_date || '']);
    else rows.push(['', '', '', 'DATE:', po.po_date || '']);
    if (po.vendor_gstin) rows.push(['GSTIN:', po.vendor_gstin]);
    rows.push([]);
    if (po.client_name) rows.push(['Client Name:', 'M/s. ' + po.client_name]);
    if (po.event_date) rows.push(['Event Date:', po.event_date]);
    rows.push([]);
    rows.push(['Sl.No', 'Item Description', 'Quantity', 'UOM', 'Per Day Amount', 'No of Days', 'Total Value (INR)']);
    items.forEach(function(item, i) {
        rows.push([i + 1, item.desc || '', item.qty || '', item.uom || '', item.per_day || 0, item.days || 0, item.total || 0]);
    });
    if (transport > 0) rows.push([items.length + 1, 'TRANSPORT CHARGES', '', '', '', '', transport]);
    rows.push([]);
    rows.push(['', 'Total', totalQty, '', '', '', subtotal]);
    if (cgstPct) rows.push(['', '', '', '', '', 'Add : CGST @ ' + cgstPct + '%', cgstAmt]);
    if (sgstPct) rows.push(['', '', '', '', '', 'Add : SGST @ ' + sgstPct + '%', sgstAmt]);
    rows.push(['', '', '', '', '', 'GRAND TOTAL', grandTotal]);
    rows.push([]);
    if (po.terms) rows.push(['TERMS OF PAYMENT:', po.terms]);
    rows.push([]);
    rows.push(['EVENT LOCATION:', po.event_name || '']);
    if (po.event_location) rows.push(['', po.event_location]);

    if (typeof XLSX === 'undefined') return alert('Excel library not loaded');
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 8 }, { wch: 38 }, { wch: 10 }, { wch: 8 }, { wch: 16 }, { wch: 16 }, { wch: 20 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PurchaseOrder');
    XLSX.writeFile(wb, fileName);
}

function listDownloadExcel(idx) {
    var po = _preparePO(idx);
    if (!po) return alert('PO not found');
    downloadPOExcelFromObject(po);
}

function exportAllPOsExcel(sourceList) {
    var list = Array.isArray(sourceList) ? sourceList.slice() : (Array.isArray(window._poList) ? window._poList.slice() : []);
    if (!list.length) {
        showToast('No POs available to export.', 'warning');
        return;
    }

    var fileName = 'All_Purchase_Orders_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    _downloadAllPOsExcelStyled_(list, fileName)
        .then(function(ok) {
            if (ok) {
                showToast('All POs exported successfully.', 'success');
                return;
            }
            _downloadAllPOsExcelFallback_(list, fileName);
            showToast('All POs exported (basic format).', 'info');
        })
        .catch(function(err) {
            console.error('All PO export failed:', err);
            _downloadAllPOsExcelFallback_(list, fileName);
            showToast('All POs exported (fallback).', 'warning');
        });
}

async function _downloadAllPOsExcelStyled_(list, fileName) {
    if (typeof ExcelJS === 'undefined') return false;

    var wb = new ExcelJS.Workbook();
    wb.creator = 'PO Manager';
    wb.created = new Date();
    var ws = wb.addWorksheet('All POs');

    ws.columns = [
        { width: 7 },   // S.No
        { width: 16 },  // PO No
        { width: 13 },  // PO Date
        { width: 26 },  // Vendor
        { width: 20 },  // Client
        { width: 20 },  // Event
        { width: 28 },  // Event Location
        { width: 13 },  // Event Date
        { width: 34 },  // Item
        { width: 9 },   // Qty
        { width: 9 },   // NOS
        { width: 14 },  // Rate
        { width: 10 },  // Days
        { width: 16 },  // Item Amount
        { width: 16 }   // PO Total
    ];

    var moneyFmt = '#,##0.00';
    var qtyFmt = '#,##0.##';

    function applyBorder(cell) {
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
    }

    ws.mergeCells('A1:O1');
    ws.getCell('A1').value = 'ALL PURCHASE ORDERS - DETAILED EXPORT';
    ws.getCell('A1').font = { bold: true, size: 16, color: { argb: 'FF1E3A8A' } };
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 24;

    ws.mergeCells('A2:O2');
    ws.getCell('A2').value = 'Generated on: ' + new Date().toLocaleString();
    ws.getCell('A2').font = { size: 10, color: { argb: 'FF64748B' } };
    ws.getCell('A2').alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(2).height = 18;

    var headerRow = 4;
    ws.getRow(headerRow).values = ['S.No', 'PO No', 'PO Date', 'Vendor', 'Client', 'Event', 'Event Location', 'Event Date', 'Item Description', 'Qty', 'NOS', 'Rate', 'Days', 'Item Amount', 'PO Total'];
    ws.getRow(headerRow).height = 22;
    ws.getRow(headerRow).eachCell(function(cell) {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
        applyBorder(cell);
    });

    var row = headerRow + 1;
    var grandAll = 0;
    var bandA = 'FFF8FAFC';
    var bandB = 'FFF0F9FF';

    for (var i = 0; i < list.length; i++) {
        var po = list[i] || {};
        var items = po.items || [];
        if (typeof items === 'string') {
            try { items = JSON.parse(items); } catch (e) { items = []; }
        }
        if (!Array.isArray(items) || !items.length) items = [{ desc: '-', qty: 0, uom: 'NOS', per_day: 0, days: 0, total: 0 }];

        var poTotal = parseFloat(po.total) || 0;
        if (!poTotal) {
            var transport = parseFloat(po.transport) || 0;
            var itemsTotal = items.reduce(function(s, it) { return s + (parseFloat(it.total) || 0); }, 0);
            var sub = itemsTotal + transport;
            var cgst = sub * (parseFloat(po.cgst_percent) || 0) / 100;
            var sgst = sub * (parseFloat(po.sgst_percent) || 0) / 100;
            poTotal = Math.round(sub + cgst + sgst);
        }
        grandAll += poTotal;

        var blockStart = row;
        for (var j = 0; j < items.length; j++) {
            var item = items[j] || {};
            var qty = parseFloat(item.qty) || 0;
            var rate = parseFloat(item.per_day) || 0;
            var days = parseFloat(item.days) || 0;
            var amount = parseFloat(item.total) || 0;

            ws.getRow(row).values = [
                i + 1,
                po.po_no || '',
                formatDateDDMMMYYYY(po.po_date) || '',
                po.vendor_name || '',
                po.client_name || '',
                po.event_name || '',
                po.event_location || '',
                formatDateDDMMMYYYY(po.event_date) || '',
                item.desc || '-',
                qty,
                item.uom || 'NOS',
                rate,
                days,
                amount,
                poTotal
            ];

            var fillColor = (i % 2 === 0) ? bandA : bandB;
            for (var c = 1; c <= 15; c++) {
                var cell = ws.getRow(row).getCell(c);
                applyBorder(cell);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
                if (c === 9) cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
                else if (c === 7) cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
                else if ([10, 11, 12, 13, 14, 15].indexOf(c) !== -1) cell.alignment = { horizontal: 'right', vertical: 'middle' };
                else cell.alignment = { horizontal: 'left', vertical: 'middle' };
            }

            ws.getCell('J' + row).numFmt = qtyFmt;
            ws.getCell('L' + row).numFmt = moneyFmt;
            ws.getCell('M' + row).numFmt = qtyFmt;
            ws.getCell('N' + row).numFmt = moneyFmt;
            ws.getCell('O' + row).numFmt = moneyFmt;
            row++;
        }

        var blockEnd = row - 1;
        if (blockEnd > blockStart) {
            ['A','B','C','D','E','F','G','H','O'].forEach(function(col) {
                ws.mergeCells(col + blockStart + ':' + col + blockEnd);
                var mc = ws.getCell(col + blockStart);
                mc.alignment = { horizontal: (col === 'O' ? 'right' : 'left'), vertical: 'middle', wrapText: true };
                mc.font = { bold: col === 'O' || col === 'B' };
            });
        } else {
            ws.getCell('B' + blockStart).font = { bold: true };
            ws.getCell('O' + blockStart).font = { bold: true };
        }

        // Visual separator between PO blocks.
        if (i !== list.length - 1) {
            ws.getRow(row).height = 6;
            for (var sc = 1; sc <= 15; sc++) {
                var sepCell = ws.getRow(row).getCell(sc);
                sepCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
                sepCell.border = {
                    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                    left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                    bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                    right: { style: 'thin', color: { argb: 'FFFFFFFF' } }
                };
            }
            row++;
        }
    }

    var summaryRow = row + 1;
    ws.mergeCells('A' + summaryRow + ':N' + summaryRow);
    ws.getCell('A' + summaryRow).value = 'GRAND TOTAL OF ALL POs';
    ws.getCell('A' + summaryRow).font = { bold: true, color: { argb: 'FF92400E' } };
    ws.getCell('A' + summaryRow).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell('A' + summaryRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
    ws.getCell('O' + summaryRow).value = grandAll;
    ws.getCell('O' + summaryRow).numFmt = moneyFmt;
    ws.getCell('O' + summaryRow).font = { bold: true, color: { argb: 'FF92400E' } };
    ws.getCell('O' + summaryRow).alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getCell('O' + summaryRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
    applyBorder(ws.getCell('A' + summaryRow));
    applyBorder(ws.getCell('O' + summaryRow));

    ws.autoFilter = {
        from: { row: headerRow, column: 1 },
        to: { row: headerRow, column: 15 }
    };
    ws.views = [{ state: 'frozen', ySplit: headerRow }];

    var buffer = await wb.xlsx.writeBuffer();
    _downloadBlobFile_(fileName, new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    return true;
}

function _downloadAllPOsExcelFallback_(list, fileName) {
    if (typeof XLSX === 'undefined') {
        showToast('Excel export library not loaded.', 'error');
        return;
    }

    var rows = [['S.No', 'PO No', 'PO Date', 'Vendor', 'Client', 'Event', 'Event Location', 'Event Date', 'Item Description', 'Qty', 'NOS', 'Rate', 'Days', 'Item Amount', 'PO Total']];
    var merges = [];
    var ptr = 2;

    list.forEach(function(po, i) {
        var items = po.items || [];
        if (typeof items === 'string') {
            try { items = JSON.parse(items); } catch (e) { items = []; }
        }
        if (!Array.isArray(items) || !items.length) items = [{ desc: '-', qty: 0, uom: 'NOS', per_day: 0, days: 0, total: 0 }];

        var poTotal = parseFloat(po.total) || 0;
        var start = ptr;
        items.forEach(function(item, idx) {
            rows.push([
                i + 1,
                po.po_no || '',
                formatDateDDMMMYYYY(po.po_date) || '',
                po.vendor_name || '',
                po.client_name || '',
                po.event_name || '',
                po.event_location || '',
                formatDateDDMMMYYYY(po.event_date) || '',
                item.desc || '-',
                parseFloat(item.qty) || 0,
                item.uom || 'NOS',
                parseFloat(item.per_day) || 0,
                parseFloat(item.days) || 0,
                parseFloat(item.total) || 0,
                idx === 0 ? poTotal : ''
            ]);
            ptr++;
        });
        var end = ptr - 1;
        if (end > start) {
            [0,1,2,3,4,5,6,7,14].forEach(function(c) {
                merges.push({ s: { r: start - 1, c: c }, e: { r: end - 1, c: c } });
            });
        }
    });

    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!merges'] = merges;
    ws['!cols'] = [
        { wch: 7 }, { wch: 14 }, { wch: 11 }, { wch: 24 }, { wch: 18 },
        { wch: 18 }, { wch: 26 }, { wch: 12 }, { wch: 34 }, { wch: 8 },
        { wch: 8 }, { wch: 12 }, { wch: 9 }, { wch: 14 }, { wch: 14 }
    ];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'All POs');
    XLSX.writeFile(wb, fileName);
}

// Print PDF for a PO from the list page
function listPrintPDF(idx) {
    var po = _preparePO(idx);
    if (!po) return alert('PO not found');
    var doc = buildPODocument({
        po_no: po.po_no, po_date: po.po_date,
        vendor_name: po.vendor_name, vendor_address: po.vendor_address, vendor_gstin: po.vendor_gstin,
        client_name: po.client_name, event_name: po.event_name, event_location: po.event_location, event_date: po.event_date,
        transport: po.transport, cgst_percent: po.cgst_percent, sgst_percent: po.sgst_percent,
        terms: po.terms, items: po._items,
        includeSign: true
    });
    if (doc) {
        var blob = doc.output('blob');
        var url = URL.createObjectURL(blob);
        var printWin = window.open(url, '_blank');
        if (printWin) {
            printWin.onload = function() { printWin.focus(); printWin.print(); };
        }
    }
}

// ============================================================
// PO FORM (Add PO page)
// ============================================================
let poItems = [];

// ── Set item type (Master / Custom) ──
function poSetType(idx, isMaster) {
    poItems[idx].isMaster = isMaster;
    if (isMaster) { poItems[idx].desc = ''; poItems[idx].uom = 'NOS'; poItems[idx].per_day = 0; }
    renderPOItemsTable();
    setTimeout(function() { var el = document.getElementById('piInput' + idx); if (el) el.focus(); }, 30);
}

// ── Search dropdown: open on focus ──
function poItemSearchOpen(idx) {
    var dd = document.getElementById('piDD' + idx);
    if (dd) { dd.classList.add('open'); }
}

// ── Search dropdown: hide on blur (delayed so mousedown on item fires first) ──
function poItemSearchHide(idx) {
    setTimeout(function() {
        var dd = document.getElementById('piDD' + idx);
        if (dd) dd.classList.remove('open');
    }, 200);
}

// ── Filter dropdown items on input ──
function poItemSearchFilter(idx, val) {
    poItems[idx].desc = val;
    var dd = document.getElementById('piDD' + idx);
    if (!dd) return;
    var q = (val || '').toLowerCase().trim();
    var els = dd.querySelectorAll('.po-item-dd-item');
    var visible = 0;
    els.forEach(function(el) {
        var show = !q || (el.dataset.name || '').toLowerCase().indexOf(q) !== -1;
        el.style.display = show ? '' : 'none';
        if (show) visible++;
    });
    dd.classList.add('open');
}

// ── Select an item from the master dropdown ──
function poItemSearchSelect(idx, masterIdx) {
    if (!itemMasterCache || !itemMasterCache[masterIdx]) return;
    var mi = itemMasterCache[masterIdx];
    poItems[idx].desc = mi.desc || '';
    poItems[idx].uom = mi.uom || 'NOS';
    poItems[idx].per_day = parseFloat(mi.rate) || 0;
    poItems[idx].masterIdx = masterIdx;
    poItems[idx].total = (parseInt(poItems[idx].qty) || 1) * poItems[idx].per_day * (parseInt(poItems[idx].days) || 1);
    renderPOItemsTable();
    setTimeout(function() {
        var row = document.getElementById('poRow' + idx);
        if (row) { var qi = row.querySelector('.po-qty-cell input'); if (qi) qi.focus(); }
    }, 30);
}

// ── Keyboard navigation in search dropdown ──
function poItemSearchKey(e, idx) {
    var dd = document.getElementById('piDD' + idx);
    if (!dd) return;
    var vis = Array.from(dd.querySelectorAll('.po-item-dd-item')).filter(function(el) { return el.style.display !== 'none'; });
    if (!vis.length && e.key !== 'Escape') return;
    var hl = dd.querySelector('.po-item-dd-item.highlighted');
    var hi = hl ? vis.indexOf(hl) : -1;
    if (e.key === 'ArrowDown') {
        e.preventDefault();
        var next = vis[hi + 1] || vis[0];
        if (hl) hl.classList.remove('highlighted');
        next.classList.add('highlighted');
        next.scrollIntoView({ block: 'nearest' });
    } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        var prev = vis[hi - 1] || vis[vis.length - 1];
        if (hl) hl.classList.remove('highlighted');
        prev.classList.add('highlighted');
        prev.scrollIntoView({ block: 'nearest' });
    } else if (e.key === 'Enter' && hl) {
        e.preventDefault();
        poItemSearchSelect(idx, parseInt(hl.dataset.idx));
    } else if (e.key === 'Escape') {
        dd.classList.remove('open');
    }
}

// ── Enter key in number inputs → advance to next field ──
function poNumKeyNav(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        var all = Array.from(document.querySelectorAll('#poItemsTable .po-cell-input'));
        var i = all.indexOf(e.target);
        if (i !== -1 && all[i + 1]) all[i + 1].focus();
    }
}

// ── Recalculate & update ONE row's total cell without full re-render ──
// Called by oninput on Qty / Rate / Days so other rows are never cleared.
function poCalcRow(idx) {
    if (!poItems[idx]) return;
    var qty  = parseFloat(poItems[idx].qty)  || 0;
    var rate = parseFloat(poItems[idx].per_day) || 0;
    var days = parseFloat(poItems[idx].days) || 0;
    var rowTotal = qty * rate * days;
    poItems[idx].total = rowTotal;

    // Update the total-amount cell
    var totalAmt = document.querySelector('#poRow' + idx + ' .po-total-amount');
    if (totalAmt) {
        totalAmt.className = 'po-total-amount' + (rowTotal ? '' : ' zero');
        totalAmt.textContent = rowTotal > 0
            ? '\u20b9' + rowTotal.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 })
            : '\u2014';
        // Flash the cell to give feedback that the calc updated
        totalAmt.classList.remove('calc-flash');
        void totalAmt.offsetWidth;
        totalAmt.classList.add('calc-flash');
    }

    // Update or create the calc-preview beneath the total
    var totalCell = document.querySelector('#poRow' + idx + ' .po-total-cell');
    var calcPrev  = document.querySelector('#poRow' + idx + ' .po-calc-preview');
    if (qty && rate && days) {
        var previewTxt = qty + '\u00d7' + rate + '\u00d7' + days;
        if (calcPrev) {
            calcPrev.textContent = previewTxt;
        } else if (totalCell) {
            var cp = document.createElement('div');
            cp.className = 'po-calc-preview';
            cp.textContent = previewTxt;
            totalCell.appendChild(cp);
        }
    } else if (calcPrev) {
        calcPrev.remove();
    }

    // Persist to hidden field and refresh summary bar only
    var hidden = document.getElementById('items');
    if (hidden) hidden.value = JSON.stringify(poItems);
    updatePOTotal();
    if (typeof generateLivePreview === 'function') generateLivePreview();
}

// ── Duplicate a row ──
function duplicatePORow(idx) {
    var copy = JSON.parse(JSON.stringify(poItems[idx]));
    poItems.splice(idx + 1, 0, copy);
    renderPOItemsTable();
    showToast('Row duplicated', 'info');
}

// ── Backward-compat stubs (called by edit-po pages) ──
function buildItemMasterOptionsHtml() {
    var html = '<option value="">-- Select Item --</option>';
    (itemMasterCache || []).forEach(function(item, idx) {
        html += '<option value="' + idx + '">' + escHtml(item.desc) + ' (' + escHtml(item.uom) + ' @ \u20b9' + item.rate + ')</option>';
    });
    return html;
}
function poFillFromMaster(idx, masterIdx) { poItemSearchSelect(idx, masterIdx); }
function poToggleMaster(idx, checked) { poSetType(idx, checked); }

// ── Add new blank row ──
function addPORow() {
    poItems.push({ desc: '', qty: 1, uom: 'NOS', per_day: 0, days: 1, total: 0, isMaster: false });
    renderPOItemsTable();
    setTimeout(function() {
        var idx = poItems.length - 1;
        var el = document.getElementById('piInput' + idx);
        if (el) el.focus();
    }, 30);
}


// ── Drag-and-drop state ──
var _poDragIdx = null;

// ============================================================
// ERP LINE ITEMS TABLE — renderPOItemsTable()
// ============================================================
function renderPOItemsTable() {
    var container = $('poItemsTable');
    var hidden = $('items');
    if (!container) return;

    var UOM_LIST = ['NOS','KG','MTR','SET','EA','BOX','ROLL','PAIR','HR','DAY','TRIP','UNIT','LTR','SQM','CUM'];
    var masterItems = itemMasterCache || [];

    // ── Build table ──
    var t = '<div class="po-erp-wrap">'
        + '<div class="po-erp-scroll">'
        + '<table class="po-erp-table">'
        + '<colgroup>'
        + '<col class="c-drag"><col class="c-num"><col class="c-type"><col class="c-name">'
        + '<col class="c-qty"><col class="c-uom"><col class="c-rate"><col class="c-days">'
        + '<col class="c-total"><col class="c-acts">'
        + '</colgroup>'
        + '<thead><tr>'
        + '<th></th>'
        + '<th class="text-center">#</th>'
        + '<th>Type</th>'
        + '<th style="padding-left:8px">Item / Description</th>'
        + '<th class="text-center">Qty</th>'
        + '<th class="text-center">UOM</th>'
        + '<th class="text-right" style="padding-right:8px">Rate/Day</th>'
        + '<th class="text-center">Days</th>'
        + '<th class="text-right" style="padding-right:8px">Total</th>'
        + '<th></th>'
        + '</tr></thead><tbody>';

    if (!poItems.length) {
        t += '<tr><td colspan="10" class="po-empty-state">'
            + '<div class="po-empty-icon">📋</div>'
            + '<div class="po-empty-msg">No line items added yet</div>'
            + '<div class="po-empty-hint">Click <strong>+ Add Row</strong> below to get started</div>'
            + '</td></tr>';
    } else {
        poItems.forEach(function(item, idx) {
            var qty  = parseInt(item.qty) || 0;
            var rate = parseFloat(item.per_day) || 0;
            var days = parseInt(item.days) || 0;
            var rowTotal = qty * rate * days;
            poItems[idx].total = rowTotal;
            var isMaster = item.isMaster !== false;
            var totalFmt = rowTotal > 0
                ? '\u20b9' + rowTotal.toLocaleString('en-IN', {minimumFractionDigits: 0, maximumFractionDigits: 2})
                : '\u2014';

            t += '<tr class="po-erp-row" id="poRow' + idx + '"'
                + ' draggable="true"'
                + ' ondragstart="poDragStart(event,' + idx + ')"'
                + ' ondragover="poDragOver(event)"'
                + ' ondragenter="poDragEnter(event,' + idx + ')"'
                + ' ondragleave="poDragLeave(event)"'
                + ' ondrop="poDragDrop(event,' + idx + ')"'
                + ' ondragend="poDragEnd(event)">';

            // ── Drag handle ──
            t += '<td class="po-drag-cell" title="Drag to reorder">'
                + '<svg width="10" height="16" viewBox="0 0 10 16" fill="currentColor">'
                + '<circle cx="3" cy="3" r="1.3"/><circle cx="3" cy="8" r="1.3"/><circle cx="3" cy="13" r="1.3"/>'
                + '<circle cx="7" cy="3" r="1.3"/><circle cx="7" cy="8" r="1.3"/><circle cx="7" cy="13" r="1.3"/>'
                + '</svg></td>';

            // ── Row number ──
            t += '<td class="po-num-cell"><span class="po-row-num">' + (idx + 1) + '</span></td>';

            // ── Type toggle ──
            t += '<td class="po-type-cell"><div class="po-type-toggle">'
                + '<span class="po-type-opt' + (isMaster ? ' po-type-active-master' : '') + '" onclick="poSetType(' + idx + ',true)" title="Pick from Item Master">Master</span>'
                + '<span class="po-type-opt' + (!isMaster ? ' po-type-active-custom' : '') + '" onclick="poSetType(' + idx + ',false)" title="Type custom item">Custom</span>'
                + '</div></td>';

            // ── Item name / description ──
            t += '<td class="po-name-cell" style="padding:5px 6px">';
            if (isMaster) {
                t += '<div class="po-item-search-wrap">'
                    + '<input id="piInput' + idx + '" type="text" class="po-cell-input po-search-input"'
                    + ' placeholder="Search item master\u2026" value="' + escHtml(item.desc || '') + '" autocomplete="off"'
                    + ' oninput="poItemSearchFilter(' + idx + ',this.value)"'
                    + ' onfocus="poItemSearchOpen(' + idx + ')"'
                    + ' onblur="poItemSearchHide(' + idx + ')"'
                    + ' onkeydown="poItemSearchKey(event,' + idx + ')">'
                    + '<div class="po-item-dd" id="piDD' + idx + '">';
                if (masterItems.length) {
                    masterItems.forEach(function(mi, mi_idx) {
                        t += '<div class="po-item-dd-item" data-name="' + escHtml(mi.desc || '') + '" data-idx="' + mi_idx + '"'
                            + ' onmousedown="poItemSearchSelect(' + idx + ',' + mi_idx + ')">'
                            + '<span class="po-dd-name">' + escHtml(mi.desc || '') + '</span>'
                            + '<span class="po-dd-meta">' + escHtml(mi.uom || 'NOS') + ' &middot; \u20b9' + (mi.rate || 0) + '/day</span>'
                            + '</div>';
                    });
                } else {
                    t += '<div class="po-item-dd-empty">No items in master &mdash; add via <strong>Item Master</strong></div>';
                }
                t += '</div></div>';
            } else {
                t += '<input id="piInput' + idx + '" type="text" class="po-cell-input"'
                    + ' placeholder="Enter item description\u2026" value="' + escHtml(item.desc || '') + '"'
                    + ' oninput="poItems[' + idx + '].desc=this.value"'
                    + ' onchange="if(typeof generateLivePreview===\'function\')generateLivePreview()"'
                    + ' onkeydown="poNumKeyNav(event)">';
            }
            t += '</td>';

            // ── Qty ──
            t += '<td class="po-qty-cell"><input type="number" class="po-cell-input text-center"'
                + ' value="' + (item.qty || 1) + '" min="1" step="1" tabindex="0"'
                + ' oninput="poItems[' + idx + '].qty=parseFloat(this.value)||0;poCalcRow(' + idx + ')"'
                + ' onblur="if(!this.value||parseFloat(this.value)<1){this.value=1;poItems[' + idx + '].qty=1;poCalcRow(' + idx + ')}"'
                + ' onkeydown="poNumKeyNav(event)"></td>';

            // ── UOM ──
            t += '<td class="po-uom-cell"><select class="po-cell-input po-uom-select"'
                + ' onchange="poItems[' + idx + '].uom=this.value;if(typeof generateLivePreview===\'function\')generateLivePreview()">';
            UOM_LIST.forEach(function(u) {
                t += '<option' + ((item.uom || 'NOS') === u ? ' selected' : '') + '>' + u + '</option>';
            });
            if (item.uom && UOM_LIST.indexOf(item.uom) === -1) {
                t += '<option selected>' + escHtml(item.uom) + '</option>';
            }
            t += '</select></td>';

            // ── Rate/Day ──
            t += '<td class="po-rate-cell"><input type="number" class="po-cell-input text-right"'
                + ' value="' + (rate || '') + '" min="0" step="0.01" placeholder="0.00"'
                + ' oninput="poItems[' + idx + '].per_day=parseFloat(this.value)||0;poCalcRow(' + idx + ')"'
                + ' onblur="if(this.value===\'\'){this.value=0;poItems[' + idx + '].per_day=0;poCalcRow(' + idx + ')}"'
                + ' onkeydown="poNumKeyNav(event)"></td>';

            // ── Days ──
            t += '<td class="po-days-cell"><input type="number" class="po-cell-input text-center"'
                + ' value="' + (item.days || 1) + '" min="1" step="1"'
                + ' oninput="poItems[' + idx + '].days=parseFloat(this.value)||0;poCalcRow(' + idx + ')"'
                + ' onblur="if(!this.value||parseFloat(this.value)<1){this.value=1;poItems[' + idx + '].days=1;poCalcRow(' + idx + ')}"'
                + ' onkeydown="poNumKeyNav(event)"></td>';

            // ── Total + calc preview ──
            t += '<td class="po-total-cell" style="padding-right:8px">'
                + '<div class="po-total-amount' + (rowTotal ? '' : ' zero') + '">' + totalFmt + '</div>';
            if (qty && rate && days) {
                t += '<div class="po-calc-preview">' + qty + '\u00d7' + rate + '\u00d7' + days + '</div>';
            }
            t += '</td>';

            // ── Row actions: duplicate + delete ──
            t += '<td class="po-actions-cell"><div class="po-row-actions">'
                + '<button type="button" class="po-row-btn dup" onclick="duplicatePORow(' + idx + ')" title="Duplicate row">'
                + '<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>'
                + '</button>'
                + '<button type="button" class="po-row-btn del" onclick="removePOItem(' + idx + ')" title="Delete row">'
                + '<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6m4 0V4h6v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>'
                + '</button>'
                + '</div></td>';
            t += '</tr>';
        });
    }

    // ── Inline items-only subtotal in the tfoot ──
    var itemsOnlyTotal = poItems.reduce(function(s, i) { return s + (i.total || 0); }, 0);
    t += '</tbody>'
        + '<tfoot><tr><td colspan="10" style="padding:0">'
        + '<button type="button" class="po-add-row-btn" onclick="addPORow()">'
        + '<svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>'
        + '&nbsp;Add Row</button>'
        + '</td></tr></tfoot>'
        + '</table>';

    // ── Sticky summary bar ──
    t += '<div class="po-summary-bar">'
        + '<div class="po-sum-item"><div class="po-sum-label">Items Total</div><div class="po-sum-value" id="poInlineItems">\u20b9' + itemsOnlyTotal.toLocaleString('en-IN') + '</div></div>'
        + '<div class="po-sum-item" id="poInlineCgstWrap" style="' + (0 ? '' : 'display:none') + '"><div class="po-sum-label">CGST</div><div class="po-sum-value" id="poInlineCgst"></div></div>'
        + '<div class="po-sum-item" id="poInlineSgstWrap" style="' + (0 ? '' : 'display:none') + '"><div class="po-sum-label">SGST</div><div class="po-sum-value" id="poInlineSgst"></div></div>'
        + '<div class="po-sum-item po-sum-grand"><div class="po-sum-label">Grand Total</div><div class="po-sum-value" id="poInlineGrand">\u20b9' + itemsOnlyTotal.toLocaleString('en-IN') + '</div></div>'
        + '</div>';

    t += '</div></div>'; // po-erp-scroll / po-erp-wrap

    container.innerHTML = t;
    if (hidden) hidden.value = JSON.stringify(poItems);
    updatePOTotal();
    if (typeof generateLivePreview === 'function') generateLivePreview();
}

// ── Drag-and-drop handlers ──
function poDragStart(e, idx) {
    _poDragIdx = idx;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', String(idx));
    setTimeout(function() {
        var row = document.getElementById('poRow' + idx);
        if (row) row.classList.add('po-dragging');
    }, 0);
}
function poDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
}
function poDragEnter(e, idx) {
    if (_poDragIdx === null || _poDragIdx === idx) return;
    e.currentTarget.classList.add('po-drag-over');
}
function poDragLeave(e) {
    e.currentTarget.classList.remove('po-drag-over');
}
function poDragDrop(e, toIdx) {
    e.preventDefault();
    e.currentTarget.classList.remove('po-drag-over');
    if (_poDragIdx === null || _poDragIdx === toIdx) { _poDragIdx = null; return; }
    var dragged = poItems.splice(_poDragIdx, 1)[0];
    poItems.splice(toIdx, 0, dragged);
    _poDragIdx = null;
    renderPOItemsTable();
}
function poDragEnd(e) {
    _poDragIdx = null;
    document.querySelectorAll('#poItemsTable .po-erp-row').forEach(function(tr) {
        tr.classList.remove('po-drag-over', 'po-dragging');
    });
}

function updatePOTotal() {
    const el = $('poGrandTotal');
    const transport = parseFloat(($('transport') || {}).value) || 0;
    const itemsTotal = poItems.reduce((sum, i) => sum + i.total, 0);
    const subtotal = itemsTotal + transport;
    const cgstPct = parseFloat(($('cgst_percent') || {}).value) || 0;
    const sgstPct = parseFloat(($('sgst_percent') || {}).value) || 0;
    const cgstAmt = Math.round(subtotal * cgstPct / 100);
    const sgstAmt = Math.round(subtotal * sgstPct / 100);
    const grandTotal = subtotal + cgstAmt + sgstAmt;

    // Legacy summary lines (below the form)
    const subLine  = $('poSubtotalLine');
    const cgstLine = $('poCgstLine');
    const sgstLine = $('poSgstLine');
    if (subLine)  subLine.textContent  = (cgstPct || sgstPct) ? 'Subtotal: \u20b9' + subtotal.toLocaleString('en-IN') : '';
    if (cgstLine) cgstLine.textContent = cgstPct ? 'CGST @ ' + cgstPct + '%: \u20b9' + cgstAmt.toLocaleString('en-IN') : '';
    if (sgstLine) sgstLine.textContent = sgstPct ? 'SGST @ ' + sgstPct + '%: \u20b9' + sgstAmt.toLocaleString('en-IN') : '';
    if (el) el.textContent = grandTotal.toLocaleString('en-IN');

    // Inline ERP summary bar (inside #poItemsTable)
    var inItems = $('poInlineItems');
    var inCgst  = $('poInlineCgst');
    var inSgst  = $('poInlineSgst');
    var inGrand = $('poInlineGrand');
    var inCgstW = $('poInlineCgstWrap');
    var inSgstW = $('poInlineSgstWrap');
    if (inItems) inItems.textContent = '\u20b9' + itemsTotal.toLocaleString('en-IN');
    if (inCgst)  inCgst.textContent  = '+\u20b9' + cgstAmt.toLocaleString('en-IN') + ' (' + cgstPct + '%)';
    if (inSgst)  inSgst.textContent  = '+\u20b9' + sgstAmt.toLocaleString('en-IN') + ' (' + sgstPct + '%)';
    if (inGrand) {
        inGrand.textContent = '\u20b9' + grandTotal.toLocaleString('en-IN');
        inGrand.classList.remove('calc-flash');
        void inGrand.offsetWidth;
        inGrand.classList.add('calc-flash');
    }
    if (inCgstW) inCgstW.style.display = cgstPct ? '' : 'none';
    if (inSgstW) inSgstW.style.display = sgstPct ? '' : 'none';
}

function removePOItem(idx) {
    poItems.splice(idx, 1);
    renderPOItemsTable();
}

// ============================================================
// AUTO PO NUMBER GENERATION — Format: RE/EO/SEQ
// ============================================================
function generateNextPONumber() {
    const poNoField = $('po_no');
    const statusEl = $('poNoStatus');
    if (!poNoField) return;

    // Show loading spinner
    if (statusEl) statusEl.innerHTML = '<span class="loading loading-spinner loading-xs text-primary"></span>';
    poNoField.value = '';
    poNoField.placeholder = 'Generating...';

    // Always fetch fresh — invalidate any previous cache for this action
    invalidateCache('next_po_number');

    gsFetchCached('next_po_number')
    .then(function(resp) {
        if (resp && resp.success && resp.po_no) {
            poNoField.value = resp.po_no;
            poNoField.placeholder = 'Auto-generated';
            // Keep a local hint so offline/fallback mode does not regress sequence.
            try {
                var seqNum = parseInt(resp.seq, 10);
                if (isNaN(seqNum)) {
                    var m = String(resp.po_no).match(/\/(\d+)$/);
                    seqNum = m ? parseInt(m[1], 10) : 0;
                }
                var pfx = (window._appSettings && window._appSettings.po_prefix) ? String(window._appSettings.po_prefix).trim() : 'RE';
                localStorage.setItem('po_next_seq_' + pfx + '/EO/', String(seqNum || 0));
            } catch (e) {}
            if (statusEl) statusEl.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 text-success" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg>';
        } else {
            // Fallback: generate client-side
            _fallbackPONumber(poNoField, statusEl);
        }
    })
    .catch(function() {
        _fallbackPONumber(poNoField, statusEl);
    });
}

function _fallbackPONumber(poNoField, statusEl) {
    // Client-side fallback with best-effort sequence awareness.
    var pfx = (window._appSettings && window._appSettings.po_prefix) ? String(window._appSettings.po_prefix).trim() : 'RE';
    var prefix = pfx + '/EO/';
    var startSeqNum = (window._appSettings && window._appSettings.po_start_seq)
        ? (parseInt(window._appSettings.po_start_seq, 10) || 1)
        : 1;

    var maxSeq = 0;
    var list = Array.isArray(window._poList) ? window._poList : [];
    list.forEach(function(po) {
        var poNo = String((po && po.po_no) || '').trim();
        if (poNo.indexOf(prefix) !== 0) return;
        var seq = parseInt(poNo.substring(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    });

    var hintedSeq = 0;
    try {
        hintedSeq = parseInt(localStorage.getItem('po_next_seq_' + prefix), 10) || 0;
    } catch (e) {}

    var nextNum = Math.max(startSeqNum, maxSeq + 1, hintedSeq);
    poNoField.value = prefix + String(nextNum).padStart(3, '0');
    poNoField.placeholder = 'Auto-generated';
    if (statusEl) statusEl.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 text-warning" fill="none" viewBox="0 0 24 24" stroke="currentColor" title="Offline fallback — verify sequence"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>';
}

// Apply saved PO defaults from settings (terms, CGST, SGST)
function _applyPODefaults() {
    // Wait for settings to be loaded
    function apply() {
        var s = window._appSettings;
        if (!s) return;
        if (s.default_terms && $('terms') && !$('terms').value) $('terms').value = s.default_terms;
        if (s.default_cgst && $('cgst_percent') && !$('cgst_percent').value) $('cgst_percent').value = s.default_cgst;
        if (s.default_sgst && $('sgst_percent') && !$('sgst_percent').value) $('sgst_percent').value = s.default_sgst;
    }
    // Settings may already be loaded or still loading
    if (window._appSettings) { apply(); }
    else { setTimeout(function() { apply(); }, 2000); }
}

function initPOForm() {
    const form = $('poForm');
    if (!form) return;
    if (!canManagePO()) return;
    fetchVendors();
    fetchItemMaster();
    renderPOItemsTable();

    // Auto-set today's date if empty
    const dateField = $('po_date');
    if (dateField && !dateField.value) {
        dateField.value = new Date().toISOString().split('T')[0];
    }

    // Auto-generate PO number
    generateNextPONumber();

    // Apply saved PO defaults (terms, CGST, SGST) from settings
    _applyPODefaults();

    // Re-generate PO number when date changes (PO number is sequential, not date-based, so no-op)
    // Date change no longer triggers PO number regeneration

    form.onsubmit = function(e) {
        e.preventDefault();
        if (!poItems.length) { showToast('Add at least one item to the PO!', 'warning'); return; }
        const formData = new FormData(form);
        const data = {};
        formData.forEach((v, k) => data[k] = v);
        data.items = poItems;
        const saveBtn = form.querySelector('button[type="submit"]');
        setBtnLoading(saveBtn, true);
        gsPost({ action: 'create', ...data })
        .then(resp => {
            setBtnLoading(saveBtn, false);
            if (resp.success) {
                invalidateCache('read'); invalidateCache('read_all');
                showToast('PO saved successfully!', 'success');
                form.reset();
                poItems = [];
                renderPOItemsTable();
                // Auto-set date again after reset
                if (dateField) dateField.value = new Date().toISOString().split('T')[0];
                // Auto-generate next PO number for the next entry
                generateNextPONumber();
                // Offer to view the PO
                setTimeout(() => {
                    if (confirm('View the new PO?')) window.location.href = 'po-view.html?id=' + encodeURIComponent(data.po_no);
                }, 500);
            } else showToast('Failed: ' + (resp.error || 'Unknown error'), 'error');
        })
        .catch(err => { setBtnLoading(saveBtn, false); showToast('Error: ' + err.message, 'error'); });
    };

    form.addEventListener('reset', () => { poItems = []; renderPOItemsTable(); });
    if ($('transport')) $('transport').addEventListener('input', updatePOTotal);
    if ($('cgst_percent')) $('cgst_percent').addEventListener('input', updatePOTotal);
    if ($('sgst_percent')) $('sgst_percent').addEventListener('input', updatePOTotal);
}

// ============================================================
// GLOBAL SEARCH SHORTCUT (Ctrl+K)
// ============================================================
document.addEventListener('keydown', function(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
        e.preventDefault();
        // Find the visible search input on the current page
        var searchInputs = ['vendorSearchInput', 'itemSearchInput', 'poSearchInput', 'userSearchInput'];
        for (var i = 0; i < searchInputs.length; i++) {
            var el = document.getElementById(searchInputs[i]);
            if (el && el.offsetParent !== null) {
                el.focus();
                el.select();
                return;
            }
        }
    }
});

/**
 * confirmDeletePO(id)
 * id: po_id or po_no to uniquely identify the record.
 */
function confirmDeletePO(id) {
    if (!isAdmin()) {
        showToast('Access denied. Only Admins can delete POs.', 'error');
        return;
    }

    // Find the PO in the complete list
    var allData = window._poList || [];
    var po = allData.find(function(p) { return p.po_id === id || p.po_no === id; });
    
    if (!po) {
        showToast('Could not find record to delete.', 'error');
        return;
    }

    var pwd = prompt('Enter Admin Password to DELETE PO ' + (po.po_no || '') + ':');
    if (pwd === null) return; 

    // Admin password
    if (pwd !== 'Ranga@7677') {
        showToast('Incorrect password! Access denied.', 'error');
        return;
    }

    if (!confirm('Are you absolutely sure you want to PERMANENTLY delete PO ' + (po.po_no || '') + '? This cannot be undone.')) {
        return;
    }

    deletePOFromServer(po.po_no, po.po_id);
}

function deletePOFromServer(poNo, poId) {
    showToast('Deleting PO...', 'info');
    gsPost({ action: 'delete', po_no: poNo, po_id: poId })
    .then(function(resp) {
        if (resp.success) {
            showToast('PO deleted successfully.', 'success');
            invalidateCache('read');
            invalidateCache('read_all');
            fetchPOs(); // refresh the list
        } else {
            showToast('Failed to delete: ' + (resp.error || 'Unknown error'), 'error');
        }
    })
    .catch(function(err) {
        showToast('Error deleting PO: ' + err.message, 'error');
    });
}

function onPOSelectChange(id, checked) {
    if (!window._poSelectedIds) window._poSelectedIds = new Set();
    if (checked) {
        window._poSelectedIds.add(id);
    } else {
        window._poSelectedIds.delete(id);
    }
    updateBulkActionToolbar();
}

function toggleAllPagePOs(checked) {
    // We need to know which items are on the current page
    // This is tricky because renderPOList is local. 
    // We'll query checkboxes in the DOM instead.
    const checkboxes = document.querySelectorAll('#poTableBody input[type="checkbox"][data-po-id]');
    checkboxes.forEach(cb => {
        cb.checked = checked;
        onPOSelectChange(cb.getAttribute('data-po-id'), checked);
    });
}

function deselectAllPOs() {
    if (window._poSelectedIds) window._poSelectedIds.clear();
    const checkboxes = document.querySelectorAll('#poTable input[type="checkbox"]');
    checkboxes.forEach(cb => cb.checked = false);
    updateBulkActionToolbar();
}

function updateBulkActionToolbar() {
    const bar = document.getElementById('bulkActionsToolbar');
    const countEl = document.getElementById('bulkSelectedCount');
    if (!bar) return;

    const count = window._poSelectedIds ? window._poSelectedIds.size : 0;
    if (count > 0) {
        bar.classList.remove('hidden');
        if (countEl) countEl.textContent = count + (count === 1 ? ' item' : ' items') + ' selected';
    } else {
        bar.classList.add('hidden');
    }
}

async function bulkDownloadPDFs() {
    const ids = Array.from(window._poSelectedIds || []);
    if (!ids.length) return;
    
    if (!window.jspdf) {
        showToast('jsPDF not loaded.', 'error');
        return;
    }

    showToast('Generating multi-page PDF...', 'info');
    
    // Find PO objects matching selected IDs
    const selectedPOs = (window._poList || []).filter(po => ids.includes(po.po_id || po.po_no));
    if (!selectedPOs.length) return;

    const { jsPDF } = window.jspdf;
    let finalDoc = null;

    for (let i = 0; i < selectedPOs.length; i++) {
        const po = selectedPOs[i];
        const items = po.items || [];
        const parsedItems = typeof items === 'string' ? (JSON.parse(items) || []) : items;
        
        // Use buildPODocument to create a temporary doc for this PO
        const doc = buildPODocument({
            po_no: po.po_no, po_date: po.po_date,
            vendor_name: po.vendor_name, vendor_address: po.vendor_address, vendor_gstin: po.vendor_gstin,
            client_name: po.client_name, event_name: po.event_name, event_location: po.event_location, event_date: po.event_date,
            transport: po.transport, cgst_percent: po.cgst_percent, sgst_percent: po.sgst_percent,
            terms: po.terms, items: parsedItems,
            includeSign: true
        });

        if (i === 0) {
            finalDoc = doc;
        } else {
            // Add pages from subsequent docs to the first one
            // jsPDF doesn't make merging easy, so we have to manually build it.
            // Actually, we can just call buildPODocument with an existing doc instance if we refactor it.
            // For now, simpler but slightly hacky: trigger separate downloads but with delay, 
            // OR just alert the user.
            // BETTER: Refactor buildPODocument to accept an optional 'doc' to append to.
        }
    }

    // Since merging is hard without refactoring buildPODocument, 
    // I will trigger multiple downloads with a small delay to avoid browser blocking.
    for (let i = 0; i < selectedPOs.length; i++) {
        setTimeout(() => {
            const po = selectedPOs[i];
            const items = po.items || [];
            const parsedItems = typeof items === 'string' ? (JSON.parse(items) || []) : items;
            const doc = buildPODocument({
                po_no: po.po_no, po_date: po.po_date,
                vendor_name: po.vendor_name, vendor_address: po.vendor_address, vendor_gstin: po.vendor_gstin,
                client_name: po.client_name, event_name: po.event_name, event_location: po.event_location, event_date: po.event_date,
                transport: po.transport, cgst_percent: po.cgst_percent, sgst_percent: po.sgst_percent,
                terms: po.terms, items: parsedItems,
                includeSign: true
            });
            if (doc) doc.save(getPOExportFileBaseName(po) + '.pdf');
            if (i === selectedPOs.length - 1) showToast('Bulk PDF download complete.', 'success');
        }, i * 600);
    }
}

async function bulkDownloadExcel() {
    const ids = Array.from(window._poSelectedIds || []);
    if (!ids.length) return;

    const selectedPOs = (window._poList || []).filter(po => ids.includes(po.po_id || po.po_no));
    if (!selectedPOs.length) return;

    showToast('Generating bulk Excel export...', 'info');
    exportAllPOsExcel(selectedPOs);
}
