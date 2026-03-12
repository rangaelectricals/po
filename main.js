// ============================================================
// main.js - PO Manager - Shared Logic
// Replace 'YOUR_GOOGLE_SCRIPT_URL' with your deployed Google Apps Script Web App URL
// ============================================================
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx4V3wD2cM8dF2tz0mD2DJUFyJLLAfji-7yHJKNuqNHLZP30ylkmBJG_gB3serhxNg/exec';

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
    if (isAdmin()) return; // Admin sees everything

    // Hide all elements marked as admin-only
    document.querySelectorAll('[data-admin-only]').forEach(function(el) {
        el.style.display = 'none';
    });

    // Hide elements with class .admin-only
    document.querySelectorAll('.admin-only').forEach(function(el) {
        el.style.display = 'none';
    });

    // Disable all form submit/add/delete buttons (non-admin = view only)
    document.querySelectorAll('.crud-action').forEach(function(el) {
        el.style.display = 'none';
    });

    // Hide "Create PO" sidebar link for non-admin
    document.querySelectorAll('.menu a').forEach(function(a) {
        if (a.getAttribute('href') && (a.getAttribute('href').indexOf('add-po') !== -1)) {
            a.closest('li').style.display = 'none';
        }
    });
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
    var menus = document.querySelectorAll('.menu');
    menus.forEach(function(menu) {
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

// Master init for auth on every page — call from window.onload
function initPageAuth() {
    if (!requireAuth()) return false;
    injectNavbarAuth();
    injectSidebarAuthLinks();
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

    var start = (pg.page - 1) * pg.size + 1;
    var end = Math.min(pg.page * pg.size, totalItems);

    var html = '<div class="flex flex-wrap items-center justify-between gap-3 px-4 py-3 border-t border-base-300 bg-base-100 text-sm">';

    // Left: showing info + page size selector
    html += '<div class="flex items-center gap-3">';
    html += '<span class="text-base-content/60">Showing <b>' + start + '</b>–<b>' + end + '</b> of <b>' + totalItems + '</b></span>';
    html += '<select class="select select-bordered select-xs w-auto" onchange="_pgSize(\'' + key + '\',' + 'this.value,' + renderFn + ')">';
    PAGE_SIZE_OPTIONS.forEach(function(s) {
        html += '<option value="' + s + '"' + (pg.size === s ? ' selected' : '') + '>' + s + ' / page</option>';
    });
    html += '</select>';
    html += '</div>';

    // Right: page buttons
    html += '<div class="join">';
    html += '<button class="join-item btn btn-xs' + (pg.page <= 1 ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\',' + 1 + ',' + renderFn + ')" title="First">&laquo;</button>';
    html += '<button class="join-item btn btn-xs' + (pg.page <= 1 ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\',' + (pg.page - 1) + ',' + renderFn + ')" title="Previous">&lsaquo;</button>';

    // Page number buttons (show max 5)
    var startP = Math.max(1, pg.page - 2);
    var endP = Math.min(totalPages, startP + 4);
    if (endP - startP < 4) startP = Math.max(1, endP - 4);
    for (var p = startP; p <= endP; p++) {
        html += '<button class="join-item btn btn-xs' + (p === pg.page ? ' btn-active btn-primary' : '') + '" onclick="_pgGo(\'' + key + '\',' + p + ',' + renderFn + ')">' + p + '</button>';
    }

    html += '<button class="join-item btn btn-xs' + (pg.page >= totalPages ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\',' + (pg.page + 1) + ',' + renderFn + ')" title="Next">&rsaquo;</button>';
    html += '<button class="join-item btn btn-xs' + (pg.page >= totalPages ? ' btn-disabled' : '') + '" onclick="_pgGo(\'' + key + '\',' + totalPages + ',' + renderFn + ')" title="Last">&raquo;</button>';
    html += '</div>';

    html += '</div>';
    return html;
}

function _pgGo(key, page, renderFn) {
    _pagination[key].page = page;
    renderFn();
}

function _pgSize(key, size, renderFn) {
    _pagination[key].size = parseInt(size);
    _pagination[key].page = 1;
    renderFn();
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
    var transport = parseFloat(d.transport) || 0;
    var itemsTotal = items.reduce(function(s, i) { return s + (parseFloat(i.total) || 0); }, 0);
    var subtotal = itemsTotal + transport;
    var cgstPct = parseFloat(d.cgst_percent) || 0;
    var sgstPct = parseFloat(d.sgst_percent) || 0;
    var cgstAmt = Math.round(subtotal * cgstPct / 100);
    var sgstAmt = Math.round(subtotal * sgstPct / 100);
    var grandTotal = subtotal + cgstAmt + sgstAmt;
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
    doc.text('E-mail: info@rangaelectricals.com', pageW / 2, y, { align: 'center' }); y += 4;
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
    doc.text(d.po_date || '____', rightBoxX + boxW - 4, ry, { align: 'right' }); ry += 6;
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
        doc.text(d.event_date, rightBoxX + 28, ry);
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

    let html = '<table class="table table-zebra w-full"><thead><tr><th>#</th><th>Vendor Name</th><th>Address</th><th>GSTIN</th>' + (isAdmin() ? '<th>Action</th>' : '') + '</tr></thead><tbody>';
    pageData.forEach((vendor, i) => {
        var displayIdx = startIdx + i;
        var origIdx = allData.indexOf(vendor);
        html += '<tr><td class="font-medium text-slate-500">' + (displayIdx + 1) + '</td><td class="font-semibold text-slate-800">' + (vendor.name || '') + '</td><td class="max-w-xs truncate text-slate-600">' + (vendor.address || '-') + '</td><td><code class="text-xs bg-slate-100 px-1.5 py-0.5 rounded">' + (vendor.gstin || '-') + '</code></td>' + (isAdmin() ? '<td><button class="btn btn-xs btn-error btn-outline btn-modern" onclick="deleteVendor(' + origIdx + ')">Delete</button></td>' : '') + '</tr>';
    });
    html += '</tbody></table>';
    if (total > pg.size) html += _buildPaginationBar('vendor', total, renderVendorList);
    $('vendorList').innerHTML = html;
}

function addVendor() {
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

function deleteVendor(idx) {
    if (!confirm('Delete this vendor?')) return;
    gsPost({ action: 'delete_vendor', idx })
    .then(resp => { if (resp.success) { invalidateCache('read_vendors'); invalidateCache('read_all'); fetchVendors(); } else alert('Failed to delete vendor: ' + (resp.error || 'Unknown error')); })
    .catch(err => alert('Error deleting vendor: ' + err.message));
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

    let html = '<table class="table table-zebra w-full"><thead><tr><th>#</th><th>Description</th><th>UOM</th><th>Rate/Day</th><th>Default Qty</th>' + (isAdmin() ? '<th>Action</th>' : '') + '</tr></thead><tbody>';
    pageData.forEach((item, i) => {
        var displayIdx = startIdx + i;
        var origIdx = allData.indexOf(item);
        html += '<tr><td class="font-medium text-slate-500">' + (displayIdx + 1) + '</td><td class="font-semibold text-slate-800">' + item.desc + '</td><td><span class="badge badge-ghost badge-sm">' + item.uom + '</span></td><td class="font-mono text-slate-700">' + item.rate + '</td><td class="text-slate-600">' + (item.default_qty || '') + '</td>' + (isAdmin() ? '<td><button class="btn btn-xs btn-error btn-outline btn-modern" onclick="deleteItemMaster(' + origIdx + ')">Delete</button></td>' : '') + '</tr>';
    });
    html += '</tbody></table>';
    if (total > pg.size) html += _buildPaginationBar('item', total, renderItemMasterList);
    $('itemMasterList').innerHTML = html;
}

function addItemMaster() {
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

function deleteItemMaster(idx) {
    if (!confirm('Delete this item?')) return;
    gsPost({ action: 'delete_item', idx })
    .then(resp => { if (resp.success) { invalidateCache('read_items'); invalidateCache('read_all'); fetchItemMaster(); } else alert('Failed to delete item: ' + (resp.error || 'Unknown error')); })
    .catch(err => alert('Error deleting item: ' + err.message));
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
function filterPOList(term) {
    _poSearchTerm = (term || '').toLowerCase().trim();
    _pagination['po'] = { page: 1, size: (_pagination['po'] || {}).size || 10 };
    renderPOList();
}
function renderPOList(data) {
    if (data) window._poList = data;
    var allData = window._poList || [];
    // Apply search filter
    var filtered = allData;
    if (_poSearchTerm) {
        filtered = allData.filter(function(po) {
            return (po.po_no || '').toLowerCase().includes(_poSearchTerm)
                || (po.vendor_name || '').toLowerCase().includes(_poSearchTerm)
                || (po.event_name || '').toLowerCase().includes(_poSearchTerm)
                || (po.event_location || '').toLowerCase().includes(_poSearchTerm)
                || (po.po_date || '').toLowerCase().includes(_poSearchTerm)
                || String(po.total || '').includes(_poSearchTerm);
        });
    }
    if (!Array.isArray(allData) || allData.length === 0) {
        $('poList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" /></svg><p class="text-sm">No purchase orders found.</p></div>';
        return;
    }
    if (filtered.length === 0 && _poSearchTerm) {
        $('poList').innerHTML = '<div class="p-6 text-center text-base-content/50"><svg xmlns="http://www.w3.org/2000/svg" class="h-8 w-8 mx-auto mb-2 opacity-30" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg><p class="text-sm">No POs match "<strong>' + escHtml(_poSearchTerm) + '</strong>"</p></div>';
        return;
    }
    // Hero stats
    if ($('poHeroStats')) {
        var statsHtml = '<div class="hero-stat"><div class="stat-label">Total POs</div><div class="stat-value">' + allData.length + '</div></div>';
        if (_poSearchTerm) statsHtml += '<div class="hero-stat"><div class="stat-label">Matched</div><div class="stat-value">' + filtered.length + '</div></div>';
        $('poHeroStats').innerHTML = statsHtml;
    }
    var pg = _initPagination('po', 10);
    var total = filtered.length;
    var startIdx = (pg.page - 1) * pg.size;
    var pageData = _pgSlice('po', filtered);

    let html = '<table class="table table-zebra w-full"><thead><tr><th>#</th><th>PO No</th><th>Date</th><th>Vendor</th><th>Event</th><th>Location</th><th>Total</th><th>Actions</th></tr></thead><tbody>';
    pageData.forEach((po, i) => {
        var displayIdx = startIdx + i;
        var idx = allData.indexOf(po);
        html += '<tr>';
        html += '<td class="font-medium text-slate-500">' + (displayIdx + 1) + '</td>';
        html += '<td><span class="font-semibold text-primary">' + (po.po_no || '') + '</span></td>';
        html += '<td class="text-slate-600">' + (po.po_date || '') + '</td>';
        html += '<td class="font-medium">' + (po.vendor_name || '') + '</td>';
        html += '<td class="text-slate-600">' + (po.event_name || '') + '</td>';
        html += '<td class="text-slate-500 text-xs">' + (po.event_location || '') + '</td>';
        html += '<td><span class="font-mono font-semibold text-emerald-600">' + (po.total ? '₹' + Number(po.total).toLocaleString('en-IN') : '') + '</span></td>';
        html += '<td>';
        html += '<div class="flex flex-wrap gap-1">';
        html += '<a href="po-view.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-primary btn-outline btn-modern" title="View"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg></a>';
        if (isAdmin()) {
        html += '<a href="po-edit.html?id=' + encodeURIComponent(String(po.po_no || '').trim()) + '" class="btn btn-xs btn-warning btn-outline btn-modern" title="Edit"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg></a>';
        }
        html += '<button onclick="listDownloadPDF(' + idx + ')" class="btn btn-xs btn-error btn-outline btn-modern" title="Download PDF"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg></button>';
        html += '<button onclick="listDownloadExcel(' + idx + ')" class="btn btn-xs btn-success btn-outline btn-modern" title="Download Excel"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"/></svg></button>';
        html += '<button onclick="listPrintPDF(' + idx + ')" class="btn btn-xs btn-info btn-outline btn-modern" title="Print PDF"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg></button>';
        html += '<button onclick="showPOItemsModal(' + idx + ')" class="btn btn-xs btn-neutral btn-outline btn-modern" title="Show Items"><svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 10h16M4 14h10"/></svg> Items</button>';
        html += '</div>';
        html += '</td>';
        html += '</tr>';
    });
    html += '</tbody></table>';
    if (total > pg.size) html += _buildPaginationBar('po', total, renderPOList);
    $('poList').innerHTML = html;
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
    if (!items.length && !transport) {
        html = '<div class="py-10 text-center text-base-content/40 text-sm">No line items recorded for this PO.</div>';
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

        // Transport row
        if (transport > 0) {
            html += '<tr class="italic text-base-content/60">';
            html += '<td class="text-center">+</td>';
            html += '<td colspan="5">Transport Charges</td>';
            html += '<td class="text-right font-mono">₹' + transport.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</td>';
            html += '</tr>';
        }

        // Footer subtotal row
        html += '<tr class="font-bold bg-base-200">';
        html += '<td></td>';
        html += '<td>Total</td>';
        html += '<td class="text-right font-mono">' + totalQty + '</td>';
        html += '<td colspan="3"></td>';
        html += '<td class="text-right font-mono">₹' + subtotal.toLocaleString('en-IN', { minimumFractionDigits: 2 }) + '</td>';
        html += '</tr>';
        html += '</tbody></table>';

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
    if (doc) doc.save('PO_' + (po.po_no || 'unknown').replace(/[\/\\]/g, '_') + '.pdf');
}

// Download Excel for a PO from the list page
function listDownloadExcel(idx) {
    var po = _preparePO(idx);
    if (!po) return alert('PO not found');
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
    XLSX.writeFile(wb, 'PO_' + (po.po_no || 'unknown').replace(/[\/\\]/g, '_') + '.xlsx');
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
    if (inGrand) inGrand.textContent = '\u20b9' + grandTotal.toLocaleString('en-IN');
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
    // Client-side fallback: PREFIX/EO/001 (no sequence awareness)
    var pfx = (window._appSettings && window._appSettings.po_prefix) ? window._appSettings.po_prefix : 'RE';
    var startSeq = (window._appSettings && window._appSettings.po_start_seq) ? String(parseInt(window._appSettings.po_start_seq) || 1).padStart(3, '0') : '001';
    poNoField.value = pfx + '/EO/' + startSeq;
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
