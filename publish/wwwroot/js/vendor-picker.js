/* ============================================================================
   vendor-picker.js — interaction for _VendorPicker.cshtml.
   Every picker on the page is keyed by its id prefix (e.g. "mk", "animal"),
   so multiple pickers can coexist without colliding.
   ========================================================================== */
(function () {
    'use strict';

    // Per-picker snapshot (used for Cancel to restore).
    var snapshots = {};

    function $(id) { return document.getElementById(id); }
    function chks(id) { return document.querySelectorAll('.' + id + '-chk'); }

    function getSelectedNames(id) {
        var out = [];
        chks(id).forEach(function (c) {
            if (c.checked) out.push({ id: c.value, name: c.getAttribute('data-name') || '' });
        });
        return out;
    }

    function updateItemHighlights(id) {
        chks(id).forEach(function (c) {
            var label = c.closest('.vp-item');
            if (!label) return;
            label.classList.toggle('vp-sel', c.checked);
        });
    }

    function renderChips(id) {
        var host = $(id + '-chips');
        if (!host) return;
        var sel = getSelectedNames(id);
        host.innerHTML = '';
        if (sel.length === 0) {
            var empty = document.createElement('span');
            empty.className = 'vp-chips-empty';
            empty.textContent = 'No vendors selected';
            host.appendChild(empty);
        } else {
            sel.forEach(function (v) {
                var chip = document.createElement('span');
                chip.className = 'vp-chip';
                chip.appendChild(document.createTextNode(v.name));
                var x = document.createElement('span');
                x.className = 'vp-chip-x';
                x.textContent = '×';
                x.title = 'Remove ' + v.name;
                x.addEventListener('click', function (e) {
                    e.stopPropagation();
                    deselect(id, v.id);
                });
                chip.appendChild(x);
                host.appendChild(chip);
            });
        }
        var cnt = $(id + '-count');
        if (cnt) cnt.textContent = sel.length + ' selected';
        updateItemHighlights(id);
    }

    function deselect(id, vendorId) {
        chks(id).forEach(function (c) {
            if (c.value === vendorId) c.checked = false;
        });
        renderChips(id);
    }

    window.VP_toggle = function (id) {
        var panel   = $(id + '-panel');
        var trigger = $(id + '-trigger');
        if (!panel) return;
        var willOpen = !panel.classList.contains('vp-open');

        // Only one picker open at a time
        document.querySelectorAll('.vp-panel.vp-open').forEach(function (p) {
            if (p !== panel) p.classList.remove('vp-open');
        });

        if (willOpen) {
            snapshots[id] = Array.from(chks(id)).map(function (c) {
                return { v: c.value, on: c.checked };
            });
            panel.classList.add('vp-open');
            trigger.setAttribute('aria-expanded', 'true');
            renderChips(id);
            setTimeout(function () { var s = $(id + '-search'); if (s) s.focus(); }, 10);
        } else {
            panel.classList.remove('vp-open');
            trigger.setAttribute('aria-expanded', 'false');
        }
    };

    window.VP_filter = function (id, q) {
        var qq = (q || '').toLowerCase().trim();
        var grid = $(id + '-grid');
        if (!grid) return;
        grid.querySelectorAll('.vp-item').forEach(function (lbl) {
            var name = (lbl.getAttribute('data-vendor-name') || '').toLowerCase();
            lbl.style.display = (!qq || name.indexOf(qq) !== -1) ? '' : 'none';
        });
    };

    window.VP_onChange = function (id) {
        renderChips(id);
    };

    window.VP_selectAll = function (id, val) {
        chks(id).forEach(function (c) {
            // Toggle only visible items so the user can Select-all within a search.
            var lbl = c.closest('.vp-item');
            if (lbl && lbl.style.display === 'none') return;
            c.checked = val;
        });
        renderChips(id);
    };

    window.VP_cancel = function (id) {
        var snap = snapshots[id];
        if (snap) {
            var map = {};
            snap.forEach(function (s) { map[s.v] = s.on; });
            chks(id).forEach(function (c) { c.checked = !!map[c.value]; });
            renderChips(id);
        }
        var panel   = $(id + '-panel');
        var trigger = $(id + '-trigger');
        if (panel)   panel.classList.remove('vp-open');
        if (trigger) trigger.setAttribute('aria-expanded', 'false');
    };

    window.VP_apply = function (id) {
        var sel     = getSelectedNames(id);
        var ids     = sel.map(function (v) { return v.id; }).join(',');
        var hidden  = $(id + '-hidden');
        if (hidden) hidden.value = ids;

        var total    = chks(id).length;
        var labelEl  = $(id + '-label');
        if (labelEl) {
            if (sel.length === 0)            labelEl.textContent = 'All vendors (' + (total || 0) + ')';
            else if (sel.length === total)   labelEl.textContent = 'All vendors selected';
            else                             labelEl.textContent = sel.length + ' of ' + total + ' vendors selected';
        }

        var panel   = $(id + '-panel');
        var trigger = $(id + '-trigger');
        if (panel)   panel.classList.remove('vp-open');
        if (trigger) trigger.setAttribute('aria-expanded', 'false');

        // Fire a DOM event so page scripts can hook in (e.g. auto-submit form).
        var ev = new CustomEvent('vp:apply', {
            detail: { id: id, vendorIds: ids, count: sel.length }
        });
        document.dispatchEvent(ev);
    };

    // Outside click closes
    document.addEventListener('mousedown', function (e) {
        if (e.target.closest('.vp-panel') || e.target.closest('.vp-trigger')) return;
        document.querySelectorAll('.vp-panel.vp-open').forEach(function (p) {
            p.classList.remove('vp-open');
            var wrap = p.closest('.vp');
            var trig = wrap && wrap.querySelector('.vp-trigger');
            if (trig) trig.setAttribute('aria-expanded', 'false');
        });
    });

    // Escape closes the open picker
    document.addEventListener('keydown', function (e) {
        if (e.key !== 'Escape') return;
        var open = document.querySelector('.vp-panel.vp-open');
        if (!open) return;
        var wrap = open.closest('.vp');
        if (!wrap) return;
        var id = wrap.id.replace(/-vp$/, '');
        window.VP_cancel(id);
    });
}());
