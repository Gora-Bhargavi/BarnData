/* ============================================================================
   column-filter.js — Access / Excel-style column header filter. (v2)
   
   v2 changes vs v1:
     - Whole TH is clickable, not just the tiny funnel icon. Huge hit target.
     - Funnel icon is fully opaque and larger so it reads on dark headers.
     - Scroll inside the table no longer auto-closes the popover.
     - Popover hardened against sticky-header stacking and clip regions.
   
   Drop-in usage:
     <table data-colfilter-table>
       <thead><tr>
         <th data-colfilter="text">Vendor</th>
         <th data-colfilter="number" class="text-right">Live Wt</th>
         <th data-colfilter="date">Purchase Date</th>
         <th>Actions</th>                        <!-- no filter -->
       </tr></thead>
       <tbody> ... </tbody>
     </table>
   
   After the table is in the DOM, auto-install runs on DOMContentLoaded.
   For tables added dynamically: ColumnFilter.install(tableEl)
   ========================================================================== */
(function (global) {
    'use strict';

    var FUNNEL_SVG =
        '<svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true" focusable="false">' +
        '<path d="M2 3h12l-4.5 6v4l-3 1.5V9L2 3z"/></svg>';

    // Popover instance (only one at a time on the page)
    var _activePop = null;

    function closePopovers() {
        if (_activePop && _activePop.parentNode) {
            _activePop.parentNode.removeChild(_activePop);
        }
        _activePop = null;
    }

    function getCellText(row, colIndex) {
        var cell = row.cells[colIndex];
        if (!cell) return '';
        var input = cell.querySelector('input[type="text"], input[type="number"], input:not([type]), select, textarea');
        if (input) {
            if (input.tagName === 'SELECT') {
                return (input.options[input.selectedIndex] && input.options[input.selectedIndex].text || '').trim();
            }
            return (input.value || '').trim();
        }
        return (cell.textContent || '').trim();
    }

    function uniqueSorted(values, kind) {
        var set = new Set();
        values.forEach(function (v) { set.add(v); });
        var arr = Array.from(set);
        arr.sort(function (a, b) {
            if (kind === 'number') {
                var an = parseFloat(String(a).replace(/,/g, ''));
                var bn = parseFloat(String(b).replace(/,/g, ''));
                var aNaN = isNaN(an), bNaN = isNaN(bn);
                if (aNaN && bNaN) return String(a).localeCompare(String(b));
                if (aNaN) return 1;
                if (bNaN) return -1;
                return an - bn;
            }
            if (kind === 'date') {
                var ad = Date.parse(a), bd = Date.parse(b);
                if (isNaN(ad) && isNaN(bd)) return String(a).localeCompare(String(b));
                if (isNaN(ad)) return 1;
                if (isNaN(bd)) return -1;
                return ad - bd;
            }
            return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
        });
        return arr;
    }

    function buildPopover(th, table) {
        closePopovers();

        var kind      = th.getAttribute('data-colfilter') || 'text';
        var colIndex  = th.cellIndex;
        var rect      = th.getBoundingClientRect();

        var rows = table.tBodies[0] ? Array.from(table.tBodies[0].rows) : [];
        var allValues = rows.map(function (r) { return getCellText(r, colIndex); });
        var distinct = uniqueSorted(allValues, kind);

        var prevState = th._cfState;
        var selectedSet = (prevState && prevState.selected)
            ? new Set(prevState.selected)
            : new Set(distinct);

        var pop = document.createElement('div');
        pop.className = 'cf-pop';
        // Anchor below the header, clamped to viewport.
        pop.style.top  = Math.max(4, rect.bottom + 2) + 'px';
        pop.style.left = Math.min(window.innerWidth - 278, Math.max(8, rect.left)) + 'px';

        var sortAscLabel = kind === 'number' ? 'Sort 0 → 9'
                         : kind === 'date'   ? 'Sort oldest → newest'
                         :                     'Sort A → Z';
        var sortDescLabel = kind === 'number' ? 'Sort 9 → 0'
                          : kind === 'date'   ? 'Sort newest → oldest'
                          :                     'Sort Z → A';

        pop.innerHTML =
            '<div class="cf-sort" role="group" aria-label="Sort">' +
              '<div data-sort="asc"  class="cf-sort-item" tabindex="0">↑&nbsp;&nbsp;' + sortAscLabel  + '</div>' +
              '<div data-sort="desc" class="cf-sort-item" tabindex="0">↓&nbsp;&nbsp;' + sortDescLabel + '</div>' +
            '</div>' +
            '<div class="cf-search-wrap">' +
              '<input class="cf-search" type="text" placeholder="Search values…" aria-label="Search values" />' +
            '</div>' +
            '<div class="cf-list" role="listbox"></div>' +
            '<div class="cf-foot">' +
              '<button class="cf-btn cf-btn-clear" type="button">Clear filter</button>' +
              '<span class="cf-foot-spacer"></span>' +
              '<button class="cf-btn cf-btn-cancel" type="button">Cancel</button>' +
              '<button class="cf-btn cf-btn-primary" type="button">Apply</button>' +
            '</div>';

        document.body.appendChild(pop);
        _activePop = pop;

        var listEl   = pop.querySelector('.cf-list');
        var searchEl = pop.querySelector('.cf-search');

        function renderList() {
            var q = (searchEl.value || '').toLowerCase().trim();
            var visible = distinct.filter(function (v) { return !q || v.toLowerCase().indexOf(q) !== -1; });
            var allChecked = visible.length > 0 && visible.every(function (v) { return selectedSet.has(v); });

            var frag = document.createDocumentFragment();

            var allLbl = document.createElement('label');
            allLbl.className = 'cf-item cf-item-all';
            allLbl.innerHTML = '<input type="checkbox" class="cf-all"' + (allChecked ? ' checked' : '') +
                '> <span>(Select all)</span>';
            frag.appendChild(allLbl);

            visible.forEach(function (v) {
                var lbl = document.createElement('label');
                lbl.className = 'cf-item';
                var display = v === '' ? '(Blanks)' : v;
                var span = document.createElement('span');
                span.textContent = display;
                var cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.className = 'cf-val';
                cb.checked = selectedSet.has(v);
                cb.addEventListener('change', function () {
                    if (cb.checked) selectedSet.add(v); else selectedSet.delete(v);
                    var allAfter = visible.every(function (x) { return selectedSet.has(x); });
                    var allBox = listEl.querySelector('.cf-all');
                    if (allBox) allBox.checked = allAfter;
                });
                lbl.appendChild(cb);
                lbl.appendChild(span);
                frag.appendChild(lbl);
            });

            listEl.innerHTML = '';
            listEl.appendChild(frag);

            var allBox = listEl.querySelector('.cf-all');
            allBox.addEventListener('change', function () {
                if (allBox.checked) visible.forEach(function (v) { selectedSet.add(v); });
                else                visible.forEach(function (v) { selectedSet.delete(v); });
                renderList();
            });
        }
        renderList();

        searchEl.addEventListener('input', renderList);

        pop.querySelectorAll('.cf-sort-item').forEach(function (el) {
            el.addEventListener('click', function (e) {
                e.stopPropagation();
                applySort(table, colIndex, el.getAttribute('data-sort'), kind);
                th._cfState = {
                    active:   !!(th._cfState && th._cfState.active),
                    selected: th._cfState ? th._cfState.selected : null,
                    sort:     el.getAttribute('data-sort')
                };
                refreshHeaderState(th);
                closePopovers();
            });
        });

        pop.querySelector('.cf-btn-cancel').addEventListener('click', function (e) {
            e.stopPropagation();
            closePopovers();
        });
        pop.querySelector('.cf-btn-clear').addEventListener('click', function (e) {
            e.stopPropagation();
            th._cfState = null;
            applyFilters(table);
            refreshHeaderState(th);
            closePopovers();
        });
        pop.querySelector('.cf-btn-primary').addEventListener('click', function (e) {
            e.stopPropagation();
            th._cfState = {
                active:   selectedSet.size < distinct.length,
                selected: selectedSet,
                sort:     th._cfState ? th._cfState.sort : null
            };
            applyFilters(table);
            refreshHeaderState(th);
            closePopovers();
        });

        // Escape closes the popover
        pop.addEventListener('keydown', function (e) {
            if (e.key === 'Escape') closePopovers();
        });

        // Focus search for quick typing
        setTimeout(function () { try { searchEl.focus(); } catch (_) {} }, 10);
    }

    function applySort(table, colIndex, direction, kind) {
        if (!table.tBodies[0]) return;
        var tbody = table.tBodies[0];
        var rows = Array.from(tbody.rows);
        rows.sort(function (a, b) {
            var av = getCellText(a, colIndex);
            var bv = getCellText(b, colIndex);
            if (kind === 'number') {
                var an = parseFloat(String(av).replace(/,/g, ''));
                var bn = parseFloat(String(bv).replace(/,/g, ''));
                var aNaN = isNaN(an), bNaN = isNaN(bn);
                if (aNaN && bNaN) return 0;
                if (aNaN) return 1;
                if (bNaN) return -1;
                return direction === 'desc' ? bn - an : an - bn;
            }
            if (kind === 'date') {
                var ad = Date.parse(av), bd = Date.parse(bv);
                if (isNaN(ad) && isNaN(bd)) return 0;
                if (isNaN(ad)) return 1;
                if (isNaN(bd)) return -1;
                return direction === 'desc' ? bd - ad : ad - bd;
            }
            return direction === 'desc'
                ? String(bv).localeCompare(String(av), undefined, { numeric: true })
                : String(av).localeCompare(String(bv), undefined, { numeric: true });
        });
        rows.forEach(function (r) { tbody.appendChild(r); });
    }

    function applyFilters(table) {
        if (!table.tBodies[0] || !table.tHead) return;
        var ths = Array.from(table.tHead.rows[0].cells);
        Array.from(table.tBodies[0].rows).forEach(function (r) {
            var visible = true;
            ths.forEach(function (th) {
                var st = th._cfState;
                if (!st || !st.active || !st.selected) return;
                var val = getCellText(r, th.cellIndex);
                if (!st.selected.has(val)) visible = false;
            });
            if (visible) r.classList.remove('cf-hidden');
            else         r.classList.add('cf-hidden');
        });
    }

    function refreshHeaderState(th) {
        if (th._cfState && th._cfState.active) th.classList.add('cf-filtered');
        else                                   th.classList.remove('cf-filtered');
    }

    function install(table) {
        if (!table || !table.tHead || table.dataset.cfInstalled === '1') return;
        Array.from(table.tHead.rows[0].cells).forEach(function (th) {
            var kind = th.getAttribute('data-colfilter');
            if (!kind) return;

            // Mark whole header as clickable
            th.classList.add('cf-clickable');

            // Add the funnel icon (visual indicator; click anywhere on the TH works too)
            if (!th.querySelector('.cf-icon')) {
                var icon = document.createElement('span');
                icon.className = 'cf-icon';
                icon.innerHTML = FUNNEL_SVG;
                icon.title = 'Filter / sort';
                icon.setAttribute('aria-hidden', 'true');
                th.appendChild(icon);
            }

            // Click handler on the whole TH
            th.addEventListener('click', function (e) {
                // Ignore clicks that bubbled from inside the popover
                if (e.target.closest && e.target.closest('.cf-pop')) return;
                e.preventDefault();
                e.stopPropagation();
                // Toggle if already open for this header
                if (_activePop && _activePop._ownerTh === th) {
                    closePopovers();
                    return;
                }
                buildPopover(th, table);
                if (_activePop) _activePop._ownerTh = th;
            });

            // Keyboard activation
            th.setAttribute('tabindex', '0');
            th.setAttribute('role', 'button');
            th.addEventListener('keydown', function (e) {
                if (e.key !== 'Enter' && e.key !== ' ') return;
                e.preventDefault();
                if (_activePop && _activePop._ownerTh === th) {
                    closePopovers();
                    return;
                }
                buildPopover(th, table);
                if (_activePop) _activePop._ownerTh = th;
            });
        });
        table.dataset.cfInstalled = '1';
    }

    function installAll(root) {
        (root || document).querySelectorAll('[data-colfilter-table]').forEach(install);
    }

    global.ColumnFilter = { install: install, installAll: installAll };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () { installAll(); });
    } else {
        installAll();
    }

    // Close on outside click (single document-level listener)
    document.addEventListener('mousedown', function (e) {
        if (!_activePop) return;
        if (e.target.closest('.cf-pop')) return;
        if (e.target.closest('.cf-clickable')) return; // let the TH click toggle handle it
        closePopovers();
    });

    // Reposition / close popover on window resize.
    // Crucially NOT listening to 'scroll' in capture phase — that would close
    // the popover any time the user scrolled inside the data table.
    window.addEventListener('resize', closePopovers);
}(window));
