// animal-form.js — client-side validation helpers
// Duplicate tag check fires on blur of Tag Number 1
// Weight range highlight fires on input

(function () {
    'use strict';

    // ── Vendor typeahead search ──────────────────────────────────────────
    const vendorSearch   = document.getElementById('vendorSearch');
    const vendorDropdown = document.getElementById('vendorDropdown');
    const vendorIdHidden = document.getElementById('vendorIdHidden');
    const vendorFreeText = document.getElementById('vendorNameFreeText');
    const vendorHint     = document.getElementById('vendorHint');

    if (vendorSearch) {
        let vendorTimer;

        vendorSearch.addEventListener('input', function () {
            clearTimeout(vendorTimer);
            const term = this.value.trim();

            // Reset hidden fields when user types
            vendorIdHidden.value = '0';
            vendorFreeText.value = term;

            if (!term) {
                vendorDropdown.style.display = 'none';
                return;
            }

            vendorTimer = setTimeout(async () => {
                try {
                    const url = `/Animal/SearchVendors?term=${encodeURIComponent(term)}`;
                    const res  = await fetch(url);
                    const data = await res.json();

                    vendorDropdown.innerHTML = '';

                    // Render existing matches
                    data.forEach(v => {
                        const row = document.createElement('div');
                        row.textContent  = v.name;
                        row.style.cssText = 'padding:9px 14px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--color-border-tertiary)';
                        row.addEventListener('mouseenter', () => row.style.background = 'var(--color-background-secondary)');
                        row.addEventListener('mouseleave', () => row.style.background = '');
                        row.addEventListener('mousedown', (e) => {
                            e.preventDefault();
                            vendorSearch.value   = v.name;
                            vendorIdHidden.value = v.id;
                            vendorFreeText.value = v.name;
                            vendorDropdown.style.display = 'none';
                            vendorHint.textContent = 'Existing vendor selected';
                            vendorHint.style.color = 'var(--color-text-success)';
                            vendorIdHidden.dispatchEvent(new Event('change'));
                        });
                        vendorDropdown.appendChild(row);
                    });

                    // Add "Use new vendor: X" option at the bottom
                    if (term.length > 1) {
                        const newRow = document.createElement('div');
                        newRow.textContent  = `+ Add new vendor: "${term}"`;
                        newRow.style.cssText = 'padding:9px 14px;cursor:pointer;font-size:13px;color:var(--color-text-info);font-weight:500;background:var(--color-background-info)';
                        newRow.addEventListener('mousedown', (e) => {
                            e.preventDefault();
                            vendorSearch.value   = term;
                            vendorIdHidden.value = '0';
                            vendorFreeText.value = term;
                            vendorDropdown.style.display = 'none';
                            vendorHint.textContent = `New vendor "${term}" will be created on save`;
                            vendorHint.style.color = 'var(--color-text-info)';
                            vendorIdHidden.dispatchEvent(new Event('change'));
                        });
                        vendorDropdown.appendChild(newRow);
                    }

                    vendorDropdown.style.display = vendorDropdown.children.length ? 'block' : 'none';
                } catch {
                    vendorDropdown.style.display = 'none';
                }
            }, 250);
        });

        // Hide dropdown when clicking elsewhere
        vendorSearch.addEventListener('blur', () => {
            setTimeout(() => { vendorDropdown.style.display = 'none'; }, 200);
        });

    }


    const tag1Input       = document.getElementById('tag1Input');
    const killDateInput   = document.getElementById('killDateInput');
    const vendorIdInput   = document.getElementById('vendorIdHidden');
    const liveWeightInput = document.getElementById('liveWeightInput');
    const tag1Feedback    = document.getElementById('tag1Feedback');

    // ── Duplicate tag check ──────────────────────────────────────────────
    if (tag1Input && killDateInput && vendorIdInput) {
        tag1Input.addEventListener('blur', checkDuplicateTag);
        killDateInput.addEventListener('change', checkDuplicateTag);
        vendorIdInput.addEventListener('change', checkDuplicateTag);
    }

    async function checkDuplicateTag() {
        const tag1     = tag1Input?.value?.trim();
        const killDate = killDateInput?.value;
        const vendorId = vendorIdInput?.value;

        if (!tag1 || !killDate || !vendorId) return;

        tag1Feedback.textContent = 'Checking…';
        tag1Feedback.className   = 'field-hint';

        try {
            const url = `${CHECK_TAG_URL}?tag1=${encodeURIComponent(tag1)}&killDate=${encodeURIComponent(killDate)}&vendorId=${encodeURIComponent(vendorId)}&controlNo=${CONTROL_NO}`;
            const res  = await fetch(url);
            const data = await res.json();

            if (data.isDuplicate) {
                tag1Input.classList.add('input-error');
                tag1Feedback.textContent = `Tag "${tag1}" already exists for this vendor on this kill date.`;
                tag1Feedback.className   = 'field-error';
            } else {
                tag1Input.classList.remove('input-error');
                tag1Feedback.textContent = 'Tag is available.';
                tag1Feedback.className   = 'field-ok';
            }
        } catch {
            tag1Feedback.textContent = '';
        }
    }

    // ── Live weight range highlight ──────────────────────────────────────
    const WEIGHT_MIN = 300;
    const WEIGHT_MAX = 2500;

    if (liveWeightInput) {
        liveWeightInput.addEventListener('input', function () {
            const val = parseFloat(this.value);
            if (!this.value) {
                this.classList.remove('input-warn', 'input-ok');
                return;
            }
            if (val < WEIGHT_MIN || val > WEIGHT_MAX) {
                this.classList.add('input-warn');
                this.classList.remove('input-ok');
            } else {
                this.classList.remove('input-warn');
                this.classList.add('input-ok');
            }
        });
    }

    // ── Kill date — purchase date check ──────────────────────────────────
    const purchaseDateInput = document.querySelector('input[name="PurchaseDate"]');
    if (killDateInput && purchaseDateInput) {
        killDateInput.addEventListener('change', validateDates);
        purchaseDateInput.addEventListener('change', validateDates);
    }

    function validateDates() {
        const kill     = new Date(killDateInput.value);
        const purchase = new Date(purchaseDateInput.value);
        if (killDateInput.value && purchaseDateInput.value && kill < purchase) {
            killDateInput.classList.add('input-error');
        } else {
            killDateInput.classList.remove('input-error');
        }
    }

    // ── Auto-dismiss flash messages after 4 seconds ──────────────────────
    const flash = document.querySelector('.flash');
    if (flash) {
        setTimeout(() => {
            flash.style.opacity = '0';
            flash.style.transition = 'opacity 0.4s';
            setTimeout(() => flash.remove(), 400);
        }, 4000);
    }

})();
