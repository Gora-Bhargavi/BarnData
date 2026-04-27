// animal-form.js — BarnData entry form client-side logic
(function () {
    'use strict';

    //  Vendor typeahead 
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
            vendorIdHidden.value = '0';
            vendorFreeText.value = term;
            if (!term) { vendorDropdown.style.display = 'none'; return; }

            vendorTimer = setTimeout(async () => {
                try {
                    const res  = await fetch(`/Animal/SearchVendors?term=${encodeURIComponent(term)}`);
                    const data = await res.json();
                    vendorDropdown.innerHTML = '';

                    data.forEach(v => {
                        const row = document.createElement('div');
                        row.textContent  = v.name;
                        row.style.cssText = 'padding:9px 14px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--color-border-tertiary)';
                        row.addEventListener('mouseenter', () => row.style.background = 'var(--color-background-secondary)');
                        row.addEventListener('mouseleave', () => row.style.background = '');
                        row.addEventListener('mousedown', e => {
                            e.preventDefault();
                            vendorSearch.value   = v.name;
                            vendorIdHidden.value = v.id;
                            vendorFreeText.value = '';
                            vendorDropdown.style.display = 'none';
                            vendorHint.textContent = 'Existing vendor selected';
                            vendorHint.style.color = 'var(--color-text-success)';
                        });
                        vendorDropdown.appendChild(row);
                    });

                    if (term.length > 1) {
                        const newRow = document.createElement('div');
                        newRow.textContent = `+ Add new vendor: "${term}"`;
                        newRow.style.cssText = 'padding:9px 14px;cursor:pointer;font-size:13px;color:var(--color-text-info);font-weight:500;background:var(--color-background-info)';
                        newRow.addEventListener('mousedown', e => {
                            e.preventDefault();
                            vendorSearch.value   = term;
                            vendorIdHidden.value = '0';
                            vendorFreeText.value = term;
                            vendorDropdown.style.display = 'none';
                            vendorHint.textContent = `New vendor "${term}" will be created on save`;
                            vendorHint.style.color = 'var(--color-text-info)';
                        });
                        vendorDropdown.appendChild(newRow);
                    }

                    vendorDropdown.style.display = vendorDropdown.children.length ? 'block' : 'none';
                } catch { vendorDropdown.style.display = 'none'; }
            }, 250);
        });

        vendorSearch.addEventListener('blur', () => {
            setTimeout(() => { vendorDropdown.style.display = 'none'; }, 200);
        });
        vendorSearch.addEventListener('focus', () => {
            if (vendorSearch.value.trim()) vendorSearch.dispatchEvent(new Event('input'));
        });
    }

    //  Sale Bill / Consignment Bill dynamic switching 
    const purchaseTypeSelect  = document.getElementById('purchaseTypeSelect');
    const liveRateField       = document.getElementById('liveRateField');
    const consRateField       = document.getElementById('consRateField');
    const consignmentFields   = document.getElementById('consignmentFields');
    const saleBillWeightHint  = document.getElementById('saleBillWeightHint');
    const consWeightHint      = document.getElementById('consWeightHint');
    const liveWtHint          = document.getElementById('liveWtHint');
    const hotWtReq            = document.getElementById('hotWtReq');
    const costPreview         = document.getElementById('costPreview');
    const costPreviewVal      = document.getElementById('costPreviewVal');
    const liveWeightInput     = document.getElementById('liveWeightInput');
    const hotWeightInput      = document.getElementById('hotWeightInput');
    const liveRateInput       = document.getElementById('liveRateInput');
    const consignmentRateInput = document.getElementById('consignmentRateInput');

    function applyPurchaseType(type) {
        const isCons = type === 'Consignment Bill';

        if (liveRateField)     liveRateField.style.display     = isCons ? 'none' : '';
        if (consRateField)     consRateField.style.display     = isCons ? '' : 'none';
        if (consignmentFields) consignmentFields.style.display = isCons ? '' : 'none';
        if (saleBillWeightHint)saleBillWeightHint.style.display= isCons ? 'none' : '';
        if (consWeightHint)    consWeightHint.style.display    = isCons ? '' : 'none';
        if (liveWtHint)        liveWtHint.textContent          = isCons
            ? 'For records only — pricing uses hot weight'
            : 'Expected range: 300–2,500 lbs';
        if (hotWtReq)          hotWtReq.style.display          = isCons ? '' : 'none';

        updateCostPreview();
    }

    function readNumber(input) {
    if (!input) return 0;
    const value = parseFloat(input.value);
    return Number.isFinite(value) ? value : 0;
}

function updateCostPreview() {
    if (!costPreview || !costPreviewVal) return;

    const type = purchaseTypeSelect?.value || '';
    const isCons = type === 'Consignment Bill';

    const liveWt = readNumber(liveWeightInput);
    const hotWt = readNumber(hotWeightInput);
    const liveRate = readNumber(liveRateInput);
    const consRate = readNumber(consignmentRateInput);

    const cost = isCons
        ? hotWt * consRate
        : liveWt * liveRate;

    if (cost > 0) {
        costPreview.style.display = 'block';
        costPreviewVal.textContent =
            '$' + cost.toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
    } else {
        costPreview.style.display = 'none';
    }
}

    if (purchaseTypeSelect) {
        purchaseTypeSelect.addEventListener('change', () => applyPurchaseType(purchaseTypeSelect.value));
        // Apply on load
        applyPurchaseType(purchaseTypeSelect.value);
    }

    // Update cost preview when weights or rates change
    [liveWeightInput, hotWeightInput, liveRateInput, consignmentRateInput]
    .filter(Boolean)
    .forEach(el => {
        el.addEventListener('input', updateCostPreview);
        el.addEventListener('change', updateCostPreview);
    });

    // Live weight range highlight 
    const WEIGHT_MIN = 300, WEIGHT_MAX = 2500;
    if (liveWeightInput) {
        liveWeightInput.addEventListener('input', function () {
            const val = parseFloat(this.value);
            if (!this.value) { this.classList.remove('input-warn', 'input-ok'); return; }
            this.classList.toggle('input-warn', val < WEIGHT_MIN || val > WEIGHT_MAX);
            this.classList.toggle('input-ok',   val >= WEIGHT_MIN && val <= WEIGHT_MAX);
        });
    }

    // Duplicate tag check 
    const tag1Input   = document.getElementById('tag1Input');
    const tag1Feedback= document.getElementById('tag1Feedback');

    if (tag1Input) {
        tag1Input.addEventListener('blur', checkDuplicateTag);
    }

    async function checkDuplicateTag() {
        const tag1     = tag1Input?.value?.trim();
        const vendorId = vendorIdHidden?.value;
        if (!tag1 || !vendorId || vendorId === '0') return;

        tag1Feedback.textContent = 'Checking…';
        tag1Feedback.className   = 'field-hint';

        try {
            const url = `/Animal/CheckTag?tag1=${encodeURIComponent(tag1)}&vendorId=${encodeURIComponent(vendorId)}&controlNo=${typeof CONTROL_NO !== 'undefined' ? CONTROL_NO : 0}`;
            const res  = await fetch(url);
            const data = await res.json();

            if (data.isDuplicate) {
                tag1Input.classList.add('input-error');
                tag1Feedback.textContent = `Tag "${tag1}" already exists for this vendor.`;
                tag1Feedback.className   = 'field-error';
            } else {
                tag1Input.classList.remove('input-error');
                tag1Feedback.textContent = 'Tag is available.';
                tag1Feedback.className   = 'field-ok';
            }
        } catch { tag1Feedback.textContent = ''; }
    }

    // Auto-dismiss flash messages 
    const flash = document.querySelector('.flash');
    if (flash) {
        setTimeout(() => {
            flash.style.opacity = '0';
            flash.style.transition = 'opacity 0.4s';
            setTimeout(() => flash.remove(), 400);
        }, 4000);
    }

})();
