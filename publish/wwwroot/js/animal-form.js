// animal-form.js — BarnData entry form client-side logic
(function () {
    'use strict';

    //  Vendor typeahead 
    const vendorSearch   = document.getElementById('vendorSearch');
    const vendorDropdown = document.getElementById('vendorDropdown');
    const vendorIdHidden = document.getElementById('vendorIdHidden');
    const vendorFreeText = document.getElementById('vendorNameFreeText');
    const vendorHint     = document.getElementById('vendorHint');
    const animalForm = document.getElementById('animalForm');

    /*let formSubmitting = false;
    const animalForm = document.getElementById('animalForm');
    if(animalForm){
        animalForm.addEventListener('submit', function () {
            formSubmitting = true;
            if(tag1Feedback) tag1Feedback.textContent = ''; 
        });
    }*/

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

        if(liveWeightInput){
            liveWeightInput.required = !isCons;
            if(isCons){
                liveWeightInput.setCustomValidity('');
            }
        }

        if (liveRateField)     liveRateField.style.display     = isCons ? 'none' : '';
        if (consRateField)     consRateField.style.display     = isCons ? '' : 'none';
        if (consignmentFields) consignmentFields.style.display = isCons ? '' : 'none';
        if (saleBillWeightHint)saleBillWeightHint.style.display= isCons ? 'none' : '';
        if (consWeightHint)    consWeightHint.style.display    = isCons ? '' : 'none';
        if (liveWtHint)        liveWtHint.textContent          = isCons
            ? 'For records only — pricing uses hot weight'
            : 'Expected range: 300–3,000 lbs';
        if (hotWtReq)          hotWtReq.style.display          = isCons ? '' : 'none';

        updateCostPreview();
    }

    if (liveRateInput) {
            liveRateInput.addEventListener('focus', function () {
                if (this.value === '0' || this.value === '0.0' || this.value === '0.0000') {
                    this.value = '';
                }
            });
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
    const WEIGHT_MIN = 300, WEIGHT_MAX = 3000;
    if (liveWeightInput) {
        liveWeightInput.addEventListener('input', function () {
            const val = parseFloat(this.value);
            if (!this.value) { this.classList.remove('input-warn', 'input-ok'); return; }
            this.classList.toggle('input-warn', val < WEIGHT_MIN || val > WEIGHT_MAX);
            this.classList.toggle('input-ok',   val >= WEIGHT_MIN && val <= WEIGHT_MAX);
        });
    }

        // Duplicate tag check
    const tag1Input    = document.getElementById('tag1Input');
    const tag1Feedback = document.getElementById('tag1Feedback');

    let formSubmitting = false;
    let ajaxSaving = false;

    
    // Wire Save & add another as a direct click (type="button") - avoids e.submitter browser issues
    if (animalForm) {
    animalForm.addEventListener('click', async function (e) {
        const btn = e.target.closest('#btnSaveAndAdd');
        if (!btn) return;

        e.preventDefault();
        if (ajaxSaving) return;

        ajaxSaving = true;
        if (tag1Feedback) tag1Feedback.textContent = '';

        const oldText = btn.textContent;
        btn.disabled = true;
        btn.textContent = 'Saving...';

        try {
            var typedVendor = vendorSearch ? String(vendorSearch.value || '').trim() : '';
            if (vendorIdHidden && vendorFreeText) {
                if (vendorIdHidden.value === '0') {
                    vendorFreeText.value = typedVendor;   // allow new vendor creation
                } else {
                    vendorFreeText.value = '';            // existing selected vendor
                }
            }
            var validationError = validateBeforeSave();
            if(validationError){
                showAjaxError(validationError);
                btn.disabled=false;
                btn.textContent=oldText;
                ajaxSaving=false;
                return;
            }
            const formData = new FormData(animalForm);
            formData.set('saveAndAdd', '1');

            const resp = await fetch(
                (window.BarnUrls && window.BarnUrls.saveAndAddAjax) || '/Animal/CreateAjaxSaveAndAdd',
                {
                    method: 'POST',
                    body: formData,
                    headers: { 'X-Requested-With': 'XMLHttpRequest' }
                }
            );

            let data = null;
            const ct = resp.headers.get('content-type') || '';
            if (ct.indexOf('application/json') >= 0) {
                data = await resp.json();
            } else {
                const raw = await resp.text();
                data = { success: false, message: 'Unexpected response from server. ' + raw };
            }

            if (!resp.ok || !data.success) {
                showAjaxError(data);
                return;
            }

            upsertSessionRecord(data.record);
            clearForNextEntry();

        } catch {
            showAjaxError({ message: 'Unable to save right now. Please try again.' });
        } finally {
            btn.disabled = false;
            btn.textContent = oldText;
            ajaxSaving = false;
        }
    });
}

    // For regular submit buttons (Save, Save and go to list)
    if (animalForm) {
        animalForm.addEventListener('submit', function () {
            formSubmitting = true;
            if (tag1Feedback) tag1Feedback.textContent = '';
        });
    }

    if (tag1Input) {
        tag1Input.addEventListener('blur', function () {
            if (!formSubmitting && !ajaxSaving) checkDuplicateTag();
        });
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

    function upsertSessionRecord(record) {
        var records = Array.isArray(window.sessionRecords)
            ? window.sessionRecords
            : JSON.parse(sessionStorage.getItem('barnSessionRecords') || '[]');

        var incomingId = Number(record.controlNo || 0);
        var idx = incomingId > 0
            ? records.findIndex(function (r) { return Number(r.controlNo || 0) === incomingId; })
            : -1;

        if (idx >= 0) records[idx] = record;
        else records.push(record);

        window.sessionRecords = records;
        sessionStorage.setItem('barnSessionRecords', JSON.stringify(records));

        if (typeof window.renderPreview === 'function') window.renderPreview();
    }

    function clearForNextEntry() {
        ['tag1Input', 'TagNumber2', 'Tag3', 'AnimalControlNumber',
         'liveWeightInput', 'hotWeightInput', 'FetalBlood',
         'Grade', 'State', 'BuyerName', 'VetName', 'Comment', 'OfficeUse2']
        .forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.value = '';
        });

        var condemned = document.getElementById('IsCondemned');
        if (condemned) condemned.checked = false;

        if (tag1Input) tag1Input.classList.remove('input-error');
        if (tag1Feedback) { tag1Feedback.textContent = ''; tag1Feedback.className = 'field-hint'; }

        updateCostPreview();
        if (tag1Input) tag1Input.focus();
    }

    function validateBeforeSave() {
    var isCons = purchaseTypeSelect && purchaseTypeSelect.value === 'Consignment Bill';

    var purchaseDateInput = document.getElementById('PurchaseDate');
    var animalTypeInput = document.getElementById('AnimalType');

    var errors = {};

    function addError(key, message, el) {
        errors[key] = [message];
        if (el) el.classList.add('input-error');
    }

    function clearError(el) {
        if (el) el.classList.remove('input-error');
    }

    var vendorId = vendorIdHidden ? String(vendorIdHidden.value || '').trim() : '';
    var vendorText = vendorSearch ? String(vendorSearch.value || '').trim() : '';
    var hasVendor = (vendorId && vendorId !== '0') || vendorText.length > 0;

    if (!hasVendor) addError('VendorID', 'Vendor is required.', vendorSearch);
    else clearError(vendorSearch);

    if (!purchaseTypeSelect || !purchaseTypeSelect.value) addError('PurchaseType', 'Purchase type is required.', purchaseTypeSelect);
    else clearError(purchaseTypeSelect);

    if (!purchaseDateInput || !purchaseDateInput.value) addError('PurchaseDate', 'Purchase date is required.', purchaseDateInput);
    else clearError(purchaseDateInput);

    if (!tag1Input || !tag1Input.value || !tag1Input.value.trim()) addError('TagNumber1', 'Tag Number 1 is required.', tag1Input);
    else clearError(tag1Input);

    if (!animalTypeInput || !animalTypeInput.value) addError('AnimalType', 'Animal type is required.', animalTypeInput);
    else clearError(animalTypeInput);

    if (!isCons) {
        if (!liveWeightInput || !liveWeightInput.value || parseFloat(liveWeightInput.value) <= 0)
            addError('LiveWeight', 'Live weight is required for Sale Bill.', liveWeightInput);
        else clearError(liveWeightInput);

        if (!liveRateInput || !liveRateInput.value || parseFloat(liveRateInput.value) <= 0)
            addError('LiveRate', 'Live rate is required for Sale Bill.', liveRateInput);
        else clearError(liveRateInput);
    } else {
        clearError(liveWeightInput);
        clearError(liveRateInput);
    }

    return Object.keys(errors).length ? { success: false, errors: errors } : null;
}

    function showAjaxError(data) {
    var fieldLabels = {
        VendorID: 'Vendor',
        PurchaseType: 'Purchase type',
        PurchaseDate: 'Purchase date',
        TagNumber1: 'Tag 1',
        AnimalType: 'Animal type',
        LiveWeight: 'Live weight',
        LiveRate: 'Live rate'
    };

    var lines = [];
    if (data && data.errors && typeof data.errors === 'object') {
        Object.keys(data.errors).forEach(function (key) {
            var msgs = data.errors[key];
            var label = fieldLabels[key] || key;
            if (Array.isArray(msgs) && msgs.length) {
                var msg = String(msgs[0] || '').trim();

                // If server message is already readable ("Vendor is required."), show only that
                if (/is required/i.test(msg)) {
                    lines.push(msg);
                } else {
                    lines.push(label + ': ' + msg);
                }
            }
        });
    }

    if (!lines.length) {
        lines.push((data && data.message) ? data.message : 'Validation failed. Please check required fields.');
    }

    var summary = document.querySelector('[data-valmsg-summary], .validation-summary-errors, .alert-error');
    if (summary) {
        summary.style.display = '';
        summary.innerHTML = '<ul><li>' + lines.join('</li><li>') + '</li></ul>';
    } else {
        alert('Please fix the following:\n• ' + lines.join('\n• '));
    }
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
