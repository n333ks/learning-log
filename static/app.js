/* ── Inventory Table Filters ──────────────────────────────────────────────── */
function applyInventoryFilters() {
    const q         = (document.getElementById('inventorySearch')?.value || '').toLowerCase().trim();
    const container = (document.getElementById('containerFilter')?.value || '');
    const rows      = Array.from(document.querySelectorAll('#inventoryTable tbody tr.inv-row'));

    if (container) {
        // Container selected: filter every row individually — show only exact matches
        rows.forEach(row => {
            const containerMatch = row.dataset.container === container;
            const searchMatch    = !q ||
                row.dataset.design.toLowerCase().includes(q) ||
                row.dataset.sku.toLowerCase().includes(q);
            row.style.display = (containerMatch && searchMatch) ? '' : 'none';
        });
    } else {
        // No container filter: group by variant so all rows of a variant show/hide together
        const groups = {};
        rows.forEach(row => {
            const key = row.dataset.design + '|||' + row.dataset.sku;
            if (!groups[key]) groups[key] = [];
            groups[key].push(row);
        });

        Object.entries(groups).forEach(([key, groupRows]) => {
            const [design, sku] = key.split('|||');
            const matches = !q ||
                design.toLowerCase().includes(q) ||
                sku.toLowerCase().includes(q);
            groupRows.forEach(row => {
                row.style.display = matches ? '' : 'none';
            });
        });
    }

    reapplyGaps();
    updateCount('inventoryCount', '#inventoryTable tbody tr.inv-row');
}

// Recompute design-gap / variant-gap / variant-lead on the currently visible rows
function reapplyGaps() {
    const allRows = Array.from(document.querySelectorAll('#inventoryTable tbody tr.inv-row'));
    const visible = allRows.filter(r => r.style.display !== 'none');

    // Clear dynamic classes and wipe any previously injected counts from all rows
    allRows.forEach(row => {
        row.classList.remove('design-gap', 'variant-gap', 'variant-lead');
        clearCountCells(row);
    });

    let prevDesign  = null;
    let prevVariant = null;

    visible.forEach(row => {
        const design  = row.dataset.design;
        const variant = design + '|||' + row.dataset.sku;

        if (prevDesign !== null && design !== prevDesign) {
            row.classList.add('design-gap');
        } else if (prevVariant !== null && variant !== prevVariant) {
            row.classList.add('variant-gap');
        }

        if (variant !== prevVariant) {
            row.classList.add('variant-lead');
            injectCountCells(row);
        }

        prevDesign  = design;
        prevVariant = variant;
    });
}

// Columns 6-10 (0-indexed) are In Stock, In Prod, Pre-Sale, Optimal, Variance
function injectCountCells(row) {
    const tds      = row.querySelectorAll('td');
    const variance = parseInt(row.dataset.variance ?? '', 10);
    const values   = [
        row.dataset.inStock  ?? '',
        row.dataset.inProd   ?? '',
        row.dataset.preSale  ?? '',
        row.dataset.optimal  ?? '',
        isNaN(variance) ? '' :
            variance < 0
                ? `<span class="text-red font-bold">${variance}</span>`
                : `<span class="text-muted">${variance}</span>`
    ];
    values.forEach((val, i) => {
        if (tds[6 + i]) {
            tds[6 + i].innerHTML  = val;
            tds[6 + i].style.textAlign = 'center';
            tds[6 + i].dataset.injected = '1';
        }
    });
}

function clearCountCells(row) {
    const tds = row.querySelectorAll('td');
    for (let i = 6; i <= 10; i++) {
        if (tds[i] && tds[i].dataset.injected === '1') {
            tds[i].innerHTML = '';
            delete tds[i].dataset.injected;
        }
    }
}

function filterInventory(query) { applyInventoryFilters(); }

// ── Unit counter helper ────────────────────────────────────────────────────────
function updateCount(counterId, selector) {
    const el = document.getElementById(counterId);
    if (!el) return;
    const total   = document.querySelectorAll(selector).length;
    const visible = Array.from(document.querySelectorAll(selector)).filter(r => r.style.display !== 'none').length;
    el.textContent = visible === total ? `${total} units` : `${visible} of ${total} units`;
}

/* ── Flash Auto-Dismiss + Page Init ──────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
    const flashes = document.querySelectorAll('.flash');
    flashes.forEach(flash => {
        setTimeout(() => {
            flash.classList.add('fade-out');
            setTimeout(() => flash.remove(), 300);
        }, 3000);
    });

    // Initialize counters on page load
    updateCount('inventoryCount', '#inventoryTable tbody tr.inv-row');
    updateCount('salesCount', '.sales-row');
    updateCount('warehouseCount', '.warehouse-row');
});
