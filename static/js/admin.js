/* ── Admin Dashboard JS ───────────────────────────────────────
   Handles:
   1. Delete image confirmation (replaces bare confirm())
   2. JSON syntax validation before saving data.json
   3. Auto-resize the data.json textarea
   4. File input label update
──────────────────────────────────────────────────────────────── */

/* 1. Delete confirmation — called inline from each delete form */
function confirmDelete(imageName) {
    return window.confirm(`Delete "${imageName}"?\n\nThis cannot be undone.`);
}

document.addEventListener('DOMContentLoaded', function () {

    /* 2. JSON syntax validation before saving data.json */
    const dataJsonTextarea = document.getElementById('data_json');
    if (dataJsonTextarea) {
        const saveForm = dataJsonTextarea.closest('form');

        if (saveForm) {
            saveForm.addEventListener('submit', function (e) {
                const raw = dataJsonTextarea.value.trim();
                try {
                    const parsed = JSON.parse(raw);

                    // Warn if key sections are empty
                    const warnings = [];
                    if (!parsed.sample_paragraphs ||
                        (!parsed.sample_paragraphs.easy?.length &&
                         !parsed.sample_paragraphs.medium?.length &&
                         !parsed.sample_paragraphs.hard?.length)) {
                        warnings.push('• sample_paragraphs appears empty');
                    }
                    if (!parsed.excel_practical_tasks?.length) {
                        warnings.push('• excel_practical_tasks is empty');
                    }
                    if (!parsed.excel_quiz_questions?.length) {
                        warnings.push('• excel_quiz_questions is empty');
                    }

                    if (warnings.length > 0) {
                        const proceed = window.confirm(
                            `Warning — the following sections look empty:\n\n` +
                            warnings.join('\n') +
                            `\n\nSave anyway?`
                        );
                        if (!proceed) e.preventDefault();
                    }
                } catch (err) {
                    e.preventDefault();
                    showAdminToast(
                        `Invalid JSON — cannot save.\n${err.message}`,
                        'error'
                    );
                    dataJsonTextarea.focus();
                    dataJsonTextarea.style.borderColor = 'var(--red-500)';
                    dataJsonTextarea.style.boxShadow   = '0 0 0 3px rgba(239,68,68,0.15)';
                }
            });

            // Reset error highlight on edit
            dataJsonTextarea.addEventListener('input', function () {
                this.style.borderColor = '';
                this.style.boxShadow   = '';
            });
        }

        /* 3. Auto-resize textarea as content changes */
        function autoResize(el) {
            el.style.height = 'auto';
            el.style.height = Math.min(el.scrollHeight, 600) + 'px';
        }
        dataJsonTextarea.addEventListener('input', function () {
            autoResize(this);
        });
        autoResize(dataJsonTextarea); // run once on load
    }

    /* 4. File input — show selected filename next to the input */
    const fileInput = document.getElementById('file');
    if (fileInput) {
        fileInput.addEventListener('change', function () {
            const existing = this.parentElement.querySelector('.js-filename');
            const label    = existing || document.createElement('span');
            label.className = 'js-filename';
            label.style.cssText = `
                display: inline-block;
                margin-top: var(--sp-2);
                font-size: var(--text-xs);
                color: var(--teal-600);
                font-family: var(--font-mono);
                font-weight: 500;
            `;
            label.textContent = this.files[0]
                ? `✓ ${this.files[0].name}`
                : '';
            if (!existing) this.parentElement.appendChild(label);
        });
    }

});

/* ── Toast notification ───────────────────────────────────────
   Lightweight alternative to alert() for non-blocking feedback
──────────────────────────────────────────────────────────────── */
function showAdminToast(msg, type = 'error') {
    document.querySelector('.js-admin-toast')?.remove();

    const typeStyles = {
        error:   'background:var(--red-50);color:var(--red-700);border:1px solid var(--red-200);',
        success: 'background:var(--green-50);color:var(--green-700);border:1px solid var(--green-200);',
        warning: 'background:var(--amber-50);color:var(--amber-700);border:1px solid var(--amber-200);',
    };

    const toast = document.createElement('div');
    toast.className = 'js-admin-toast';
    toast.style.cssText = `
        position: fixed;
        bottom: 80px;
        left: 50%;
        transform: translateX(-50%);
        z-index: 3000;
        min-width: 280px;
        max-width: 460px;
        padding: 12px 18px;
        border-radius: 10px;
        font-family: var(--font-body);
        font-size: 13px;
        font-weight: 500;
        line-height: 1.5;
        text-align: center;
        box-shadow: 0 8px 24px rgba(14,26,92,0.15);
        white-space: pre-line;
        animation: fadeUp 0.25s ease forwards;
        ${typeStyles[type] || typeStyles.error}
    `;
    toast.textContent = msg;
    document.body.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity    = '0';
        toast.style.transition = 'opacity 0.3s ease';
        setTimeout(() => toast.remove(), 320);
    }, 4000);
}