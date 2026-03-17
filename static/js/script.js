document.addEventListener('DOMContentLoaded', function () {

    /* ── Utility: disable autofill / suggestions ─────────────── */
    function disableAutofill(elements) {
        elements.forEach(el => {
            el.setAttribute('autocomplete', 'off');
            if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                el.setAttribute('autocorrect', 'off');
                el.setAttribute('autocapitalize', 'off');
                el.setAttribute('spellcheck', 'false');
            }
        });
    }

    // Apply globally
    disableAutofill(document.querySelectorAll('input, textarea, select'));


    /* ════════════════════════════════════════════════════════════
       1. SIGNUP FORM VALIDATION
    ════════════════════════════════════════════════════════════ */
    const signupForm = document.getElementById('signup-form');
    if (signupForm) {
        signupForm.addEventListener('submit', function (e) {
            const name          = document.getElementById('name')?.value.trim();
            const location      = document.getElementById('location')?.value.trim();
            const distance      = document.getElementById('distance')?.value.trim();
            const attemptNumber = document.getElementById('attempt_number')?.value;
            const dob           = document.getElementById('dob')?.value.trim();

            if (!name || !location || !distance || !attemptNumber || !dob) {
                e.preventDefault();
                showFormError('Please fill out all fields.');
                return;
            }

            if (isNaN(distance) || parseFloat(distance) < 0) {
                e.preventDefault();
                showFormError('Distance must be a valid non-negative number.');
                return;
            }

            if (!['1st', '2nd', '3rd'].includes(attemptNumber)) {
                e.preventDefault();
                showFormError('Please select a valid attempt number (1st, 2nd, or 3rd).');
                return;
            }

            if (dob) {
                const dobDate = new Date(dob);
                const today   = new Date();
                const age     = today.getFullYear() - dobDate.getFullYear() -
                    ((today.getMonth() < dobDate.getMonth() ||
                      (today.getMonth() === dobDate.getMonth() &&
                       today.getDate() < dobDate.getDate())) ? 1 : 0);

                if (isNaN(dobDate.getTime()) || dobDate > today) {
                    e.preventDefault();
                    showFormError('Please enter a valid date of birth.');
                    return;
                }
                if (age < 18) {
                    e.preventDefault();
                    showFormError('You must be at least 18 years old to sign up.');
                    return;
                }
            }
        });

        // Inline validation error — uses existing .alert styling
        function showFormError(msg) {
            let existing = signupForm.querySelector('.js-form-error');
            if (!existing) {
                existing = document.createElement('div');
                existing.className = 'alert alert-error js-form-error';
                signupForm.prepend(existing);
            }
            existing.textContent = msg;
            existing.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }


    /* ════════════════════════════════════════════════════════════
       2. HANDWRITTEN VERIFICATION
       Timer is driven here (not duplicated in the template).
       The inline <script> in handwritten_round.html handles only
       the visual timer bar (colour changes) — this handles the
       AJAX submit logic and the master countdown.
    ════════════════════════════════════════════════════════════ */
    const handwrittenForm  = document.getElementById('handwrittenForm');
    const imageElement     = document.getElementById('current-image');
    const inputField       = document.getElementById('handwritten_input');
    const hwLoader         = document.getElementById('loader');
    const hwTimerDisplay   = document.getElementById('hw-timer');       // updated id
    const hwTimerBar       = document.getElementById('hw-timer-bar');   // progress bar
    const hwTimerIcon      = document.getElementById('timerIcon');

    if (handwrittenForm && imageElement && inputField && hwLoader) {
        disableAutofill([inputField]);

        const TOTAL_HW = 10 * 60; // 600 seconds
        let hwTimeLeft = TOTAL_HW;
        let hwInterval = null;

        function updateHwTimer(s) {
            const m   = Math.floor(s / 60);
            const sec = s % 60;
            const fmt = `${m}:${String(sec).padStart(2, '0')}`;

            if (hwTimerDisplay) hwTimerDisplay.textContent = fmt;

            const pct = (s / TOTAL_HW) * 100;
            if (hwTimerBar) hwTimerBar.style.width = pct + '%';

            // Colour transitions matching the template's inline script
            if (pct > 50) {
                if (hwTimerBar) hwTimerBar.style.background =
                    'linear-gradient(90deg,var(--teal-500),var(--teal-400))';
                if (hwTimerDisplay) hwTimerDisplay.style.color = 'var(--navy-900)';
                if (hwTimerIcon)    hwTimerIcon.textContent = '⏱️';
            } else if (pct > 20) {
                if (hwTimerBar) hwTimerBar.style.background =
                    'linear-gradient(90deg,var(--amber-500),var(--amber-400))';
                if (hwTimerDisplay) hwTimerDisplay.style.color = 'var(--amber-700)';
                if (hwTimerIcon)    hwTimerIcon.textContent = '⚠️';
            } else {
                if (hwTimerBar) hwTimerBar.style.background =
                    'linear-gradient(90deg,var(--red-500),var(--red-400))';
                if (hwTimerDisplay) hwTimerDisplay.style.color = 'var(--red-600)';
                if (hwTimerIcon)    hwTimerIcon.textContent = '🚨';
            }
        }

        function submitHandwritten(formData) {
            if (hwLoader) hwLoader.style.display = 'block';

            fetch('/submit_handwritten', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    if (hwLoader) hwLoader.style.display = 'none';

                    if (data.completed) {
                        clearInterval(hwInterval);
                        // Show completion message inside the card
                        const card = handwrittenForm.closest('.card');
                        if (card) {
                            card.innerHTML = `
                                <div style="text-align:center;padding:var(--sp-10);">
                                    <div style="font-size:2.5rem;margin-bottom:var(--sp-4);">✅</div>
                                    <h2 style="text-align:center;color:var(--green-600);
                                        margin-bottom:var(--sp-3);">All Done!</h2>
                                    <p style="color:var(--gray-500);">
                                        Handwritten verification completed successfully.
                                    </p>
                                </div>`;
                        }
                        if (data.redirect) {
                            setTimeout(() => { window.location.href = data.redirect; }, 1200);
                        }
                    } else {
                        // Next image
                        if (imageElement && data.next_image_url) {
                            imageElement.src = data.next_image_url;
                        }
                        inputField.value = '';
                        disableAutofill([inputField]);
                        inputField.focus();
                    }
                })
                .catch(err => {
                    if (hwLoader) hwLoader.style.display = 'none';
                    console.error('Handwritten submit error:', err);
                    showToast('Submission failed. Please try again.', 'error');
                });
        }

        // Start countdown
        hwInterval = setInterval(() => {
            if (hwTimeLeft <= 0) {
                clearInterval(hwInterval);
                updateHwTimer(0);
                // Auto-submit with empty input on timeout
                const fd = new FormData(handwrittenForm);
                fd.set('handwritten_input', '');
                submitHandwritten(fd);
            } else {
                updateHwTimer(hwTimeLeft);
                hwTimeLeft--;
            }
        }, 1000);

        // Manual submit
        handwrittenForm.addEventListener('submit', function (e) {
            e.preventDefault();
            submitHandwritten(new FormData(e.target));
        });
    }


    /* ════════════════════════════════════════════════════════════
       3. EXCEL QUIZ — selected option highlight + submit
    ════════════════════════════════════════════════════════════ */
    const excelQuizForm = document.getElementById('excelQuizForm');
    const quizLoader    = document.getElementById('submission-loader');

    if (excelQuizForm) {
        // Highlight selected option
        excelQuizForm.querySelectorAll('input[type="radio"]').forEach(radio => {
            disableAutofill([radio]);
            radio.addEventListener('change', function () {
                const optionGroup = this.closest('.quiz-options');
                if (optionGroup) {
                    optionGroup.querySelectorAll('.quiz-option')
                        .forEach(lbl => lbl.classList.remove('selected'));
                }
                const parentOption = this.closest('.quiz-option');
                if (parentOption) parentOption.classList.add('selected');
            });

            // Restore selected state on page load (browser back/forward cache)
            if (radio.checked) {
                const parentOption = radio.closest('.quiz-option');
                if (parentOption) parentOption.classList.add('selected');
            }
        });

        excelQuizForm.addEventListener('submit', function (e) {
            e.preventDefault();
            if (quizLoader) quizLoader.style.display = 'block';

            fetch('/excel_quiz', { method: 'POST', body: new FormData(excelQuizForm) })
                .then(response => {
                    if (quizLoader) quizLoader.style.display = 'none';
                    if (response.redirected) {
                        window.location.href = response.url;
                    } else {
                        return response.json().then(data => {
                            if (data.error) showToast(data.error, 'error');
                            else window.location.reload();
                        });
                    }
                })
                .catch(err => {
                    if (quizLoader) quizLoader.style.display = 'none';
                    console.error('Quiz submit error:', err);
                    showToast('An error occurred while submitting the quiz.', 'error');
                });
        });
    }


    /* ════════════════════════════════════════════════════════════
       4. TOAST NOTIFICATION (replaces bare alert())
    ════════════════════════════════════════════════════════════ */
    function showToast(msg, type = 'error') {
        // Remove any existing toast
        document.querySelector('.js-toast')?.remove();

        const toast = document.createElement('div');
        toast.className = `js-toast alert alert-${type}`;
        toast.textContent = msg;
        toast.style.cssText = `
            position: fixed;
            bottom: 80px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 3000;
            min-width: 260px;
            max-width: 420px;
            text-align: center;
            box-shadow: 0 8px 24px rgba(14,26,92,0.18);
            animation: fadeUp 0.25s ease forwards;
        `;
        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transition = 'opacity 0.3s ease';
            setTimeout(() => toast.remove(), 320);
        }, 3500);
    }

});