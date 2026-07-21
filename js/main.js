/* Site scripts — language toggle + side-nav mobile toggle + access gate */

(function () {
    'use strict';

    /* ---------- Google Analytics (GA4) ----------
       Loads on every page (this script is included site-wide). Property: philiphong.org */
    (function initAnalytics() {
        var GA_ID = 'G-0SR83J0TBZ';
        var s = document.createElement('script');
        s.async = true;
        s.src = 'https://www.googletagmanager.com/gtag/js?id=' + GA_ID;
        document.head.appendChild(s);
        window.dataLayer = window.dataLayer || [];
        function gtag() { window.dataLayer.push(arguments); }
        window.gtag = gtag;
        gtag('js', new Date());
        gtag('config', GA_ID);
    })();

    /* ---------- Access gate (client-side friction only, NOT real security) ----------
       SHA-256 of the 6-digit code. A determined visitor can find this in the JS and
       brute-force the 1M-space combination. It keeps casual visitors out, no more. */
    const ACCESS_HASH = 'e0bc60c82713f64ef8a57c0c40d02ce24fd0141d5cc3086259c19b1e62a62bea';
    const ACCESS_KEY  = 'ph-access-ok';

    async function sha256Hex(str) {
        const buf = await crypto.subtle.digest(
            'SHA-256',
            new TextEncoder().encode(str)
        );
        return Array.from(new Uint8Array(buf))
                    .map(function (b) { return b.toString(16).padStart(2, '0'); })
                    .join('');
    }

    function unlockPage() {
        document.body.classList.remove('requires-access');
    }

    function initAccessGate() {
        if (!document.body.classList.contains('requires-access')) return;

        let unlocked = false;
        try { unlocked = localStorage.getItem(ACCESS_KEY) === '1'; } catch (e) {}
        if (unlocked) { unlockPage(); return; }

        const gate = document.createElement('div');
        gate.className = 'access-gate';
        gate.innerHTML =
            '<div class="access-card">' +
              '<h2>Course Access</h2>' +
              '<p>Please enter the access code provided by your instructor.</p>' +
              '<form class="access-form">' +
                '<input type="text" class="access-input" inputmode="numeric" pattern="[0-9]*" ' +
                       'autocomplete="off" aria-label="Access code">' +
                '<button type="submit" class="access-submit">Continue</button>' +
              '</form>' +
              '<div class="access-error" role="alert" hidden>Incorrect code. Please try again.</div>' +
            '</div>';
        document.body.appendChild(gate);

        const card  = gate.querySelector('.access-card');
        const input = gate.querySelector('.access-input');
        const form  = gate.querySelector('.access-form');
        const error = gate.querySelector('.access-error');

        setTimeout(function () { input.focus(); }, 40);

        form.addEventListener('submit', async function (e) {
            e.preventDefault();
            const val = (input.value || '').trim();
            if (!val) return;

            let hash = '';
            try {
                hash = await sha256Hex(val);
            } catch (err) {
                error.textContent = 'Browser does not support secure hashing. Please use a modern browser.';
                error.hidden = false;
                return;
            }

            if (hash === ACCESS_HASH) {
                try { localStorage.setItem(ACCESS_KEY, '1'); } catch (e) {}
                unlockPage();
                gate.remove();
            } else {
                error.hidden = false;
                card.classList.remove('shake');
                void card.offsetWidth;  /* reflow to replay animation */
                card.classList.add('shake');
                input.value = '';
                input.focus();
            }
        });
    }

    /* ---------- Language toggle (week pages) ---------- */
    const STORAGE_KEY = 'ph-lang';

    function setLanguage(lang) {
        const groups = document.querySelectorAll('.resources');
        if (!groups.length) return;

        groups.forEach(function (group) {
            if (group.getAttribute('data-lang') === lang) {
                group.classList.add('active');
            } else {
                group.classList.remove('active');
            }
        });

        const buttons = document.querySelectorAll('.language-toggle button');
        buttons.forEach(function (btn) {
            if (btn.getAttribute('data-lang') === lang) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        try { localStorage.setItem(STORAGE_KEY, lang); } catch (e) {}
    }

    function initLanguageToggle() {
        const buttons = document.querySelectorAll('.language-toggle button');
        if (!buttons.length) return;

        let initial = 'en';
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            if (saved === 'en' || saved === 'ko') initial = saved;
        } catch (e) {}

        setLanguage(initial);

        buttons.forEach(function (btn) {
            btn.addEventListener('click', function () {
                const lang = btn.getAttribute('data-lang');
                if (lang) setLanguage(lang);
            });
        });
    }

    /* ---------- Mobile side-nav toggle ---------- */
    function initSideNav() {
        const toggle = document.querySelector('.mobile-menu-toggle');
        const nav = document.querySelector('.side-nav');
        const backdrop = document.querySelector('.nav-backdrop');
        if (!toggle || !nav) return;

        function open() {
            nav.classList.add('open');
            if (backdrop) backdrop.classList.add('visible');
            toggle.setAttribute('aria-expanded', 'true');
        }
        function close() {
            nav.classList.remove('open');
            if (backdrop) backdrop.classList.remove('visible');
            toggle.setAttribute('aria-expanded', 'false');
        }

        toggle.addEventListener('click', function () {
            if (nav.classList.contains('open')) close(); else open();
        });

        if (backdrop) backdrop.addEventListener('click', close);

        nav.querySelectorAll('a').forEach(function (a) {
            a.addEventListener('click', close);
        });
    }

    document.addEventListener('DOMContentLoaded', function () {
        initAccessGate();
        initLanguageToggle();
        initSideNav();
    });
})();
