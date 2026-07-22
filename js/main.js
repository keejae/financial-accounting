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

    /* Pure-JS SHA-256 fallback (ASCII input). Used when window.crypto.subtle is
       unavailable — e.g. in-app browsers (KakaoTalk, Instagram, mail apps) or when
       the page is loaded over plain http://. Verified to match crypto.subtle. */
    function sha256Fallback(ascii) {
        function rightRotate(value, amount) {
            return (value >>> amount) | (value << (32 - amount));
        }
        var mathPow = Math.pow;
        var maxWord = mathPow(2, 32);
        var result = '';
        var words = [];
        var asciiBitLength = ascii.length * 8;

        var hash = sha256Fallback.h = sha256Fallback.h || [];
        var k = sha256Fallback.k = sha256Fallback.k || [];
        var primeCounter = k.length;

        var isComposite = {};
        for (var candidate = 2; primeCounter < 64; candidate++) {
            if (!isComposite[candidate]) {
                for (var i = 0; i < 313; i += candidate) {
                    isComposite[i] = candidate;
                }
                hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
                k[primeCounter++] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
            }
        }

        ascii += '\x80';
        while (ascii.length % 64 - 56) ascii += '\x00';
        for (var i = 0; i < ascii.length; i++) {
            var j = ascii.charCodeAt(i);
            if (j >> 8) return '';  /* ASCII-only; non-ASCII never matches a numeric code */
            words[i >> 2] |= j << ((3 - i) % 4) * 8;
        }
        words[words.length] = ((asciiBitLength / maxWord) | 0);
        words[words.length] = (asciiBitLength);

        for (var j = 0; j < words.length;) {
            var w = words.slice(j, j += 16);
            var oldHash = hash;
            hash = hash.slice(0, 8);

            for (var i = 0; i < 64; i++) {
                var w15 = w[i - 15], w2 = w[i - 2];
                var a = hash[0], e = hash[4];
                var temp1 = hash[7]
                    + (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25))
                    + ((e & hash[5]) ^ ((~e) & hash[6]))
                    + k[i]
                    + (w[i] = (i < 16) ? (w[i] | 0) : (
                            w[i - 16]
                            + (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3))
                            + w[i - 7]
                            + (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))
                        ) | 0
                    );
                var temp2 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22))
                    + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));

                hash = [(temp1 + temp2) | 0].concat(hash);
                hash[4] = (hash[4] + temp1) | 0;
            }

            for (var i = 0; i < 8; i++) {
                hash[i] = (hash[i] + oldHash[i]) | 0;
            }
        }

        for (var i = 0; i < 8; i++) {
            for (var j = 3; j + 1; j--) {
                var b = (hash[i] >> (j * 8)) & 255;
                result += ((b < 16) ? 0 : '') + b.toString(16);
            }
        }
        return result;
    }

    async function sha256Hex(str) {
        /* Prefer native Web Crypto when available (secure context). */
        if (window.crypto && window.crypto.subtle) {
            try {
                const buf = await crypto.subtle.digest(
                    'SHA-256',
                    new TextEncoder().encode(str)
                );
                return Array.from(new Uint8Array(buf))
                            .map(function (b) { return b.toString(16).padStart(2, '0'); })
                            .join('');
            } catch (e) { /* fall through to pure-JS implementation */ }
        }
        return sha256Fallback(str);
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
            /* Normalize: full-width digits (from Korean IME) -> ASCII, and strip any
               whitespace / zero-width characters that sneak in via copy-paste. */
            const val = (input.value || '')
                .replace(/[０-９]/g, function (d) {
                    return String.fromCharCode(d.charCodeAt(0) - 0xFEE0);
                })
                .replace(/[\s\u200B-\u200D\uFEFF]/g, '');
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
