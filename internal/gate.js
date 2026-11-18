/* ===== gate.js — the shared lock for every /internal/ page =================
   Added 2026-11-18. One password for the whole internal site, asked once per
   browser session.

   WHAT THIS IS: a privacy screen. It keeps the lab's pages out of casual view
   on a shared or unattended machine, and makes a forwarded link useless to
   whoever receives it.

   WHAT IT IS NOT: security. These pages are served by GitHub Pages out of a
   PUBLIC repository, so the HTML is already on the wire before any of this
   runs — `curl`, view-source or JavaScript-off walks straight past it. It also
   cannot cover internal/methods_files/ at all: those are 109 direct file
   downloads and no script ever executes for them. What actually protects lab
   DATA is that it lives in the private lab-internal / lab-reagents repos and
   is fetched with a token held in the reader's own browser.

   WHY A REDIRECT rather than an overlay: every data page boots itself on parse
   (reagents.html's `(async function boot()` is the clearest case) and would pop
   its GitHub-token prompt behind a merely-hidden page. Navigating away in <head>
   stops the document loading before any page script runs at all.

   The password lives in gate.html, XOR-obfuscated, so it is not greppable in a
   public repository. Same trick tasks.html has used since September.
   ========================================================================= */
(function () {
    'use strict';
    var FLAG = 'yeh-internal-unlocked';

    /* Storage unavailable (private window, blocked site data) means the flag
       can never be written, and failing CLOSED there would redirect to the gate
       forever. Fail open: a lock that bricks the lab out of its own site is
       worse than one that is already bypassable by design. */
    try {
        if (sessionStorage.getItem(FLAG) === '1') return;
    } catch (e) {
        return;
    }

    var here = location.pathname.replace(/^.*\//, '') || 'index.html';
    if (here === 'gate.html') return;          // never gate the gate

    location.replace('gate.html?r='
        + encodeURIComponent(here + location.search + location.hash));
})();
