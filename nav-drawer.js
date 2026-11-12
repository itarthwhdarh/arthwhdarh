/* ═══════════════════════════════════════════════════════════════════
   NAV-DRAWER.JS — جمعية إرث وحضارة بالقريات
   Unified Minimal-Luxury Drawer Navigation Controller
   - RTL slide-in drawer with smooth cubic-bezier motion
   - Single-open Accordion groups with grid-row height transition
   - Focus management, Focus Trap, Escape & Backdrop dismissal
   - Scroll locking without orphaned overflow:hidden
   - Active link detection & sticky header scroll state
   ═══════════════════════════════════════════════════════════════════ */

(function () {
    'use strict';

    if (window.__erthNavDrawerInit) return;
    window.__erthNavDrawerInit = true;

    var CLOSE_DURATION_MS = 380;
    var NAV_LEAVE_DELAY_MS = 240;
    var TRANSITION_KEY = 'erthNavTransitionTs';

    function prefersReducedMotion() {
        return !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
    }

    function supportsViewTransitions() {
        return 'onpageswap' in window;
    }

    function clearPageTransitionClasses() {
        document.documentElement.classList.remove('erth-page-leaving', 'erth-page-entering');
    }

    function applyFallbackEnterTransition() {
        if (prefersReducedMotion() || supportsViewTransitions()) return;
        try {
            var raw = sessionStorage.getItem(TRANSITION_KEY);
            sessionStorage.removeItem(TRANSITION_KEY);
            if (!raw) return;
            var ts = parseInt(raw, 10);
            if (!ts || Date.now() - ts > 3000) return;

            var root = document.documentElement;
            root.classList.add('erth-page-entering');

            var done = false;
            var cleanup = function () {
                if (done) return;
                done = true;
                root.classList.remove('erth-page-entering');
            };
            var wrapper = document.querySelector('.site-wrapper');
            if (wrapper) {
                wrapper.addEventListener('animationend', cleanup, { once: true });
            }
            setTimeout(cleanup, 450);
        } catch (err) {
            clearPageTransitionClasses();
        }
    }

    applyFallbackEnterTransition();

    function init() {
        var trigger = document.getElementById('navDrawerTrigger');
        var overlay = document.getElementById('navDrawerOverlay');
        var drawer  = document.getElementById('navDrawer');
        var closeBtn = document.getElementById('navDrawerClose');
        var mainNav = document.getElementById('mainNav');

        // Smart fixed header: large & transparent at top -> compact glass ivory (#F8F6F2) on scroll
        var applyNavScrollState = null;
        if (mainNav && !mainNav.dataset.scrollBound) {
            mainNav.dataset.scrollBound = '1';

            // Keep header directly under <body> so page-transition transforms on .site-wrapper never break fixed positioning
            if (mainNav.parentElement !== document.body) {
                document.body.insertBefore(mainNav, document.body.firstChild);
            }

            var isHeaderScrolled = null;
            applyNavScrollState = function (shouldScroll) {
                var next = Boolean(shouldScroll);
                if (next === isHeaderScrolled) return;
                isHeaderScrolled = next;
                mainNav.classList.toggle('scrolled', next);
            };

            applyNavScrollState(window.scrollY > 24);

            if ('IntersectionObserver' in window) {
                var sentinel = document.querySelector('.nav-scroll-sentinel');
                if (!sentinel) {
                    sentinel = document.createElement('div');
                    sentinel.className = 'nav-scroll-sentinel';
                    sentinel.setAttribute('aria-hidden', 'true');
                    document.body.insertBefore(sentinel, document.body.firstChild);
                }

                var headerObserver = new IntersectionObserver(function (entries) {
                    if (!entries || !entries.length) return;
                    if (document.documentElement.classList.contains('nav-drawer-lock')) return;
                    applyNavScrollState(!entries[0].isIntersecting);
                }, { threshold: 0 });

                headerObserver.observe(sentinel);
            } else {
                var onScrollFallback = function () {
                    if (document.documentElement.classList.contains('nav-drawer-lock')) return;
                    applyNavScrollState(window.scrollY > 24);
                };
                window.addEventListener('scroll', onScrollFallback, { passive: true });
            }
        }

        if (!trigger || !overlay || !drawer) return;

        // Ensure overlay & drawer sit directly on <body> so no parent stacking context clips z-index
        if (overlay.parentElement !== document.body) {
            document.body.appendChild(overlay);
        }
        if (drawer.parentElement !== document.body) {
            document.body.appendChild(drawer);
        }

        var isOpen = false;
        var isNavigating = false;
        var closeTimer = null;
        var navTimer = null;
        var lastFocusedEl = null;

        var groups = Array.prototype.slice.call(drawer.querySelectorAll('.nav-drawer-group'));
        var accordionBtns = Array.prototype.slice.call(drawer.querySelectorAll('.nav-drawer-accordion-btn'));
        var allLinks = Array.prototype.slice.call(drawer.querySelectorAll('a[href]'));

        /* ── Normalize paths & mark active link ── */
        function normalizePath(p) {
            return String(p || '')
                .toLowerCase()
                .replace(/index\.html$/, '')
                .replace(/^\/+|\/+$/g, '');
        }

        function syncActiveState() {
            var currentPath = normalizePath(window.location.pathname);

            allLinks.forEach(function (link) {
                var href = normalizePath(link.getAttribute('href'));
                var isMatch = (href === currentPath) || link.classList.contains('active');

                if (isMatch) {
                    link.classList.add('active');
                    link.setAttribute('aria-current', 'page');
                    var parentGroup = link.closest('.nav-drawer-group');
                    if (parentGroup) {
                        var btn = parentGroup.querySelector('.nav-drawer-accordion-btn');
                        if (btn) btn.classList.add('has-active');
                    }
                } else {
                    link.classList.remove('active');
                    link.removeAttribute('aria-current');
                }
            });
        }

        /* ── Accordion helpers ── */
        function setGroupExpanded(group, expand) {
            var btn = group.querySelector('.nav-drawer-accordion-btn');
            var panel = group.querySelector('.nav-drawer-accordion-panel');
            if (!btn || !panel) return;

            var subLinks = panel.querySelectorAll('a[href]');
            if (expand) {
                group.classList.add('is-expanded');
                btn.setAttribute('aria-expanded', 'true');
                panel.setAttribute('aria-hidden', 'false');
                subLinks.forEach(function (a) { a.removeAttribute('tabindex'); });
            } else {
                group.classList.remove('is-expanded');
                btn.setAttribute('aria-expanded', 'false');
                panel.setAttribute('aria-hidden', 'true');
                subLinks.forEach(function (a) { a.setAttribute('tabindex', '-1'); });
            }
        }

        function resetAccordions(exceptGroup) {
            groups.forEach(function (group) {
                if (group !== exceptGroup) {
                    setGroupExpanded(group, false);
                }
            });
        }

        // Initialize all accordions as collapsed and sub-links unfocusable while collapsed
        resetAccordions(null);

        accordionBtns.forEach(function (btn) {
            btn.addEventListener('click', function () {
                var group = btn.closest('.nav-drawer-group');
                if (!group) return;
                var willExpand = !group.classList.contains('is-expanded');
                resetAccordions(willExpand ? group : null);
                setGroupExpanded(group, willExpand);
            });
        });

        /* ── Non-mutating Scroll Lock (zero reflow, zero scrollbar width/padding shift) ── */
        var scrollLocked = false;
        var lockedScrollX = 0;
        var lockedScrollY = 0;
        var touchStartY = 0;

        function canScrollDrawerBody(target, deltaY) {
            var scrollable = target && target.closest ? target.closest('.nav-drawer-body') : null;
            if (!scrollable) return false;
            if (scrollable.scrollHeight <= scrollable.clientHeight + 1) return false;
            if (deltaY < 0) return scrollable.scrollTop > 0;
            if (deltaY > 0) return scrollable.scrollTop + scrollable.clientHeight < scrollable.scrollHeight - 1;
            return true;
        }

        function onWheelLock(e) {
            if (!scrollLocked) return;
            if (!canScrollDrawerBody(e.target, e.deltaY)) {
                e.preventDefault();
            }
        }

        function onTouchStartLock(e) {
            if (!scrollLocked) return;
            if (e.touches && e.touches.length === 1) {
                touchStartY = e.touches[0].clientY;
            }
        }

        function onTouchMoveLock(e) {
            if (!scrollLocked) return;
            if (e.touches && e.touches.length > 1) return;
            var currentY = (e.touches && e.touches[0]) ? e.touches[0].clientY : touchStartY;
            var deltaY = touchStartY - currentY;
            if (!canScrollDrawerBody(e.target, deltaY)) {
                e.preventDefault();
            }
        }

        function onScrollPin() {
            if (!scrollLocked) return;
            if (window.scrollX !== lockedScrollX || window.scrollY !== lockedScrollY) {
                window.scrollTo({ left: lockedScrollX, top: lockedScrollY, behavior: 'auto' });
            }
        }

        function lockScroll() {
            if (scrollLocked) return;
            lockedScrollX = window.scrollX || window.pageXOffset || 0;
            lockedScrollY = window.scrollY || window.pageYOffset || 0;
            scrollLocked = true;
            document.documentElement.classList.add('nav-drawer-lock');
            window.addEventListener('wheel', onWheelLock, { passive: false, capture: true });
            window.addEventListener('touchstart', onTouchStartLock, { passive: true, capture: true });
            window.addEventListener('touchmove', onTouchMoveLock, { passive: false, capture: true });
            window.addEventListener('scroll', onScrollPin, { passive: true });
        }

        function unlockScroll() {
            scrollLocked = false;
            document.documentElement.classList.remove('nav-drawer-lock');
            document.body.classList.remove('nav-drawer-lock');
            document.documentElement.style.removeProperty('--drawer-scrollbar-gap');
            window.removeEventListener('wheel', onWheelLock, { capture: true });
            window.removeEventListener('touchstart', onTouchStartLock, { capture: true });
            window.removeEventListener('touchmove', onTouchMoveLock, { capture: true });
            window.removeEventListener('scroll', onScrollPin);
        }

        /* ── Focusable elements inside drawer (visible & enabled only) ── */
        function getFocusableElements() {
            var candidates = drawer.querySelectorAll(
                'button:not([disabled]), a[href]:not([tabindex="-1"]), [tabindex]:not([tabindex="-1"])'
            );
            return Array.prototype.filter.call(candidates, function (el) {
                if (el.closest('[aria-hidden="true"]')) return false;
                var style = window.getComputedStyle(el);
                if (style.display === 'none' || style.visibility === 'hidden') return false;
                var li = el.closest('li');
                if (li && window.getComputedStyle(li).display === 'none') return false;
                return true;
            });
        }

        /* ── Open / Close Drawer ── */
        function openDrawer() {
            if (closeTimer) {
                clearTimeout(closeTimer);
                closeTimer = null;
            }
            if (isOpen) return;

            isOpen = true;
            lastFocusedEl = document.activeElement;

            lockScroll();
            drawer.setAttribute('aria-hidden', 'false');
            overlay.setAttribute('aria-hidden', 'false');
            trigger.setAttribute('aria-expanded', 'true');

            // Force reflow so rapid close -> open always animates smoothly
            void drawer.offsetWidth;

            overlay.classList.add('is-open');
            drawer.classList.add('is-open');

            // Focus close button without scrolling
            window.requestAnimationFrame(function () {
                if (!isOpen) return;
                var targetFocus = closeBtn || getFocusableElements()[0] || drawer;
                if (targetFocus && typeof targetFocus.focus === 'function') {
                    targetFocus.focus({ preventScroll: true });
                }
            });
        }

        function closeDrawer(options) {
            var opts = options || {};
            if (!isOpen && !drawer.classList.contains('is-open')) {
                unlockScroll();
                return;
            }

            isOpen = false;
            if (closeTimer) {
                clearTimeout(closeTimer);
                closeTimer = null;
            }

            overlay.classList.remove('is-open');
            drawer.classList.remove('is-open');
            drawer.setAttribute('aria-hidden', 'true');
            overlay.setAttribute('aria-hidden', 'true');
            trigger.setAttribute('aria-expanded', 'false');

            // Always restore scroll cleanly
            unlockScroll();

            // Restore focus before hiding completes
            if (!opts.skipFocusRestore) {
                var focusBack = (lastFocusedEl && typeof lastFocusedEl.focus === 'function' && document.contains(lastFocusedEl))
                    ? lastFocusedEl
                    : trigger;
                if (focusBack && typeof focusBack.focus === 'function') {
                    focusBack.focus({ preventScroll: true });
                }
            }
            lastFocusedEl = null;

            // Reset accordions after slide-out transition completes (or immediately under reduced motion)
            if (prefersReducedMotion()) {
                resetAccordions(null);
            } else {
                closeTimer = setTimeout(function () {
                    if (!isOpen) {
                        resetAccordions(null);
                    }
                    closeTimer = null;
                }, CLOSE_DURATION_MS);
            }
        }

        function handleDrawerLinkClick(e, link) {
            // Let modifier keys, middle clicks, new tabs, and downloads behave natively
            if (e.defaultPrevented || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey ||
                (e.button && e.button !== 0) || link.target === '_blank' || link.hasAttribute('download')) {
                closeDrawer({ skipFocusRestore: true });
                return;
            }

            var rawHref = link.getAttribute('href') || '';
            var targetUrl;
            try {
                targetUrl = new URL(link.href, window.location.href);
            } catch (err) {
                closeDrawer({ skipFocusRestore: true });
                return;
            }

            // External origin or special protocol (mailto:, tel:, etc.)
            if (targetUrl.origin !== window.location.origin || !/^https?:$/.test(targetUrl.protocol)) {
                closeDrawer({ skipFocusRestore: true });
                return;
            }

            var samePath = normalizePath(targetUrl.pathname) === normalizePath(window.location.pathname) &&
                           targetUrl.search === window.location.search;

            // In-page anchor link on current page: close drawer cleanly then scroll smoothly
            if (samePath && targetUrl.hash) {
                var targetId = decodeURIComponent(targetUrl.hash.slice(1));
                var hashTarget = targetId ? document.getElementById(targetId) : null;
                if (hashTarget) {
                    e.preventDefault();
                    closeDrawer({ skipFocusRestore: true });
                    var delay = prefersReducedMotion() ? 0 : 210;
                    setTimeout(function () {
                        hashTarget.scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
                        try {
                            if (window.history && typeof window.history.pushState === 'function') {
                                window.history.pushState(null, '', targetUrl.hash);
                            }
                        } catch (_) {}
                    }, delay);
                    return;
                }
            }

            // Clicking active current route without a hash simply dismisses the drawer
            if (samePath && !targetUrl.hash) {
                e.preventDefault();
                closeDrawer({ skipFocusRestore: true });
                return;
            }

            // Under reduced motion: close drawer and navigate immediately without artificial transitions
            if (prefersReducedMotion()) {
                closeDrawer({ skipFocusRestore: true });
                return;
            }

            // Coordinated cross-page transition: smooth drawer close -> View Transitions API or CSS fallback
            e.preventDefault();
            if (isNavigating) return;
            isNavigating = true;

            closeDrawer({ skipFocusRestore: true });

            if (!supportsViewTransitions()) {
                document.documentElement.classList.add('erth-page-leaving');
                try {
                    sessionStorage.setItem(TRANSITION_KEY, String(Date.now()));
                } catch (_) {}
            }

            navTimer = setTimeout(function () {
                navTimer = null;
                window.location.assign(rawHref.indexOf('#') === 0 ? targetUrl.href : link.href);
            }, NAV_LEAVE_DELAY_MS);
        }

        /* ── Event Bindings ── */
        trigger.addEventListener('click', function (e) {
            e.preventDefault();
            if (isOpen) closeDrawer();
            else openDrawer();
        });

        if (closeBtn) {
            closeBtn.addEventListener('click', function () {
                closeDrawer();
            });
        }

        overlay.addEventListener('click', function () {
            closeDrawer();
        });

        allLinks.forEach(function (link) {
            link.addEventListener('click', function (e) {
                handleDrawerLinkClick(e, link);
            });
        });

        document.addEventListener('keydown', function (e) {
            if (!isOpen) return;

            if (e.key === 'Escape') {
                e.preventDefault();
                closeDrawer();
                return;
            }

            if (e.key === 'PageUp' || e.key === 'PageDown' || e.key === 'Home' || e.key === 'End') {
                if (!drawer.contains(e.target) || !e.target.closest('.nav-drawer-body')) {
                    e.preventDefault();
                    return;
                }
            }

            if (e.key === ' ' && !/^(BUTTON|A|INPUT|TEXTAREA|SELECT)$/i.test((e.target && e.target.tagName) || '')) {
                e.preventDefault();
                return;
            }

            // Focus Trap inside Drawer
            if (e.key === 'Tab') {
                var focusables = getFocusableElements();
                if (!focusables.length) {
                    e.preventDefault();
                    return;
                }
                var first = focusables[0];
                var last  = focusables[focusables.length - 1];

                if (e.shiftKey) {
                    if (document.activeElement === first || !drawer.contains(document.activeElement)) {
                        e.preventDefault();
                        last.focus();
                    }
                } else {
                    if (document.activeElement === last || !drawer.contains(document.activeElement)) {
                        e.preventDefault();
                        first.focus();
                    }
                }
            }
        });

        // Guard against browser back/forward cache keeping drawer or transition classes stuck
        window.addEventListener('pageshow', function (e) {
            isNavigating = false;
            if (navTimer) {
                clearTimeout(navTimer);
                navTimer = null;
            }
            if (e.persisted) {
                clearPageTransitionClasses();
            }
            if (e.persisted || isOpen) {
                closeDrawer({ skipFocusRestore: true });
            }
            if (applyNavScrollState) {
                applyNavScrollState(window.scrollY > 24);
            }
        });

        window.addEventListener('pagehide', function () {
            isNavigating = false;
            if (navTimer) {
                clearTimeout(navTimer);
                navTimer = null;
            }
            document.documentElement.classList.remove('erth-page-leaving');
            unlockScroll();
        });

        syncActiveState();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
