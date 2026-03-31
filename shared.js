/* ============================================================
   FLAIR NETWORK SYSTEMS — Shared JavaScript
   Premium Enterprise IT Infrastructure
   ============================================================ */

(function () {
  'use strict';

  /* ── Loader ── */
  function initLoader() {
    const loader = document.getElementById('loader');
    if (!loader) return;
    document.body.classList.add('loading');
    setTimeout(() => {
      loader.classList.add('hidden');
      document.body.classList.remove('loading');
    }, 1600);
  }

  /* ── Custom Cursor ── */
  function initCursor() {
    if (window.matchMedia('(hover: none)').matches) return;

    const dot = document.getElementById('cursor');
    const ring = document.getElementById('cursor-ring');
    if (!dot || !ring) return;

    let cx = -100, cy = -100, rx = -100, ry = -100;
    let raf;

    document.addEventListener('mousemove', e => {
      cx = e.clientX; cy = e.clientY;
    });

    function loop() {
      rx += (cx - rx) * 0.12;
      ry += (cy - ry) * 0.12;
      dot.style.left = cx + 'px';
      dot.style.top  = cy + 'px';
      ring.style.left = rx + 'px';
      ring.style.top  = ry + 'px';
      raf = requestAnimationFrame(loop);
    }
    loop();

    document.querySelectorAll('a, button, [data-hover], .glass-card, .btn, .nav-links a, #float-cta').forEach(el => {
      el.addEventListener('mouseenter', () => { dot.classList.add('hover'); ring.classList.add('hover'); });
      el.addEventListener('mouseleave', () => { dot.classList.remove('hover'); ring.classList.remove('hover'); });
    });
  }

  /* ── Scroll Progress ── */
  function initScrollProgress() {
    const bar = document.getElementById('scroll-progress');
    if (!bar) return;
    window.addEventListener('scroll', () => {
      const total = document.documentElement.scrollHeight - window.innerHeight;
      const pct = total > 0 ? (window.scrollY / total) * 100 : 0;
      bar.style.width = pct + '%';
    }, { passive: true });
  }

  /* ── Nav Scroll / Hamburger ── */
  function initNav() {
    const nav = document.querySelector('.nav');
    const hamburger = document.querySelector('.nav-hamburger');
    const mobileMenu = document.querySelector('.nav-mobile');

    if (nav) {
      window.addEventListener('scroll', () => {
        nav.classList.toggle('scrolled', window.scrollY > 40);
      }, { passive: true });
    }

    if (hamburger && mobileMenu) {
      hamburger.addEventListener('click', () => {
        hamburger.classList.toggle('open');
        mobileMenu.classList.toggle('open');
      });
    }

    // Active link
    const currentPage = location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.nav-links a, .nav-mobile a').forEach(a => {
      const href = a.getAttribute('href') || '';
      if (href === currentPage || (currentPage === '' && href === 'index.html')) {
        a.classList.add('active');
      }
    });
  }

  /* ── Reveal on Scroll (IntersectionObserver) ── */
  function initReveal() {
    const items = document.querySelectorAll('.reveal, .reveal-left, .reveal-right');
    if (!items.length) return;

    const isMobile = window.matchMedia('(max-width: 768px)').matches;

    const obs = new IntersectionObserver((entries) => {
      entries.forEach((entry, i) => {
        if (entry.isIntersecting) {
          const delay = entry.target.dataset.delay || 0;
          setTimeout(() => {
            entry.target.classList.add('revealed');
          }, isMobile ? 0 : Number(delay));
          obs.unobserve(entry.target);
        }
      });
    }, { threshold: 0.12 });

    items.forEach(el => obs.observe(el));
  }

  /* ── Counter Animation ── */
  function animateCounter(el) {
    const target = parseInt(el.dataset.target, 10);
    const suffix = el.dataset.suffix || '';
    const duration = 1800;
    const start = performance.now();

    function easeOut(t) { return 1 - Math.pow(1 - t, 3); }

    function tick(now) {
      const elapsed = now - start;
      const progress = Math.min(elapsed / duration, 1);
      const val = Math.round(easeOut(progress) * target);
      el.textContent = val.toLocaleString() + suffix;
      if (progress < 1) requestAnimationFrame(tick);
    }
    requestAnimationFrame(tick);
  }

  function initCounters() {
    const counters = document.querySelectorAll('[data-counter]');
    if (!counters.length) return;

    const obs = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          animateCounter(entry.target);
          obs.unobserve(entry.target);
        }
      });
    }, { threshold: 0.5 });

    counters.forEach(el => {
      el.dataset.target = el.dataset.counter;
      obs.observe(el);
    });
  }

  /* ── Typewriter ── */
  function initTypewriter() {
    const el = document.getElementById('typewriter');
    if (!el) return;

    const words = el.dataset.words ? JSON.parse(el.dataset.words) : [];
    if (!words.length) return;

    let wordIndex = 0, charIndex = 0, deleting = false;

    function type() {
      const word = words[wordIndex];
      if (!deleting) {
        el.textContent = word.slice(0, ++charIndex);
        if (charIndex === word.length) { deleting = true; setTimeout(type, 1800); return; }
      } else {
        el.textContent = word.slice(0, --charIndex);
        if (charIndex === 0) { deleting = false; wordIndex = (wordIndex + 1) % words.length; }
      }
      setTimeout(type, deleting ? 60 : 90);
    }
    type();
  }

  /* ── Floating CTA ── */
  function initFloatCTA() {
    const btn = document.getElementById('float-cta');
    if (!btn) return;
    btn.addEventListener('mouseenter', () => { btn.classList.add('hover'); });
    btn.addEventListener('mouseleave', () => { btn.classList.remove('hover'); });
  }

  /* ── Init All ── */
  document.addEventListener('DOMContentLoaded', () => {
    initLoader();
    initCursor();
    initScrollProgress();
    initNav();
    initReveal();
    initCounters();
    initTypewriter();
    initFloatCTA();
  });

})();
