(function () {
  'use strict';

  const analyticsId = 'G-Y5VPB1HNPM';
  const consentKey = 'xtelefonja_cookie_consent_v1';
  let analyticsLoaded = false;

  function readConsent() {
    try {
      return localStorage.getItem(consentKey);
    } catch (error) {
      return null;
    }
  }

  function saveConsent(value) {
    try {
      localStorage.setItem(consentKey, value);
    } catch (error) {
      // The preference cannot be persisted in private/restricted browser modes.
    }
  }

  function loadAnalytics() {
    if (analyticsLoaded || document.querySelector('script[data-xtelefonja-ga]')) return;
    analyticsLoaded = true;

    window.dataLayer = window.dataLayer || [];
    window.gtag = window.gtag || function () {
      window.dataLayer.push(arguments);
    };
    window.gtag('consent', 'default', { analytics_storage: 'granted' });
    window.gtag('js', new Date());
    window.gtag('config', analyticsId, { anonymize_ip: true });

    const script = document.createElement('script');
    script.async = true;
    script.src = 'https://www.googletagmanager.com/gtag/js?id=' + analyticsId;
    script.dataset.xtelefonjaGa = 'true';
    document.head.appendChild(script);
  }

  function removeAnalyticsCookies() {
    const cookieNames = document.cookie
      .split(';')
      .map((item) => item.split('=')[0].trim())
      .filter((name) => name === '_ga' || name.indexOf('_ga_') === 0);
    const hostParts = location.hostname.split('.');
    const domains = [location.hostname];
    if (hostParts.length > 1) domains.push('.' + hostParts.slice(-2).join('.'));

    cookieNames.forEach((name) => {
      document.cookie = name + '=; Max-Age=0; path=/; SameSite=Lax';
      domains.forEach((domain) => {
        document.cookie = name + '=; Max-Age=0; path=/; domain=' + domain + '; SameSite=Lax';
      });
    });
  }

  function closeBanner() {
    const banner = document.getElementById('cookieBanner');
    if (banner) banner.remove();
  }

  function setConsent(value) {
    saveConsent(value);
    closeBanner();
    if (value === 'accepted') {
      loadAnalytics();
      return;
    }
    if (typeof window.gtag === 'function') {
      window.gtag('consent', 'update', { analytics_storage: 'denied' });
    }
    removeAnalyticsCookies();
  }

  function showBanner() {
    if (document.getElementById('cookieBanner')) return;
    const banner = document.createElement('section');
    banner.id = 'cookieBanner';
    banner.className = 'cookie-banner';
    banner.setAttribute('role', 'dialog');
    banner.setAttribute('aria-modal', 'true');
    banner.setAttribute('aria-label', 'Preferenze cookie');
    banner.innerHTML =
      '<button class="cookie-close" type="button" aria-label="Continua senza cookie statistici">&times;</button>' +
      '<div class="cookie-copy"><strong>Cookie e privacy</strong>' +
      '<p>Usiamo cookie tecnici necessari e, solo con il tuo consenso, Google Analytics per statistiche anonime sul sito. Puoi rifiutare senza limitazioni.</p>' +
      '<a href="privacy-cookie.html">Privacy e Cookie Policy</a></div>' +
      '<div class="cookie-actions"><button class="cookie-button" type="button" data-cookie-choice="rejected">Rifiuta</button>' +
      '<button class="cookie-button cookie-accept" type="button" data-cookie-choice="accepted">Accetta statistiche</button></div>';
    document.body.appendChild(banner);
    banner.querySelector('.cookie-close').addEventListener('click', () => setConsent('rejected'));
    banner.querySelectorAll('[data-cookie-choice]').forEach((button) => {
      button.addEventListener('click', () => setConsent(button.dataset.cookieChoice));
    });
    banner.querySelector('.cookie-button').focus();
  }

  function addCookieStyles() {
    if (document.getElementById('cookieStyles')) return;
    const style = document.createElement('style');
    style.id = 'cookieStyles';
    style.textContent =
      '.cookie-banner{position:fixed;left:16px;right:16px;bottom:16px;z-index:20000;max-width:960px;margin:auto;padding:18px 52px 18px 18px;background:#fff;color:#0f172a;border:1px solid #cbd5e1;border-radius:14px;box-shadow:0 18px 50px rgba(15,23,42,.28);display:flex;gap:18px;align-items:center}' +
      '.cookie-copy{flex:1}.cookie-copy p{margin:6px 0;font-size:14px;line-height:1.45}.cookie-copy a{color:#0f766e;font-weight:700}' +
      '.cookie-actions{display:flex;gap:10px;flex-wrap:wrap}.cookie-button,.cookie-close{border:2px solid #0f766e;background:#fff;color:#0f766e;border-radius:10px;padding:9px 14px;font-weight:800;cursor:pointer}' +
      '.cookie-accept{background:#0f766e;color:#fff}.cookie-close{position:absolute;right:10px;top:10px;width:36px;height:36px;padding:0;font-size:24px;line-height:1}' +
      '.footer-links{display:flex;gap:12px;flex-wrap:wrap;margin-top:8px}.footer-links a,.cookie-settings{color:#0f766e;background:none;border:0;padding:0;font:inherit;text-decoration:underline;cursor:pointer}' +
      '.map-consent{min-height:300px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:24px;text-align:center;background:#eef2f7}.map-consent p{margin:0;color:#475569}.map-consent button{border:0;background:#0f766e;color:#fff;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer}' +
      '.product-filter{display:flex;gap:10px;flex-wrap:wrap;margin:18px 0}.product-filter button{border:2px solid #0f766e;background:#fff;color:#0f766e;border-radius:999px;padding:8px 14px;font-weight:800;cursor:pointer}.product-filter button.active{background:#0f766e;color:#fff}.product-hidden{display:none!important}' +
      '@media(max-width:720px){.cookie-banner{left:10px;right:10px;bottom:10px;display:block;padding:16px 48px 16px 16px}.cookie-actions{margin-top:12px}.cookie-button{flex:1}.map-consent{min-height:230px}}';
    document.head.appendChild(style);
  }

  function setupCookieControls() {
    document.querySelectorAll('[data-cookie-settings]').forEach((button) => {
      button.addEventListener('click', showBanner);
    });
  }

  function setupMapConsent() {
    const iframe = document.querySelector('iframe[data-map-src]');
    if (!iframe) return;
    const container = iframe.parentElement;
    const consent = document.createElement('div');
    consent.className = 'map-consent';
    consent.innerHTML = '<strong>Google Maps è bloccato per proteggere la tua privacy.</strong><p>Caricalo solo se desideri visualizzare la mappa interattiva.</p><button type="button">Carica Google Maps</button>';
    iframe.hidden = true;
    container.insertBefore(consent, iframe);
    consent.querySelector('button').addEventListener('click', () => {
      iframe.src = iframe.dataset.mapSrc;
      iframe.hidden = false;
      consent.remove();
    });
  }

  function setupProductLinks() {
    document.querySelectorAll('.product-image').forEach((image) => {
      image.loading = 'lazy';
      image.decoding = 'async';
    });

    document.querySelectorAll('.card').forEach((card) => {
      const link = card.querySelector('a.btn-card[href^="https://wa.me/393381846260"]');
      const title = card.querySelector('.product-title');
      if (!link || !title) return;
      const brand = title.querySelector('.product-brand')?.textContent.trim();
      const model = title.querySelector('.product-model')?.textContent.trim();
      const productName = [brand, model].filter(Boolean).join(' ') || title.textContent.replace(/\s+/g, ' ').trim();
      const price = card.querySelector('.product-price');
      const isSoldOut = card.classList.contains('sold-out');
      const message = isSoldOut
        ? 'Vorrei essere avvisato quando torna disponibile: ' + productName
        : 'Vorrei informazioni su ' + productName + (price ? ' - ' + price.textContent.trim() : '');
      link.href = 'https://wa.me/393381846260?text=' + encodeURIComponent(message);
    });
  }

  function setupProductFilter() {
    const filter = document.querySelector('[data-product-filter]');
    if (!filter) return;
    const cards = Array.from(document.querySelectorAll('.section-block .card'));

    function applyFilter(value) {
      cards.forEach((card) => {
        const isSoldOut = card.classList.contains('sold-out');
        const hide = (value === 'available' && isSoldOut) || (value === 'sold' && !isSoldOut);
        card.classList.toggle('product-hidden', hide);
      });
      filter.querySelectorAll('button').forEach((button) => {
        const active = button.dataset.filter === value;
        button.classList.toggle('active', active);
        button.setAttribute('aria-pressed', active ? 'true' : 'false');
      });
    }

    filter.addEventListener('click', (event) => {
      const button = event.target.closest('button[data-filter]');
      if (button) applyFilter(button.dataset.filter);
    });
    applyFilter('available');
  }

  function setupExternalLinks() {
    document.querySelectorAll('a[target="_blank"]').forEach((link) => {
      link.rel = 'noopener noreferrer';
    });
  }

  document.addEventListener('DOMContentLoaded', () => {
    addCookieStyles();
    setupCookieControls();
    setupMapConsent();
    setupProductLinks();
    setupProductFilter();
    setupExternalLinks();
    if (readConsent() === 'accepted') loadAnalytics();
    else if (readConsent() !== 'rejected') showBanner();
  });

  window.XtelefonjaCookies = { open: showBanner };
}());
