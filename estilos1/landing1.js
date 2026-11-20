/* Medición y enlaces de las páginas de aterrizaje */
(function () {
  var lp = document.body.getAttribute('data-landing') || '';
  var q = new URLSearchParams(window.location.search);
  var keep = new URLSearchParams();
  ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term', 'fbclid', 'gclid'].forEach(function (k) {
    if (q.get(k)) keep.set(k, q.get(k));
  });
  var extra = keep.toString();

  function meta(evento, params, personalizado) {
    try { if (window.fbq) { personalizado ? fbq('trackCustom', evento, params) : fbq('track', evento, params); } } catch (e) {}
  }
  function ga(evento, params) {
    try { if (window.gtag) gtag('event', evento, params); } catch (e) {}
  }

  document.addEventListener('DOMContentLoaded', function () {
    meta('ViewContent', { content_name: lp, content_category: 'landing' });

    // Botones "Cotizar": se les agregan los datos de la campaña (utm, fbclid) y se mide el clic
    document.querySelectorAll('a[data-cta]').forEach(function (a) {
      if (extra) a.href += (a.href.indexOf('?') > -1 ? '&' : '?') + extra;
      a.addEventListener('click', function () {
        var lugar = a.getAttribute('data-cta');
        meta('ClickCotizar', { landing: lp, ubicacion: lugar }, true);
        ga('cta_click', { landing: lp, ubicacion: lugar });
      });
    });

    // Botones de WhatsApp
    document.querySelectorAll('a[data-wa]').forEach(function (a) {
      a.addEventListener('click', function () {
        meta('Contact', { landing: lp });
        ga('whatsapp_click', { landing: lp });
      });
    });
  });
})();
