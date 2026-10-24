/* Menú móvil, video de YouTube y visor de fotos */
document.addEventListener('DOMContentLoaded', function () {
  // Menú móvil
  var burger = document.querySelector('.burger');
  var menu = document.querySelector('.menu');
  if (burger && menu) {
    burger.addEventListener('click', function () {
      var abierto = menu.classList.toggle('open');
      burger.setAttribute('aria-expanded', abierto ? 'true' : 'false');
    });
    menu.addEventListener('click', function (e) {
      if (e.target.closest('a')) { menu.classList.remove('open'); burger.setAttribute('aria-expanded', 'false'); }
    });
  }

  // Video: se muestra solo si completás data-youtube="ID_DEL_VIDEO" en index.html
  var video = document.getElementById('video');
  if (video) {
    var id = (video.getAttribute('data-youtube') || '').trim();
    if (id) {
      var f = document.createElement('iframe');
      f.src = 'https://www.youtube-nocookie.com/embed/' + encodeURIComponent(id) + '?rel=0';
      f.title = 'Show de magia de Fabián Gi';
      f.loading = 'lazy';
      f.allow = 'accelerometer; encrypted-media; picture-in-picture';
      f.allowFullscreen = true;
      video.appendChild(f);
      video.hidden = false;
    }
  }

  // Filtros del portfolio
  var botones = document.querySelectorAll('.filter');
  if (botones.length) {
    botones.forEach(function (b) {
      b.addEventListener('click', function () {
        botones.forEach(function (x) { x.classList.remove('is-active'); x.setAttribute('aria-pressed', 'false'); });
        b.classList.add('is-active');
        b.setAttribute('aria-pressed', 'true');
        var f = b.getAttribute('data-filter');
        document.querySelectorAll('.gallery a').forEach(function (a) {
          a.hidden = !(f === 'todos' || a.getAttribute('data-cat') === f);
        });
      });
    });
  }

  // Visor de fotos
  var links = Array.prototype.slice.call(document.querySelectorAll('[data-lightbox]'));
  if (links.length) {
    var ov = document.createElement('div');
    ov.className = 'lightbox';
    ov.hidden = true;
    ov.innerHTML = '<button type="button" aria-label="Cerrar">×</button><img alt="">';
    document.body.appendChild(ov);
    var img = ov.querySelector('img');
    var cerrar = function () { ov.hidden = true; img.removeAttribute('src'); };
    links.forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        var mini = a.querySelector('img');
        img.src = a.getAttribute('href');
        img.alt = mini ? mini.alt : '';
        ov.hidden = false;
      });
    });
    ov.addEventListener('click', cerrar);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') cerrar(); });
  }
});
