// Oxi site analytics — Google Analytics 4 (gtag.js)
// Shared by all pages via <script src="analytics.js" defer>.
// Equivalent to the official gtag snippet for the ID below.
(function () {
  var GA_ID = 'G-87LWCWJ65R';
  if (!GA_ID || navigator.doNotTrack === '1') return;
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
