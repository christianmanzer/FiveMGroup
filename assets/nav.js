// Mobile menu: the three header links collapse behind a menu button on small screens.
(function () {
  var btn = document.querySelector(".nav-toggle");
  var nav = document.getElementById("site-nav");
  if (!btn || !nav) return;

  function setOpen(open) {
    nav.classList.toggle("open", open);
    btn.setAttribute("aria-expanded", open ? "true" : "false");
  }

  btn.addEventListener("click", function (e) {
    e.stopPropagation();
    setOpen(!nav.classList.contains("open"));
  });
  document.addEventListener("click", function (e) {
    if (!nav.contains(e.target)) setOpen(false);
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && nav.classList.contains("open")) { setOpen(false); btn.focus(); }
  });
  window.matchMedia("(min-width: 36.01rem)").addEventListener("change", function () { setOpen(false); });
})();
