(function () {
  var root = document.documentElement;
  var stored = localStorage.getItem("wplp-lang");
  document.querySelectorAll("[data-lang-link]").forEach(function (a) {
    a.addEventListener("click", function () {
      localStorage.setItem("wplp-lang", a.getAttribute("hreflang") || "");
    });
  });
  root.dataset.prefLang = stored || "";
})();
