(() => {
  const weightEl = document.querySelector("[data-weight]");
  if (weightEl) {
    const values = [42380, 42115, 41890, 42540, 42205];
    let i = 0;
    setInterval(() => {
      i = (i + 1) % values.length;
      weightEl.textContent = values[i].toLocaleString("ru-RU");
    }, 2800);
  }

  const caps = document.querySelectorAll(".cap");
  if (!caps.length) return;

  const reveal = (el) => el.classList.add("is-visible");

  if (!("IntersectionObserver" in window)) {
    caps.forEach(reveal);
    return;
  }

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          reveal(entry.target);
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.12, rootMargin: "0px 0px 80px 0px" }
  );

  caps.forEach((el) => {
    const rect = el.getBoundingClientRect();
    if (rect.top < window.innerHeight * 0.92) reveal(el);
    else observer.observe(el);
  });
})();
