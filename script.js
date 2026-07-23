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
  if (!caps.length || !("IntersectionObserver" in window)) {
    caps.forEach((el) => el.classList.add("is-visible"));
    return;
  }

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.2 }
  );

  caps.forEach((el) => observer.observe(el));
})();
