const root = document.documentElement;
const themeToggle = document.getElementById('themeToggle');
const savedTheme = localStorage.getItem('theme');
if (savedTheme) root.dataset.theme = savedTheme;

themeToggle?.addEventListener('click', () => {
  const next = root.dataset.theme === 'dark' ? 'light' : 'dark';
  root.dataset.theme = next;
  localStorage.setItem('theme', next);
});

const topbar = document.querySelector('.topbar');
if (topbar && themeToggle) {
  const topbarTools = document.createElement('div');
  topbarTools.className = 'topbar-tools';

  const cabinet = document.createElement('button');
  cabinet.className = 'power-cabinet';
  cabinet.type = 'button';
  cabinet.title = 'Проверить напряжение';
  cabinet.setAttribute('aria-label', 'Проверить напряжение на электрическом шкафу');
  cabinet.innerHTML = `
    <span class="cabinet-shell">
      <span class="cabinet-panel cabinet-left"></span>
      <span class="cabinet-panel cabinet-right"></span>
      <span class="cabinet-warning">⚡</span>
      <span class="cabinet-lamp cabinet-lamp-green"></span>
      <span class="cabinet-lamp cabinet-lamp-red"></span>
      <span class="cabinet-arc cabinet-arc-one"></span>
      <span class="cabinet-arc cabinet-arc-two"></span>
    </span>
    <span class="cabinet-burst cabinet-burst-one">⚡</span>
    <span class="cabinet-burst cabinet-burst-two">⚡</span>
    <span class="cabinet-burst cabinet-burst-three">⚡</span>
  `;

  topbar.insertBefore(topbarTools, themeToggle);
  topbarTools.append(cabinet, themeToggle);

  let cabinetZapTimer;
  cabinet.addEventListener('click', () => {
    cabinet.classList.remove('cabinet-zap');
    void cabinet.offsetWidth;
    cabinet.classList.add('cabinet-zap');
    clearTimeout(cabinetZapTimer);
    cabinetZapTimer = setTimeout(() => cabinet.classList.remove('cabinet-zap'), 900);
  });
}

const tabs = document.querySelectorAll('.tab');
const panels = document.querySelectorAll('.tab-panel');
tabs.forEach((tab) => {
  tab.addEventListener('click', () => {
    const id = tab.dataset.tab;
    tabs.forEach((item) => item.classList.toggle('active', item === tab));
    panels.forEach((panel) => panel.classList.toggle('active', panel.dataset.panel === id));
  });
});

const toast = document.getElementById('toast');
let toastTimer;
async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch (error) {
    const area = document.createElement('textarea');
    area.value = text;
    area.style.position = 'fixed';
    area.style.left = '-9999px';
    document.body.appendChild(area);
    area.focus();
    area.select();
    document.execCommand('copy');
    area.remove();
  }
  toast?.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast?.classList.remove('show'), 1600);
}

document.querySelectorAll('.copy-btn').forEach((button) => {
  button.addEventListener('click', () => {
    copyText(button.dataset.copy || button.closest('.code-block')?.querySelector('code')?.textContent || '');
    const old = button.textContent;
    button.textContent = 'Скопировано';
    setTimeout(() => { button.textContent = old; }, 1200);
  });
});

const screenshotImages = document.querySelectorAll('.wiki-screenshot img');
if (screenshotImages.length) {
  const lightbox = document.createElement('div');
  lightbox.className = 'image-lightbox';
  lightbox.setAttribute('role', 'dialog');
  lightbox.setAttribute('aria-modal', 'true');
  lightbox.setAttribute('aria-label', 'Просмотр скриншота');
  lightbox.innerHTML = `
    <div class="image-lightbox-inner">
      <button class="image-lightbox-close" type="button" aria-label="Закрыть">×</button>
      <img alt="" />
      <div class="image-lightbox-caption"></div>
    </div>
  `;
  document.body.appendChild(lightbox);

  const lightboxImage = lightbox.querySelector('img');
  const lightboxCaption = lightbox.querySelector('.image-lightbox-caption');
  const closeButton = lightbox.querySelector('.image-lightbox-close');
  let lastFocusedImage;

  function closeLightbox() {
    lightbox.classList.remove('open');
    document.body.classList.remove('lightbox-open');
    lightboxImage.removeAttribute('src');
    lastFocusedImage?.focus();
  }

  function openLightbox(image) {
    lastFocusedImage = image;
    lightboxImage.src = image.currentSrc || image.src;
    lightboxImage.alt = image.alt || '';
    lightboxCaption.textContent = image.closest('figure')?.querySelector('figcaption')?.textContent || image.alt || '';
    document.body.classList.add('lightbox-open');
    lightbox.classList.add('open');
    closeButton.focus();
  }

  screenshotImages.forEach((image) => {
    image.tabIndex = 0;
    image.setAttribute('role', 'button');
    image.setAttribute('aria-label', `Открыть скриншот: ${image.alt || 'изображение'}`);
    image.addEventListener('click', () => openLightbox(image));
    image.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openLightbox(image);
      }
    });
  });

  closeButton.addEventListener('click', closeLightbox);
  lightbox.addEventListener('click', (event) => {
    if (event.target === lightbox) closeLightbox();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && lightbox.classList.contains('open')) {
      closeLightbox();
    }
  });
}
