const root = document.documentElement;
const themeToggle = document.getElementById('themeToggle');
const savedTheme = localStorage.getItem('theme');
if (savedTheme) root.dataset.theme = savedTheme;

themeToggle?.addEventListener('click', () => {
  const next = root.dataset.theme === 'dark' ? 'light' : 'dark';
  root.dataset.theme = next;
  localStorage.setItem('theme', next);
});

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
