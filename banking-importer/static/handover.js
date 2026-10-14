document.addEventListener('click', async (event) => {
  const button = event.target.closest('[data-copy]');
  if (!button) return;
  const field = document.getElementById(button.dataset.copy);
  const status = document.getElementById('copy-status');
  try {
    await navigator.clipboard.writeText(field.value);
    status.textContent = 'Copied.';
  } catch {
    field.focus();
    field.select();
    status.textContent = 'Text selected. Press Command+C on Mac or Ctrl+C on Windows to copy.';
  }
});
