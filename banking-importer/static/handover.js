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
function showPage(number, moveFocus = true) {
  document.querySelectorAll('[data-page]').forEach(page => {
    page.hidden = page.dataset.page !== String(number);
  });
  document.querySelectorAll('[data-step-label]').forEach(label => {
    label.classList.toggle('completed', Number(label.dataset.stepLabel) < Number(number));
    if (label.dataset.stepLabel === String(number)) label.setAttribute('aria-current', 'step');
    else label.removeAttribute('aria-current');
  });
  if (moveFocus) document.querySelector(`[data-page="${number}"] h1`)?.focus();
  window.scrollTo(0, 0);
}
document.addEventListener('click', event => {
  const navigation = event.target.closest('[data-go]');
  if (navigation && !navigation.disabled) showPage(navigation.dataset.go);
});
document.getElementById('to-report')?.addEventListener('click', () => {
  const field = document.getElementById('workbook-file');
  if (!field.reportValidity()) return;
  const error = document.getElementById('file-error');
  if (!field.files[0].name.toLowerCase().endsWith('.xlsx')) {
    error.textContent = 'Choose an Excel .xlsx workbook.';
    error.hidden = false;
    return;
  }
  error.hidden = true;
  document.getElementById('selected-workbook').textContent = 'Banking file: ' + field.files[0].name;
  showPage(2);
});
document.getElementById('import-form')?.addEventListener('submit', () => {
  const button = document.getElementById('process-button');
  button.disabled = true;
  document.getElementById('processing-message').hidden = false;
});
const firstVisible = document.querySelector('[data-page]:not([hidden])');
if (firstVisible) showPage(firstVisible.dataset.page, false);

for (const [inputId, confirmationId, extension] of [
  ['workbook-file', 'workbook-selected', '.xlsx'],
  ['report-file', 'report-selected', '.pdf']
]) {
  document.getElementById(inputId)?.addEventListener('change', event => {
    const file = event.target.files[0];
    const confirmation = document.getElementById(confirmationId);
    const valid = Boolean(file && file.name.toLowerCase().endsWith(extension));
    event.target.closest('label').classList.toggle('file-present', valid);
    confirmation.hidden = !valid;
    confirmation.textContent = valid ? '✓ Selected: ' + file.name : '';
    event.target.setCustomValidity(file && !valid ? 'Choose a ' + extension + ' file.' : '');
  });
}
document.getElementById('workbook-download')?.addEventListener('click', () => {
  document.getElementById('to-handover').disabled = false;
  document.getElementById('download-status').textContent = 'Download requested. Save it in the banking folder and replace the original before finishing.';
});
