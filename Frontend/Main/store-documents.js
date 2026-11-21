(() => {
  const form = document.getElementById('store-form');
  const input = document.getElementById('purchase-files');
  const summary = document.getElementById('file-summary');
  const status = document.getElementById('store-status');
  const button = form.querySelector('button[type="submit"]');
  const dropZone = document.querySelector('.drop-zone');
  const allowed = /\.(pdf|doc|docx|xls|xlsx|xlsm|zip)$/i;

  const updateSummary = () => {
    const files = [...input.files];
    summary.textContent = files.length ? `${files.length} document(s) selected: ${files.map(file => file.name).join(', ')}` : 'No documents selected.';
  };

  input.addEventListener('change', updateSummary);

  ['dragenter', 'dragover'].forEach(eventName => dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    event.stopPropagation();
    dropZone.classList.add('dragging');
  }));
  ['dragleave', 'drop'].forEach(eventName => dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    event.stopPropagation();
    dropZone.classList.remove('dragging');
  }));
  dropZone.addEventListener('drop', event => {
    const files = [...(event.dataTransfer?.files || [])];
    if (!files.length) return;
    const transfer = new DataTransfer();
    files.forEach(file => transfer.items.add(file));
    input.files = transfer.files;
    updateSummary();
  });

  form.addEventListener('submit', async event => {
    event.preventDefault();
    const files = [...input.files];
    status.className = 'store-status';
    if (!files.length) {
      status.textContent = 'Please select at least one document.';
      status.classList.add('error');
      return;
    }
    const invalid = files.filter(file => !allowed.test(file.name));
    if (invalid.length) {
      status.textContent = `Unsupported file type: ${invalid.map(file => file.name).join(', ')}`;
      status.classList.add('error');
      return;
    }
    const data = new FormData();
    files.forEach(file => data.append('files', file));
    button.disabled = true;
    status.textContent = 'Storing documents…';
    try {
      const response = await fetch('/api/purchase-orders', { method: 'POST', body: data });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'The documents could not be stored.');
      status.textContent = `${result.saved.length} document(s) analyzed and extracted details stored.`;
      status.classList.add('ok');
      if (result.documents?.length) {
        const details = result.documents.map(document => {
          const fields = [`Tender No: ${document.tender_no || 'not found'}`,
            `Supplier: ${document.supplier_name || 'not found'}`,
            `Supplies: ${document.supply_description || 'not found'}`,
            `PO Date: ${document.po_date || 'not found'}`];
          return fields.length ? fields.join(' | ') : 'Metadata not found';
        });
        status.textContent += ` ${details.join('; ')}`;
      }
      if (result.rejected?.length) status.textContent += ` Rejected: ${result.rejected.join(', ')}.`;
      input.value = '';
      summary.textContent = 'No documents selected.';
    } catch (error) {
      status.textContent = error.message;
      status.classList.add('error');
    } finally {
      button.disabled = false;
    }
  });
})();
