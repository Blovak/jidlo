const config = window.APP_CONFIG || {};

const state = {
  items: [],
  selectedIds: new Set(),
  loading: false,
  saving: false
};

const elements = {
  date: document.querySelector('#menu-date'),
  groups: document.querySelector('#menu-groups'),
  loading: document.querySelector('#loading-state'),
  empty: document.querySelector('#empty-state'),
  error: document.querySelector('#error-state'),
  errorMessage: document.querySelector('#error-message'),
  setup: document.querySelector('#setup-notice'),
  refresh: document.querySelector('#refresh-button'),
  retry: document.querySelector('#retry-button'),
  bar: document.querySelector('#selection-bar'),
  count: document.querySelector('#selection-count'),
  total: document.querySelector('#selection-total'),
  save: document.querySelector('#save-button'),
  toast: document.querySelector('#toast')
};

const apiConfigured = Boolean(
  config.API_URL &&
  /^https:\/\/script\.google\.com\//.test(config.API_URL) &&
  /\/exec(?:\?|$)/.test(config.API_URL)
);

function jsonp(parameters, timeout = 30000) {
  return new Promise((resolve, reject) => {
    const callbackName = `cb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    const timer = window.setTimeout(() => cleanup(new Error('Server neodpověděl včas.')), timeout);
    const url = new URL(config.API_URL);

    Object.entries(parameters).forEach(([key, value]) => url.searchParams.set(key, value));
    url.searchParams.set('callback', `window.__jidloCallbacks.${callbackName}`);
    url.searchParams.set('_', Date.now());

    window.__jidloCallbacks ||= {};
    window.__jidloCallbacks[callbackName] = payload => cleanup(null, payload);

    script.src = url.toString();
    script.async = true;
    script.onerror = () => cleanup(new Error('API není dostupné.'));
    document.head.append(script);

    function cleanup(error, value) {
      window.clearTimeout(timer);
      script.remove();
      delete window.__jidloCallbacks[callbackName];
      if (error) reject(error);
      else resolve(value);
    }
  });
}

async function loadMenu() {
  if (!apiConfigured) {
    elements.loading.classList.add('is-hidden');
    elements.setup.classList.remove('is-hidden');
    elements.date.textContent = 'Aplikace čeká na propojení s Google Sheety';
    return;
  }

  setLoading(true);
  state.selectedIds.clear();
  updateSelectionBar();

  try {
    const response = await jsonp({action: 'menu'});
    if (!response.ok) throw new Error(response.error || 'Neznámá chyba API.');

    state.items = Array.isArray(response.items) ? response.items : [];
    elements.date.textContent = formatDate(response.date);
    renderMenu();
  } catch (error) {
    showError(error.message);
  } finally {
    setLoading(false);
  }
}

function renderMenu() {
  elements.groups.replaceChildren();
  elements.error.classList.add('is-hidden');
  elements.empty.classList.toggle('is-hidden', state.items.length > 0);

  if (!state.items.length) return;

  const grouped = Map.groupBy
    ? Map.groupBy(state.items, item => item.section)
    : state.items.reduce((map, item) => {
        const entries = map.get(item.section) || [];
        entries.push(item);
        map.set(item.section, entries);
        return map;
      }, new Map());

  for (const [section, items] of grouped) {
    const group = document.createElement('section');
    group.className = 'menu-group';

    const heading = document.createElement('h3');
    heading.className = 'group-title';
    heading.textContent = section;

    const list = document.createElement('div');
    list.className = 'dish-list';
    items.forEach(item => list.append(createDishCard(item)));

    group.append(heading, list);
    elements.groups.append(group);
  }
}

function createDishCard(item) {
  const label = document.createElement('label');
  label.className = 'dish-card';
  label.dataset.itemId = item.id;

  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.value = item.id;
  checkbox.setAttribute('aria-label', `Vybrat ${item.name}`);
  checkbox.addEventListener('change', () => toggleItem(item.id, checkbox.checked, label));

  const checkmark = document.createElement('span');
  checkmark.className = 'checkmark';
  checkmark.setAttribute('aria-hidden', 'true');
  checkmark.innerHTML = '<svg viewBox="0 0 24 24"><path d="m6 12 4 4 8-9"/></svg>';

  const copy = document.createElement('span');
  copy.className = 'dish-copy';
  const name = document.createElement('strong');
  name.textContent = item.name;
  copy.append(name);

  const estimatedEnergyKcal = Number(item.estimatedEnergyKcal);
  if (Number.isFinite(estimatedEnergyKcal) && estimatedEnergyKcal > 0) {
    const energy = document.createElement('span');
    energy.className = 'energy-estimate';
    energy.textContent = `AI odhad: ≈ ${formatEnergy(estimatedEnergyKcal)}`;
    energy.title = 'Orientační energetická hodnota typické porce odhadnutá pomocí AI.';
    copy.append(energy);
  }

  if (item.allergens) {
    const allergens = document.createElement('span');
    allergens.className = 'allergens';
    allergens.textContent = `Alergeny: ${item.allergens}`;
    copy.append(allergens);
  }

  const price = document.createElement('span');
  price.className = 'dish-price';
  price.textContent = formatPrice(item.price);

  label.append(checkbox, checkmark, copy, price);
  return label;
}

function toggleItem(id, selected, card) {
  if (selected) state.selectedIds.add(id);
  else state.selectedIds.delete(id);
  card.classList.toggle('is-selected', selected);
  updateSelectionBar();
}

function updateSelectionBar() {
  const selected = state.items.filter(item => state.selectedIds.has(item.id));
  const total = selected.reduce((sum, item) => sum + (Number(item.price) || 0), 0);
  const count = selected.length;

  elements.count.textContent = itemCountLabel(count);
  elements.total.textContent = formatPrice(total);
  elements.save.disabled = count === 0 || state.saving;
  elements.bar.classList.toggle('is-visible', count > 0);
}

async function saveSelection() {
  if (!state.selectedIds.size || state.saving) return;

  state.saving = true;
  elements.save.disabled = true;
  elements.save.querySelector('span').textContent = 'Ukládám…';

  const requestId = createRequestId();
  const payload = {
    action: 'save',
    requestId,
    selectedIds: [...state.selectedIds]
  };

  try {
    await fetch(config.API_URL, {
      method: 'POST',
      mode: 'no-cors',
      redirect: 'follow',
      headers: {'Content-Type': 'text/plain;charset=utf-8'},
      body: JSON.stringify(payload)
    });

    const confirmed = await waitForConfirmation(requestId);
    if (!confirmed) throw new Error('Uložení se nepodařilo potvrdit. Zkuste to znovu.');

    state.selectedIds.clear();
    document.querySelectorAll('.dish-card').forEach(card => {
      card.classList.remove('is-selected');
      card.querySelector('input').checked = false;
    });
    showToast('Výběr je uložený v Google Sheetu ✓');
  } catch (error) {
    showToast(error.message || 'Výběr se nepodařilo uložit.', true);
  } finally {
    state.saving = false;
    elements.save.querySelector('span').textContent = 'Uložit výběr';
    updateSelectionBar();
  }
}

async function waitForConfirmation(requestId) {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    await delay(attempt === 0 ? 700 : 1100);
    try {
      const response = await jsonp({action: 'status', requestId}, 10000);
      if (response.ok && response.saved) return true;
    } catch (_) {
      // Krátký výpadek při ověřování není důvod k duplicitnímu POSTu.
    }
  }
  return false;
}

function setLoading(loading) {
  state.loading = loading;
  elements.loading.classList.toggle('is-hidden', !loading);
  elements.refresh.classList.toggle('is-spinning', loading);
  elements.refresh.disabled = loading;
  if (loading) {
    elements.groups.replaceChildren();
    elements.empty.classList.add('is-hidden');
    elements.error.classList.add('is-hidden');
  }
}

function showError(message) {
  elements.errorMessage.textContent = message;
  elements.error.classList.remove('is-hidden');
  elements.empty.classList.add('is-hidden');
}

let toastTimer;
function showToast(message, isError = false) {
  window.clearTimeout(toastTimer);
  elements.toast.textContent = message;
  elements.toast.style.background = isError ? '#8d3828' : '';
  elements.toast.classList.add('is-visible');
  toastTimer = window.setTimeout(() => elements.toast.classList.remove('is-visible'), 4200);
}

function formatDate(dateKey) {
  const date = new Date(`${dateKey}T12:00:00`);
  if (Number.isNaN(date.valueOf())) return dateKey;
  const formatted = new Intl.DateTimeFormat('cs-CZ', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  }).format(date);
  return formatted.charAt(0).toUpperCase() + formatted.slice(1);
}

function formatPrice(value) {
  return new Intl.NumberFormat('cs-CZ', {
    style: 'currency',
    currency: 'CZK',
    maximumFractionDigits: 0
  }).format(Number(value) || 0);
}

function formatEnergy(value) {
  return `${new Intl.NumberFormat('cs-CZ', {maximumFractionDigits: 0}).format(value)} kcal`;
}

function itemCountLabel(count) {
  if (count === 1) return '1 položka';
  if (count >= 2 && count <= 4) return `${count} položky`;
  return `${count} položek`;
}

function createRequestId() {
  if (crypto.randomUUID) return `jidlo_${crypto.randomUUID()}`;
  const random = crypto.getRandomValues(new Uint32Array(2)).join('_');
  return `jidlo_${Date.now()}_${random}`;
}

function delay(milliseconds) {
  return new Promise(resolve => window.setTimeout(resolve, milliseconds));
}

elements.refresh.addEventListener('click', loadMenu);
elements.retry.addEventListener('click', loadMenu);
elements.save.addEventListener('click', saveSelection);

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  window.addEventListener('load', () => navigator.serviceWorker.register('service-worker.js'));
}

loadMenu();
