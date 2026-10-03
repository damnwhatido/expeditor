// ============================================================
//  Карта експедитора: кілька маршрутів, надійний геокодинг
// ============================================================

let map = null;
let markersGroup = null;
let routePolylineGroup = null;
let userMarker = null;

let lastValidLocation = null;
let lastValidTime = 0;
let runId = 0; // щоб старе завантаження не псувало нове

// База: вул. Аеропортівська, 4
const BASE_COORDS = [46.4288770, 30.6526377];

// Межі, в яких шукаємо точки (Одеса + околиці)
const CITY_BBOX = { minLon: 30.30, minLat: 46.20, maxLon: 31.20, maxLat: 46.75 };
const MAX_DIST_FROM_BASE_KM = 60;

const ROUTE_COLORS = ['#F36B21', '#2E86DE', '#27AE60', '#8E44AD', '#E74C3C', '#16A085', '#D4A017', '#C2185B'];

// ---------- Утиліти ----------
function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

function distanceKm(a, b) {
  const R = 6371;
  const dLat = (b[0] - a[0]) * Math.PI / 180;
  const dLon = (b[1] - a[1]) * Math.PI / 180;
  const x = Math.sin(dLat / 2) ** 2 +
            Math.cos(a[0] * Math.PI / 180) * Math.cos(b[0] * Math.PI / 180) * Math.sin(dLon / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}

function isInsideOdesaArea(c) {
  const [lat, lon] = c;
  if (!isFinite(lat) || !isFinite(lon)) return false;
  if (lon < CITY_BBOX.minLon || lon > CITY_BBOX.maxLon) return false;
  if (lat < CITY_BBOX.minLat || lat > CITY_BBOX.maxLat) return false;
  return distanceKm(BASE_COORDS, c) <= MAX_DIST_FROM_BASE_KM;
}

// ---------- Карта ----------
function initMap() {
  if (map) return;

  map = L.map('map', { zoomControl: false }).setView(BASE_COORDS, 12);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap',
    maxZoom: 19
  }).addTo(map);

  routePolylineGroup = L.layerGroup().addTo(map);
  markersGroup = L.layerGroup().addTo(map);

  const homeIcon = L.divIcon({
    className: 'home-marker-container',
    html: `
      <div style="background-color:#5c55e3;width:32px;height:32px;border-radius:8px;display:flex;
                  align-items:center;justify-content:center;box-shadow:0 3px 6px rgba(0,0,0,0.4);
                  border:2px solid #fff;">
        <span style="color:#fff;font-size:18px;font-weight:bold;line-height:1;">🏠</span>
      </div>`,
    iconSize: [32, 32],
    iconAnchor: [16, 16]
  });

  L.marker(BASE_COORDS, { icon: homeIcon }).addTo(map)
    .bindPopup('<b>🏠 База (Виїзд)</b><br>вул. Аеропортівська, 4');

  startGPSWithAntiSpoofing();
}

// ---------- GPS із захистом від глушилок ----------
function startGPSWithAntiSpoofing() {
  if (!('geolocation' in navigator)) return;

  navigator.geolocation.watchPosition(
    (pos) => {
      const lat = pos.coords.latitude;
      const lon = pos.coords.longitude;
      const now = Date.now();

      if (pos.coords.accuracy > 150) return;

      if (lastValidLocation && lastValidTime > 0) {
        const km = distanceKm([lastValidLocation.lat, lastValidLocation.lon], [lat, lon]);
        const sec = (now - lastValidTime) / 1000;
        if (sec > 0 && (km * 1000 / sec) * 3.6 > 130) return; // аномальний стрибок
      }

      lastValidLocation = { lat, lon };
      lastValidTime = now;

      if (!userMarker) {
        const userIcon = L.divIcon({ className: 'user-location-beacon', iconSize: [16, 16], iconAnchor: [8, 8] });
        userMarker = L.marker([lat, lon], { icon: userIcon }).addTo(map).bindPopup('<b>Ви тут</b>');
      } else {
        userMarker.setLatLng([lat, lon]);
      }
    },
    (err) => console.warn(err),
    { enableHighAccuracy: true, maximumAge: 3000, timeout: 10000 }
  );
}

document.getElementById('gpsBtn')?.addEventListener('click', () => {
  if (userMarker && lastValidLocation) map.setView([lastValidLocation.lat, lastValidLocation.lon], 16);
});

const sheet = document.getElementById('bottomSheet');
document.getElementById('sheetHandle')?.addEventListener('click', () => sheet?.classList.toggle('expanded'));

// ---------- Завантаження Excel (усі листи = маршрути) ----------
document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files ? e.target.files[0] : null;
  if (!file) return;

  document.getElementById('fileName').innerText = 'Файл: ' + file.name;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const routes = parseWorkbook(workbook);
    renderRoutes(routes);
  };
  reader.readAsArrayBuffer(file);
});

// ---------- Розбір Excel ----------
const STREET_WORDS = /(вул|ул|просп|пр-т|пров|пер|шосе|ш\.|бульв|узвіз|дорога)/i;

function extractRow(row) {
  const r = { address: '', clientName: '', weight: '', invoices: '', isNote: false, routeTitle: '' };

  row.forEach(cell => {
    if (cell === null || cell === undefined || cell === '') return;
    const str = String(cell).trim();
    const lower = str.toLowerCase();

    if (lower.includes('примечание') || lower.includes('примітка')) r.isNote = true;
    if (lower.includes('маршрут') && !r.routeTitle) r.routeTitle = str;

    if (!r.address && (
      lower.includes('ул.') || lower.includes('вул.') || lower.includes('просп') ||
      lower.includes('пер.') || lower.includes('пров.') || lower.includes('одесс') ||
      lower.includes('одес') || lower.includes('таирова') || lower.includes('черемушки') ||
      lower.includes('поскот') || lower.includes('ш.') || lower.includes('шоссе') ||
      lower.includes('шосе')
    )) {
      r.address = str;
    } else if (!r.clientName && (
      lower.includes('фоп') || lower.includes('тов') || lower.includes('пп') ||
      lower.includes('маг') || lower.includes('копійка') || lower.includes('кофе')
    )) {
      r.clientName = str;
    } else if (!r.weight && (lower.includes('кг') || lower.includes('вес') || lower.includes('вага'))) {
      r.weight = str;
    } else if (!r.invoices && (lower.includes('накл') || lower.includes('сч') || lower.includes('док'))) {
      r.invoices = str;
    }
  });

  return r;
}

// Чистить адресу: відокремлює назву закладу, прибирає "м. Одеса" / "ОДЕССА", нормалізує "ул." → "вул."
function parseAddress(raw) {
  let s = String(raw).replace(/\s+/g, ' ').trim();
  let label = '';

  // "Бон Чикен - ОДЕССА м. Одеса вул. Жуковського 40" → label + адреса
  const m = s.match(/^(.*?)\s+[-–—]\s+(.+)$/);
  if (m && !STREET_WORDS.test(m[1])) {
    label = m[1].trim();
    s = m[2].trim();
  }

  // прибираємо "м. Одеса", "г. Одесса", "ОДЕССА"
  s = s.replace(/(?:(?<![а-яіїєґ])(?:м|г)\.?\s*)?(?<![а-яіїєґ])одес+а(?![а-яіїєґ])/gi, ' ');
  // ул. → вул.
  s = s.replace(/(?<![а-яіїєґ])ул\.?\s*/gi, 'вул. ');

  s = s.replace(/\s*,\s*/g, ', ').replace(/^[,\s]+|[,\s]+$/g, '').replace(/\s+/g, ' ');

  return { label, street: s };
}

function parseWorkbook(workbook) {
  const routes = [];

  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    let current = { name: sheetName, points: [] };

    for (const row of rows) {
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      const r = extractRow(row);

      if (r.isNote) continue;

      // Рядок з написом "Маршрут ..." без адреси = початок нового маршруту
      if (r.routeTitle && !r.address) {
        if (current.points.length) routes.push(current);
        current = { name: r.routeTitle, points: [] };
        continue;
      }

      if (!r.address) continue;

      const parsed = parseAddress(r.address);
      if (!parsed.street) continue;

      let clientName = r.clientName || parsed.label;
      if (!clientName) {
        const firstText = row.find(c => typeof c === 'string' && String(c).trim() !== r.address);
        clientName = firstText ? String(firstText).trim() : '';
      }

      current.points.push({
        clientName: clientName || 'Без назви',
        street: parsed.street,
        address: parsed.street + ', Одеса',
        weight: r.weight,
        invoices: r.invoices
      });
    }

    if (current.points.length) routes.push(current);
  });

  return routes;
}

// ---------- Геокодинг ----------
let geoCache = {};
try { geoCache = JSON.parse(localStorage.getItem('geoCache_v2') || '{}'); } catch (e) { geoCache = {}; }
function saveCache() {
  try { localStorage.setItem('geoCache_v2', JSON.stringify(geoCache)); } catch (e) { /* ігноруємо */ }
}

let lastNominatimTime = 0;

async function tryPhoton(query) {
  const b = CITY_BBOX;
  const url = `https://photon.komoot.io/api/?q=${encodeURIComponent(query)}&limit=5` +
              `&lat=${BASE_COORDS[0]}&lon=${BASE_COORDS[1]}` +
              `&bbox=${b.minLon},${b.minLat},${b.maxLon},${b.maxLat}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    for (const f of (data.features || [])) {
      const [lon, lat] = f.geometry.coordinates; // Photon: [lon, lat]
      const c = [lat, lon];
      if (isInsideOdesaArea(c)) return c;
    }
  } catch (e) {
    console.warn('Photon error:', e);
  }
  return null;
}

async function tryNominatim(query) {
  // політика Nominatim: не частіше 1 запиту на секунду
  const wait = 1100 - (Date.now() - lastNominatimTime);
  if (wait > 0) await new Promise(r => setTimeout(r, wait));
  lastNominatimTime = Date.now();

  const b = CITY_BBOX;
  const url = `https://nominatim.openstreetmap.org/search?format=json&limit=5&countrycodes=ua` +
              `&viewbox=${b.minLon},${b.maxLat},${b.maxLon},${b.minLat}&bounded=1` +
              `&q=${encodeURIComponent(query)}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    for (const item of (data || [])) {
      const c = [parseFloat(item.lat), parseFloat(item.lon)];
      if (isInsideOdesaArea(c)) return c;
    }
  } catch (e) {
    console.warn('Nominatim error:', e);
  }
  return null;
}

// Повертає координати або null (ніколи не ставимо точку "на око")
async function geocode(point) {
  const key = point.address.toLowerCase();
  if (geoCache[key]) return geoCache[key];

  const coords = (await tryPhoton(point.address)) || (await tryNominatim(point.address + ', Україна'));
  if (coords) {
    geoCache[key] = coords;
    saveCache();
  }
  return coords;
}

// ---------- Лінія маршруту по дорогах (OSRM) ----------
async function fetchRoadRoute(coords) {
  try {
    const path = coords.map(c => `${c[1]},${c[0]}`).join(';');
    const res = await fetch(`https://router.project-osrm.org/route/v1/driving/${path}?overview=full&geometries=geojson`);
    const data = await res.json();
    if (data.code === 'Ok' && data.routes && data.routes[0]) {
      return data.routes[0].geometry.coordinates.map(c => [c[1], c[0]]);
    }
  } catch (e) {
    console.warn('OSRM error:', e);
  }
  return null; // тоді малюємо пряму лінію
}

// ---------- Малювання ----------
function makePinIcon(number, color) {
  return L.divIcon({
    className: 'orange-pin-marker',
    html: `
      <div style="position:relative;width:30px;height:38px;display:flex;justify-content:center;align-items:center;">
        <svg width="30" height="38" viewBox="0 0 30 38" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M15 0C6.71573 0 0 6.71573 0 15C0 26.25 15 38 15 38C15 38 30 26.25 30 15C30 6.71573 23.2843 0 15 0Z"
                fill="${color}" stroke="#FFFFFF" stroke-width="2"/>
          <circle cx="15" cy="14" r="9" fill="${color}"/>
        </svg>
        <span style="position:absolute;top:5px;color:#fff;font-weight:bold;font-size:11px;
                     text-shadow:0 1px 2px rgba(0,0,0,0.6);">${number}</span>
      </div>`,
    iconSize: [30, 38],
    iconAnchor: [15, 38],
    popupAnchor: [0, -34]
  });
}

async function renderRoutes(routes) {
  const myRun = ++runId;
  const listContainer = document.getElementById('pointsList');
  if (listContainer) listContainer.innerHTML = '';
  markersGroup.clearLayers();
  routePolylineGroup.clearLayers();

  if (!routes.length) {
    if (listContainer) listContainer.innerHTML = '<div style="padding:12px;">Адрес у файлі не знайдено.</div>';
    return;
  }

  // 1) Спочатку будуємо всі картки
  routes.forEach((route, ri) => {
    route.color = ROUTE_COLORS[ri % ROUTE_COLORS.length];
    route.bounds = [];

    const header = document.createElement('div');
    header.style.cssText = `padding:10px 12px;margin:10px 0 6px;border-left:6px solid ${route.color};
                            font-weight:bold;background:rgba(128,128,128,0.15);border-radius:6px;cursor:pointer;`;
    header.textContent = `🚚 ${route.name} (${route.points.length})`;
    header.addEventListener('click', () => {
      if (route.bounds.length) {
        map.fitBounds(L.latLngBounds(route.bounds), { padding: [40, 40] });
        sheet?.classList.remove('expanded');
      }
    });
    listContainer?.appendChild(header);

    route.points.forEach((p, i) => {
      p.number = i + 1;
      const card = document.createElement('div');
      card.className = 'point-card';
      card.style.borderLeft = `4px solid ${route.color}`;
      card.innerHTML = `
        <div class="point-title">📍 Точка №${p.number}: ${esc(p.clientName)}</div>
        <div class="point-address"><b>Адреса:</b> ${esc(p.address)}</div>
        ${p.weight ? `<div class="point-address"><b>Вага:</b> ${esc(p.weight)}</div>` : ''}
        ${p.invoices ? `<div class="point-address"><b>Накладні:</b> ${esc(p.invoices)}</div>` : ''}
        <div class="geo-status point-address">⏳ Шукаю на карті…</div>`;
      listContainer?.appendChild(card);
      p.card = card;
      p.statusEl = card.querySelector('.geo-status');
    });
  });

  sheet?.classList.remove('expanded');

  // 2) Геокодимо і малюємо по черзі
  const allBounds = [BASE_COORDS];

  for (const route of routes) {
    const routeCoords = [BASE_COORDS];

    for (const p of route.points) {
      if (myRun !== runId) return; // завантажили інший файл — зупиняємось

      const coords = await geocode(p);

      if (!coords) {
        p.statusEl.textContent = '⚠️ Не знайдено на карті — перевір адресу';
        p.statusEl.style.color = '#e74c3c';
        continue;
      }

      p.statusEl.textContent = '✅ На карті';
      p.statusEl.style.color = '#27ae60';

      routeCoords.push(coords);
      route.bounds.push(coords);
      allBounds.push(coords);

      const popup = `
        <div style="font-size:13px;color:#111;line-height:1.4;">
          <b style="font-size:15px;color:${route.color};">📍 ${esc(route.name)} · Точка №${p.number}</b><br>
          <b>Клієнт:</b> ${esc(p.clientName)}<br>
          <b>Адреса:</b> ${esc(p.address)}<br>
          ${p.weight ? `<b>Вага:</b> ${esc(p.weight)}<br>` : ''}
          ${p.invoices ? `<b>Накладні:</b> ${esc(p.invoices)}<br>` : ''}
        </div>`;

      const marker = L.marker(coords, { icon: makePinIcon(p.number, route.color) }).addTo(markersGroup);
      marker.bindPopup(popup);

      p.card.addEventListener('click', () => {
        map.setView(coords, 16);
        marker.openPopup();
        sheet?.classList.remove('expanded');
      });
    }

    // Лінія цього маршруту: по дорогах, а якщо не вийшло — пряма
    if (routeCoords.length > 1 && myRun === runId) {
      const road = await fetchRoadRoute(routeCoords);
      L.polyline(road || routeCoords, {
        color: route.color,
        weight: 5,
        opacity: 0.85,
        lineJoin: 'round',
        dashArray: road ? null : '8 8'
      }).addTo(routePolylineGroup);
    }
  }

  if (myRun === runId && allBounds.length > 1) {
    map.fitBounds(L.latLngBounds(allBounds), { padding: [40, 40] });
  }
}

// Запуск
initMap();
