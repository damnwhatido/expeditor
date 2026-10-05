// ============================================================
//  Карта експедитора: адмін-панель, нумерація точок,
//  найшвидший маршрут, кілька маршрутів
// ============================================================

// ⚠️ ЗМІНИ PIN. Він лежить у коді сайту, тому це захист від випадкових
// змін, а не від зламу. Справжній захист потребує сервера.
const ADMIN_PIN = '1234';

let map = null;
let markersGroup = null;
let routePolylineGroup = null;
let userMarker = null;

let lastValidLocation = null;
let lastValidTime = 0;

let runId = 0;        // щоб старе завантаження не псувало нове
let routes = [];      // весь поточний план
let uid = 0;
let busy = false;
let placementTarget = null; // точка, яку адмін ставить вручну

let isAdmin = false;
try { isAdmin = sessionStorage.getItem('isAdmin') === '1'; } catch (e) { /* ігноруємо */ }

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

function formatDuration(sec) {
  const m = Math.round(sec / 60);
  if (m < 60) return `${m} хв`;
  return `${Math.floor(m / 60)} год ${String(m % 60).padStart(2, '0')} хв`;
}

let toastTimer = null;
function toast(msg, ms = 3500) {
  let t = document.getElementById('adm-toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'adm-toast';
    t.className = 'adm-toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = '1';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.style.opacity = '0'; }, ms);
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

// ---------- Завантаження Excel ----------
document.getElementById('excelFile').addEventListener('change', function (e) {
  const file = e.target.files ? e.target.files[0] : null;
  if (!file) return;

  document.getElementById('fileName').innerText = 'Файл: ' + file.name;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    loadFromExcel(workbook);
  };
  reader.readAsArrayBuffer(file);
  e.target.value = ''; // щоб той самий файл можна було завантажити ще раз
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

// Відокремлює назву закладу, прибирає "м. Одеса" / "ОДЕССА", нормалізує "ул." → "вул."
function parseAddress(raw) {
  let s = String(raw).replace(/\s+/g, ' ').trim();
  let label = '';

  const m = s.match(/^(.*?)\s+[-–—]\s+(.+)$/);
  if (m && !STREET_WORDS.test(m[1])) {
    label = m[1].trim();
    s = m[2].trim();
  }

  s = s.replace(/(?:(?<![а-яіїєґ])(?:м|г)\.?\s*)?(?<![а-яіїєґ])одес+а(?![а-яіїєґ])/gi, ' ');
  s = s.replace(/(?<![а-яіїєґ])ул\.?\s*/gi, 'вул. ');
  s = s.replace(/\s*,\s*/g, ', ').replace(/^[,\s]+|[,\s]+$/g, '').replace(/\s+/g, ' ');

  return { label, street: s };
}

function parseWorkbook(workbook) {
  const result = [];

  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    let current = { name: sheetName, points: [] };

    for (const row of rows) {
      if (!row || !Array.isArray(row) || row.length === 0) continue;
      const r = extractRow(row);

      if (r.isNote) continue;

      if (r.routeTitle && !r.address) {
        if (current.points.length) result.push(current);
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
        invoices: r.invoices,
        origIndex: current.points.length
      });
    }

    if (current.points.length) result.push(current);
  });

  return result;
}

// ---------- Геокодинг ----------
let geoCache = {};
let manualFixes = {};
try { geoCache = JSON.parse(localStorage.getItem('geoCache_v3') || '{}'); } catch (e) { geoCache = {}; }
try { manualFixes = JSON.parse(localStorage.getItem('manualFixes_v1') || '{}'); } catch (e) { manualFixes = {}; }

function saveCache() {
  try { localStorage.setItem('geoCache_v3', JSON.stringify(geoCache)); } catch (e) { /* ігноруємо */ }
}
function saveManualStore() {
  try { localStorage.setItem('manualFixes_v1', JSON.stringify(manualFixes)); } catch (e) { /* ігноруємо */ }
}
function saveManual(p, coords) {
  manualFixes[p.address.toLowerCase()] = coords;
  saveManualStore();
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
    let fallback = null;
    for (const f of (data.features || [])) {
      const [lon, lat] = f.geometry.coordinates; // Photon: [lon, lat]
      const c = [lat, lon];
      if (!isInsideOdesaArea(c)) continue;
      if (f.properties && f.properties.housenumber) return { coords: c, exact: true };
      if (!fallback) fallback = { coords: c, exact: false };
    }
    return fallback;
  } catch (e) {
    console.warn('Photon error:', e);
  }
  return null;
}

async function tryNominatim(query) {
  const wait = 1100 - (Date.now() - lastNominatimTime); // політика: ≤ 1 запит/сек
  if (wait > 0) await new Promise(r => setTimeout(r, wait));
  lastNominatimTime = Date.now();

  const b = CITY_BBOX;
  const url = `https://nominatim.openstreetmap.org/search?format=json&limit=5&countrycodes=ua` +
              `&viewbox=${b.minLon},${b.maxLat},${b.maxLon},${b.minLat}&bounded=1` +
              `&q=${encodeURIComponent(query)}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    let fallback = null;
    for (const item of (data || [])) {
      const c = [parseFloat(item.lat), parseFloat(item.lon)];
      if (!isInsideOdesaArea(c)) continue;
      const exact = item.type === 'house' || item.class === 'building' || item.addresstype === 'building';
      if (exact) return { coords: c, exact: true };
      if (!fallback) fallback = { coords: c, exact: false };
    }
    return fallback;
  } catch (e) {
    console.warn('Nominatim error:', e);
  }
  return null;
}

// Повертає { coords, exact, manual } або null (ніколи не ставимо точку "на око")
async function geocode(point) {
  const key = point.address.toLowerCase();

  if (manualFixes[key]) return { coords: manualFixes[key], exact: true, manual: true };
  if (geoCache[key]) return geoCache[key];

  const ph = await tryPhoton(point.address);
  let result = ph;
  if (!ph || !ph.exact) {
    const nm = await tryNominatim(point.address + ', Україна');
    if (nm && (nm.exact || !ph)) result = nm;
  }

  if (result) {
    geoCache[key] = result;
    saveCache();
  }
  return result;
}

// ---------- Дороги (OSRM) ----------
async function fetchRoadRoute(coords) {
  try {
    const path = coords.map(c => `${c[1]},${c[0]}`).join(';');
    const res = await fetch(`https://router.project-osrm.org/route/v1/driving/${path}?overview=full&geometries=geojson`);
    const data = await res.json();
    if (data.code === 'Ok' && data.routes && data.routes[0]) {
      const r = data.routes[0];
      return {
        coords: r.geometry.coordinates.map(c => [c[1], c[0]]),
        distance: r.distance,
        duration: r.duration
      };
    }
  } catch (e) {
    console.warn('OSRM route error:', e);
  }
  return null;
}

// Порядок об'їзду від бази за реальними дорогами. Повертає індекси точок або null
async function osrmTripOrder(pts) {
  if (pts.length > 95) return null; // ліміт публічного сервера
  const all = [BASE_COORDS, ...pts.map(p => p.coords)];
  const path = all.map(c => `${c[1]},${c[0]}`).join(';');
  try {
    const res = await fetch(`https://router.project-osrm.org/trip/v1/driving/${path}?source=first&roundtrip=false&overview=false`);
    const data = await res.json();
    if (data.code !== 'Ok' || !data.waypoints || data.waypoints.length !== all.length) return null;
    const order = data.waypoints
      .map((w, i) => ({ i, pos: w.waypoint_index }))
      .filter(x => x.i > 0)
      .sort((a, b) => a.pos - b.pos)
      .map(x => x.i - 1);
    return order.length === pts.length ? order : null;
  } catch (e) {
    console.warn('OSRM trip error:', e);
    return null;
  }
}

// Запасний варіант: найближчий сусід + покращення 2-opt (по прямій)
function localOrder(pts) {
  const nodes = [BASE_COORDS, ...pts.map(p => p.coords)];
  const n = nodes.length;
  const d = (a, b) => distanceKm(nodes[a], nodes[b]);

  const left = new Set();
  for (let i = 1; i < n; i++) left.add(i);
  const path = [0];
  while (left.size) {
    const last = path[path.length - 1];
    let best = -1, bd = Infinity;
    for (const j of left) {
      const x = d(last, j);
      if (x < bd) { bd = x; best = j; }
    }
    path.push(best);
    left.delete(best);
  }

  let improved = true, guard = 0;
  while (improved && guard++ < 200) {
    improved = false;
    for (let i = 1; i < n - 1; i++) {
      for (let j = i + 1; j < n; j++) {
        const a = path[i - 1], b = path[i], c = path[j];
        const e = j + 1 < n ? path[j + 1] : null;
        const before = d(a, b) + (e !== null ? d(c, e) : 0);
        const after = d(a, c) + (e !== null ? d(b, e) : 0);
        if (after + 1e-9 < before) {
          path.splice(i, j - i + 1, ...path.slice(i, j + 1).reverse());
          improved = true;
        }
      }
    }
  }
  return path.slice(1).map(i => i - 1);
}

// ---------- Модель: маршрути і точки ----------
function makePoint(raw) {
  return {
    id: ++uid,
    clientName: raw.clientName,
    street: raw.street,
    address: raw.address,
    weight: raw.weight || '',
    invoices: raw.invoices || '',
    origIndex: raw.origIndex ?? 0,
    coords: raw.coords || null,
    exact: !!raw.exact,
    state: raw.state || 'pending', // pending | ok | approx | manual | missing | placing
    number: 0,
    marker: null,
    card: null,
    statusEl: null
  };
}

function newRoute(name, points, idx) {
  return {
    name,
    points,
    color: ROUTE_COLORS[idx % ROUTE_COLORS.length],
    line: null,
    distance: 0,
    duration: 0,
    optimized: false,
    headerEl: null,
    token: 0
  };
}

function clearMapLayers() {
  markersGroup.clearLayers();
  routePolylineGroup.clearLayers();
  placementTarget = null;
}

// ---------- Мітки на карті ----------
function makePinIcon(number, color, isFirst) {
  const len = String(number).length;
  const fs = len <= 2 ? 12 : (len === 3 ? 10 : 8); // цифри влазять і при 100+, і при 1000+
  const stroke = isFirst ? '#22c55e' : '#FFFFFF';
  const sw = isFirst ? 3 : 2;
  return L.divIcon({
    className: 'orange-pin-marker',
    html: `
      <div style="position:relative;width:30px;height:38px;">
        <svg width="30" height="38" viewBox="0 0 30 38" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M15 1.5C7.3 1.5 1.5 7.3 1.5 15C1.5 25.5 15 36.5 15 36.5C15 36.5 28.5 25.5 28.5 15C28.5 7.3 22.7 1.5 15 1.5Z"
                fill="${color}" stroke="${stroke}" stroke-width="${sw}"/>
        </svg>
        <span style="position:absolute;top:6px;left:0;width:30px;text-align:center;line-height:18px;
                     color:#fff;font-weight:800;font-size:${fs}px;text-shadow:0 1px 2px rgba(0,0,0,0.7);">${number}</span>
        ${isFirst ? '<span style="position:absolute;top:-13px;left:19px;font-size:15px;line-height:1;">🚩</span>' : ''}
      </div>`,
    iconSize: [30, 38],
    iconAnchor: [15, 38],
    popupAnchor: [0, -34]
  });
}

function popupEl(route, p) {
  const d = document.createElement('div');
  d.style.cssText = 'font-size:13px;color:#111;line-height:1.4;';
  d.innerHTML = `
    <b style="font-size:15px;color:${route.color};">📍 ${esc(route.name)} · Точка №${p.number}</b><br>
    <b>Клієнт:</b> ${esc(p.clientName)}<br>
    <b>Адреса:</b> ${esc(p.address)}<br>
    ${p.weight ? `<b>Вага:</b> ${esc(p.weight)}<br>` : ''}
    ${p.invoices ? `<b>Накладні:</b> ${esc(p.invoices)}<br>` : ''}`;

  if (isAdmin) {
    const del = document.createElement('button');
    del.textContent = '🗑 Видалити точку';
    del.style.cssText = 'margin-top:8px;padding:6px 10px;border:none;border-radius:8px;background:#e74c3c;color:#fff;font-weight:600;';
    del.addEventListener('click', () => { map.closePopup(); deletePoint(route, p); });
    d.appendChild(del);
  }
  return d;
}

function refreshMarker(route, p, isFirst) {
  if (!p.coords) {
    if (p.marker) { markersGroup.removeLayer(p.marker); p.marker = null; }
    return;
  }

  const icon = makePinIcon(p.number, route.color, isFirst);

  if (!p.marker) {
    const m = L.marker(p.coords, { icon, draggable: isAdmin }).addTo(markersGroup);
    m.bindPopup(() => popupEl(route, p));
    m.on('dragend', () => {
      const ll = m.getLatLng();
      p.coords = [ll.lat, ll.lng];
      p.state = 'manual';
      saveManual(p, p.coords);
      setStatus(p);
      redrawRoute(route);
      savePlan();
    });
    p.marker = m;
  } else {
    p.marker.setLatLng(p.coords);
    p.marker.setIcon(icon);
  }
}

// Нумерація 1..N за порядком у маршруті + підсвітка першої точки
function refreshRoute(route) {
  route.points.forEach((p, i) => { p.number = i + 1; });
  const first = route.points.find(p => p.coords);
  route.points.forEach(p => refreshMarker(route, p, p === first));
}

// ---------- Список (нижня шторка) ----------
function statusInfo(p) {
  switch (p.state) {
    case 'ok':      return ['✅ На карті', '#27ae60'];
    case 'approx':  return [isAdmin ? '⚠️ Приблизно — перетягни піну на місце' : '⚠️ Приблизно (номер будинку не знайдено)', '#e67e22'];
    case 'manual':  return ['📌 Виправлено вручну', '#2E86DE'];
    case 'placing': return ['👆 Тапни на карті, де ця точка', '#2E86DE'];
    case 'missing': return [isAdmin ? '⚠️ Не знайдено — тапни по картці, потім по карті' : '⚠️ Не знайдено на карті', '#e74c3c'];
    default:        return ['⏳ Шукаю на карті…', '#999'];
  }
}

function setStatus(p) {
  if (!p.statusEl) return;
  const [text, color] = statusInfo(p);
  p.statusEl.textContent = text;
  p.statusEl.style.color = color;
}

function updateHeader(route) {
  if (!route.headerEl) return;
  const parts = [`🚚 ${route.name}`, `${route.points.length} точок`];
  if (route.distance) parts.push(`${(route.distance / 1000).toFixed(1)} км`);
  if (route.duration) parts.push(`~${formatDuration(route.duration)}`);
  if (route.optimized) parts.push('⚡');
  route.headerEl.textContent = parts.join(' · ');
}

function toolBtn(text, onClick
