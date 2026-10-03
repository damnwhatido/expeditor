let map = null;
let markersGroup = null;
let routePolylineGroup = null;
let userMarker = null;

let lastValidLocation = null;
let lastValidTime = 0;

// Точні координати вул. Аеропортівська, 4
const BASE_COORDS = [46.4288770, 30.6526377]; 

function initMap() {
  if (!map) {
    map = L.map('map', { zoomControl: false }).setView(BASE_COORDS, 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap'
    }).addTo(map);

    routePolylineGroup = L.layerGroup().addTo(map);
    markersGroup = L.layerGroup().addTo(map);

    // Синій квадрат з білим будиночком (як у додатку)
    const homeIcon = L.divIcon({
      className: 'home-marker-container',
      html: `
        <div style="
          background-color: #5c55e3; 
          width: 32px; 
          height: 32px; 
          border-radius: 8px; 
          display: flex; 
          align-items: center; 
          justify-content: center; 
          box-shadow: 0 3px 6px rgba(0,0,0,0.4);
          border: 2px solid #ffffff;">
          <span style="color: white; font-size: 18px; font-weight: bold; line-height: 1;">🏠</span>
        </div>`,
      iconSize: [32, 32],
      iconAnchor: [16, 16]
    });

    const baseMarker = L.marker(BASE_COORDS, { icon: homeIcon }).addTo(map);
    baseMarker.bindPopup("<b>🏠 База (Виїзд)</b><br>вул. Аеропортівська, 4");

    startGPSWithAntiSpoofing();
  }
}

// Захист від глушилок GPS
function startGPSWithAntiSpoofing() {
  if ("geolocation" in navigator) {
    navigator.geolocation.watchPosition(
      (pos) => {
        const lat = pos.coords.latitude;
        const lon = pos.coords.longitude;
        const accuracy = pos.coords.accuracy;
        const currentTime = Date.now();

        if (accuracy > 150) return;

        if (lastValidLocation && lastValidTime > 0) {
          const R = 6371000;
          const dLat = (lat - lastValidLocation.lat) * Math.PI / 180;
          const dLon = (lon - lastValidLocation.lon) * Math.PI / 180;
          const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
                    Math.cos(lastValidLocation.lat * Math.PI / 180) * Math.cos(lat * Math.PI / 180) *
                    Math.sin(dLon/2) * Math.sin(dLon/2);
          const distance = R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
          const timeDiffSeconds = (currentTime - lastValidTime) / 1000;

          if (timeDiffSeconds > 0 && ((distance / timeDiffSeconds) * 3.6) > 130) {
            return; // Пропускаємо аномальний стрибок від глушилки
          }
        }

        lastValidLocation = { lat, lon };
        lastValidTime = currentTime;

        const userIcon = L.divIcon({
          className: 'user-location-beacon',
          iconSize: [16, 16],
          iconAnchor: [8, 8]
        });

        if (!userMarker) {
          userMarker = L.marker([lat, lon], { icon: userIcon }).addTo(map);
          userMarker.bindPopup("<b>Ви тут</b>");
        } else {
          userMarker.setLatLng([lat, lon]);
        }
      },
      (err) => console.warn(err),
      { enableHighAccuracy: true, maximumAge: 3000, timeout: 10000 }
    );
  }
}

document.getElementById('gpsBtn')?.addEventListener('click', () => {
  if (userMarker && lastValidLocation) {
    map.setView([lastValidLocation.lat, lastValidLocation.lon], 16);
  }
});

const sheet = document.getElementById('bottomSheet');
const handle = document.getElementById('sheetHandle');
if (handle) {
  handle.addEventListener('click', () => sheet.classList.toggle('expanded'));
}

// Завантаження Excel
document.getElementById('excelFile').addEventListener('change', function(e) {
  const file = e.target.files ? e.target.files[0] : null;
  if (!file) return;

  document.getElementById('fileName').innerText = "Файл: " + file.name;

  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    parseExcelData(jsonData);
  };
  reader.readAsArrayBuffer(file);
});

// Допоміжна функція очищення адреси
function cleanAddress(addr) {
  return addr.replace(/Кофе\s*-\s*/gi, '')
             .replace(/Копійка\s*-\s*/gi, '')
             .replace(/Хлібзавод\s*\d*\s*-\s*/gi, '')
             .replace(/Магазин\s*-\s*/gi, '')
             .replace(/Вівіат\s*-\s*/gi, '')
             .trim();
}

async function parseExcelData(rows) {
  const listContainer = document.getElementById('pointsList');
  if (listContainer) listContainer.innerHTML = '';
  
  markersGroup.clearLayers();
  routePolylineGroup.clearLayers();

  let count = 0;
  const pointsToGeocode = [];

  for (const row of rows) {
    if (!row || !Array.isArray(row) || row.length === 0) continue;

    let address = "";
    let clientName = "";
    let weight = "";
    let invoices = "";
    let isHeaderOrNote = false;

    row.forEach(cell => {
      if (!cell) return;
      const str = String(cell).trim();
      const lower = str.toLowerCase();

      if (lower.includes('примечание') || lower.includes('примітка') || lower.includes('маршрут')) {
        isHeaderOrNote = true;
      }

      if (!address && (
        lower.includes('ул.') || lower.includes('вул.') || lower.includes('просп') || 
        lower.includes('пер.') || lower.includes('пров.') || lower.includes('одесс') || 
        lower.includes('одес') || lower.includes('таирова') || lower.includes('черемушки') || 
        lower.includes('поскот') || lower.includes('ш.') || lower.includes('шоссе')
      )) {
        address = str;
      }
      else if (!clientName && (
        lower.includes('фоп') || lower.includes('тов') || lower.includes('пп') || 
        lower.includes('маг') || lower.includes('копійка') || lower.includes('кофе') || lower.includes('магазин')
      )) {
        clientName = str;
      }
      else if (!weight && (lower.includes('кг') || lower.includes('вес') || lower.includes('вага'))) {
        weight = str;
      }
      else if (!invoices && (lower.includes('накл') || lower.includes('сч') || lower.includes('док'))) {
        invoices = str;
      }
    });

    if (isHeaderOrNote || !address) continue;

    count++;

    let fullAddress = address;
    if (!fullAddress.toLowerCase().includes('одесс') && !fullAddress.toLowerCase().includes('одес')) {
      fullAddress = 'Одеса, ' + fullAddress;
    }

    if (!clientName) {
      const firstText = row.find(c => typeof c === 'string' && String(c).trim() !== address);
      clientName = firstText ? String(firstText).trim() : `Точка №${count}`;
    }

    const card = document.createElement('div');
    card.className = 'point-card';
    card.id = `point-card-${count}`;
    card.innerHTML = `
      <div class="point-title">📍 Точка №${count}: ${clientName}</div>
      <div class="point-address"><b>Адреса:</b> ${fullAddress}</div>
      ${weight ? `<div class="point-address"><b>Вага:</b> ${weight}</div>` : ''}
      ${invoices ? `<div class="point-address"><b>Накладні:</b> ${invoices}</div>` : ''}
    `;
    if (listContainer) listContainer.appendChild(card);

    pointsToGeocode.push({
      number: count,
      address: fullAddress,
      searchQuery: cleanAddress(fullAddress),
      clientName,
      weight,
      invoices,
      cardElement: card
    });
  }

  sheet.classList.remove('expanded');
  processPointsGeocoding(pointsToGeocode);
}

// Пакетне геокодування з інтервалом для запобігання блокувань
async function processPointsGeocoding(points) {
  const routeCoords = [BASE_COORDS];

  for (const p of points) {
    try {
      const res = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(p.searchQuery)}`);
      const data = await res.json();

      if (data && data.length > 0) {
        const lat = parseFloat(data[0].lat);
        const lon = parseFloat(data[0].lon);
        const pointCoords = [lat, lon];

        routeCoords.push(pointCoords);

        // Помаранчева пипка з білим обводком та номером (як у додатку)
        const orangeIcon = L.divIcon({
          className: 'orange-pin-marker',
          html: `
            <div style="position: relative; width: 30px; height: 38px; display: flex; justify-content: center; align-items: center;">
              <svg width="30" height="38" viewBox="0 0 30 38" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M15 0C6.71573 0 0 6.71573 0 15C0 26.25 15 38 15 38C15 38 30 26.25 30 15C30 6.71573 23.2843 0 15 0Z" fill="#F36B21" stroke="#FFFFFF" stroke-width="2"/>
                <circle cx="15" cy="14" r="9" fill="#F36B21"/>
              </svg>
              <span style="position: absolute; top: 5px; color: #ffffff; font-weight: bold; font-size: 11px; text-shadow: 0 1px 2px rgba(0,0,0,0.6);">${p.number}</span>
            </div>`,
          iconSize: [30, 38],
          iconAnchor: [15, 38],
          popupAnchor: [0, -34]
        });

        const popupContent = `
          <div style="font-size: 13px; color: #111; line-height: 1.4;">
            <b style="font-size: 15px; color: #F36B21;">📍 Точка №${p.number}</b><br>
            <b>Клієнт:</b> ${p.clientName}<br>
            <b>Адреса:</b> ${p.address}<br>
            ${p.weight ? `<b>Вага:</b> ${p.weight}<br>` : ''}
            ${p.invoices ? `<b>Кількість накладних:</b> ${p.invoices}<br>` : ''}
          </div>
        `;

        const marker = L.marker(pointCoords, { icon: orangeIcon }).addTo(markersGroup);
        marker.bindPopup(popupContent);

        p.cardElement.addEventListener('click', () => {
          map.setView(pointCoords, 16);
          marker.openPopup();
          sheet.classList.remove('expanded');
        });
      }
    } catch (err) {
      console.warn("Помилка геокодування точки №" + p.number, err);
    }

    // Затримка 400мс, щоб сервер OpenStreetMap не блокував запити
    await new Promise(resolve => setTimeout(resolve, 400));
  }

  // Малюємо помаранчеву лінію маршруту
  if (routeCoords.length > 1) {
    L.polyline(routeCoords, {
      color: '#F36B21',
      weight: 5,
      opacity: 0.8,
      lineJoin: 'round'
    }).addTo(routePolylineGroup);

    map.fitBounds(L.latLngBounds(routeCoords), { padding: [40, 40] });
  }
}

// Старт карти
initMap();
    
