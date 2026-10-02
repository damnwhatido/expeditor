let map = null;
let markersGroup = null;
let userMarker = null;

// Останній надійний координаційний пункт та час
let lastValidLocation = null;
let lastValidTime = 0;

// Координати бази: вул. Аеропортівська, 4, Одеса
const BASE_COORDS = [46.4395, 30.6690]; 

function initMap() {
  if (!map) {
    map = L.map('map', { zoomControl: false }).setView(BASE_COORDS, 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap'
    }).addTo(map);

    markersGroup = L.layerGroup().addTo(map);

    // Синій будиночок (База / Виїзд)
    const homeIcon = L.divIcon({
      className: 'home-icon-marker',
      html: '<div style="font-size:22px; background:#fff; border:2px solid #007bff; border-radius:50%; width:36px; height:36px; display:flex; align-items:center; justify-content:center; box-shadow:0 2px 6px rgba(0,0,0,0.4);">🏠</div>',
      iconSize: [36, 36],
      iconAnchor: [18, 18]
    });

    const baseMarker = L.marker(BASE_COORDS, { icon: homeIcon }).addTo(map);
    baseMarker.bindPopup("<b>🏠 База (Виїзд)</b><br>вул. Аеропортівська, 4");

    startGPSWithAntiSpoofing();
  }
}

// Обчислення відстані між двома точками в метрах (Формула гаверсинусів)
function getDistanceMeters(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

// Захищена геолокація (Захист від глушилок та спуфінгу)
function startGPSWithAntiSpoofing() {
  if ("geolocation" in navigator) {
    navigator.geolocation.watchPosition(
      (pos) => {
        const lat = pos.coords.latitude;
        const lon = pos.coords.longitude;
        const accuracy = pos.coords.accuracy; // Похибка в метрах
        const currentTime = Date.now();

        // 1. Фільтр за похибкою (якщо глушилка дає похибку > 150 метрів — ігноруємо)
        if (accuracy > 150) {
          console.warn("Слабкий або викривлений GPS-сигнал (похибка:", accuracy, "м)");
          return;
        }

        // 2. Фільтр телепортації (перевірка реальної швидкості)
        if (lastValidLocation && lastValidTime > 0) {
          const distance = getDistanceMeters(lastValidLocation.lat, lastValidLocation.lon, lat, lon);
          const timeDiffSeconds = (currentTime - lastValidTime) / 1000;

          if (timeDiffSeconds > 0) {
            const speedKmh = (distance / timeDiffSeconds) * 3.6;

            // Якщо швидкість > 130 км/год — це 100% стрибок від глушилки/спуфінгу
            if (speedKmh > 130) {
              console.warn(`Заблоковано аномальний стрибок GPS! Відстань: ${distance.toFixed(0)}м, Швидкість: ${speedKmh.toFixed(0)} км/год`);
              return; // Ігноруємо фальшиву точку
            }
          }
        }

        // Оновлюємо надійні координати
        lastValidLocation = { lat, lon };
        lastValidTime = currentTime;

        // Малюємо маяк
        const userIcon = L.divIcon({
          className: 'user-location-beacon',
          iconSize: [16, 16],
          iconAnchor: [8, 8]
        });

        if (!userMarker) {
          userMarker = L.marker([lat, lon], { icon: userIcon }).addTo(map);
          userMarker.bindPopup("<b>Ви тут (Реальна позиція)</b>");
        } else {
          userMarker.setLatLng([lat, lon]);
        }
      },
      (err) => {
        console.warn("Помилка GPS:", err.message);
      },
      { enableHighAccuracy: true, maximumAge: 3000, timeout: 10000 }
    );
  }
}

// Кнопка центрування на GPS
document.getElementById('gpsBtn').addEventListener('click', () => {
  if (userMarker && lastValidLocation) {
    map.setView([lastValidLocation.lat, lastValidLocation.lon], 16);
  } else {
    alert("Очікуємо надійний сигнал GPS...");
  }
});

// Шторка BottomSheet
const sheet = document.getElementById('bottomSheet');
const handle = document.getElementById('sheetHandle');

if (handle) {
  handle.addEventListener('click', () => {
    sheet.classList.toggle('expanded');
  });
}

// Читання Excel
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

async function parseExcelData(rows) {
  const listContainer = document.getElementById('pointsList');
  if (listContainer) listContainer.innerHTML = '';
  
  markersGroup.clearLayers();
  let count = 0;

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

    // Додаємо точку на карту з затримкою
    geocodeAndAddMarker(fullAddress, count, clientName, weight, invoices, card);
  }

  sheet.classList.remove('expanded');
}

// Очищення адреси для надійного геокодування
function cleanAddressForSearch(addr) {
  return addr.replace(/Кофе\s*-\s*/gi, '')
             .replace(/Копійка\s*-\s*/gi, '')
             .replace(/Хлібзавод\s*\d*\s*-\s*/gi, '')
             .replace(/Вівіат\s*-\s*/gi, '')
             .trim();
}

function geocodeAndAddMarker(address, number, clientName, weight, invoices, cardElement) {
  const cleanAddr = cleanAddressForSearch(address);

  setTimeout(() => {
    fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(cleanAddr)}`)
      .then(res => res.json())
      .then(data => {
        if (data && data.length > 0) {
          const lat = parseFloat(data[0].lat);
          const lon = parseFloat(data[0].lon);

          // Створення червоного маркера з номером точки
          const numberIcon = L.divIcon({
            className: 'custom-number-marker',
            html: `<div style="background-color:#d32f2f; color:#fff; font-weight:bold; border-radius:50%; width:28px; height:28px; display:flex; align-items:center; justify-content:center; border:2px solid #fff; box-shadow:0 2px 5px rgba(0,0,0,0.5); font-size:13px;">${number}</div>`,
            iconSize: [28, 28],
            iconAnchor: [14, 14]
          });

          const popupContent = `
            <div style="font-size: 13px; color: #111; line-height: 1.4;">
              <b style="font-size: 15px; color: #d32f2f;">📍 Точка №${number}</b><br>
              <b>Клієнт:</b> ${clientName}<br>
              <b>Адреса:</b> ${address}<br>
              ${weight ? `<b>Вага:</b> ${weight}<br>` : ''}
              ${invoices ? `<b>Кількість накладних:</b> ${invoices}<br>` : ''}
            </div>
          `;

          const marker = L.marker([lat, lon], { icon: numberIcon }).addTo(markersGroup);
          marker.bindPopup(popupContent);

          cardElement.addEventListener('click', () => {
            map.setView([lat, lon], 16);
            marker.openPopup();
            sheet.classList.remove('expanded');
          });

          const allLatLngs = markersGroup.getLayers().map(m => m.getLatLng());
          allLatLngs.push(L.latLng(BASE_COORDS[0], BASE_COORDS[1]));
          map.fitBounds(L.latLngBounds(allLatLngs), { padding: [40, 40] });
        }
      })
      .catch(err => console.log('Помилка геокодування:', err));
  }, number * 400);
}

// Запуск
initMap();
