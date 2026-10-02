let map = null;
let markersGroup = null;
let userMarker = null;

// Координати бази: вул. Аеропортівська, 4, Одеса
const BASE_COORDS = [46.4385, 30.6720]; 

function initMap() {
  if (!map) {
    map = L.map('map', { zoomControl: false }).setView(BASE_COORDS, 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap'
    }).addTo(map);

    markersGroup = L.layerGroup().addTo(map);

    // Додаємо постійну синю іконку будиночка (База / Виїзд)
    const homeIcon = L.divIcon({
      className: 'home-icon-marker',
      html: '<div style="font-size:24px; background:#fff; border:2px solid #007bff; border-radius:50%; width:36px; height:36px; display:flex; align-items:center; justify-content:center; box-shadow:0 2px 6px rgba(0,0,0,0.4);">🏠</div>',
      iconSize: [36, 36],
      iconAnchor: [18, 18]
    });

    const baseMarker = L.marker(BASE_COORDS, { icon: homeIcon }).addTo(map);
    baseMarker.bindPopup("<b>🏠 База (Виїзд)</b><br>вул. Аеропортівська, 4");

    startGPS();
  }
}

// Постійне відстеження геолокації з Синім Маяком
function startGPS() {
  if ("geolocation" in navigator) {
    navigator.geolocation.watchPosition(
      (pos) => {
        const lat = pos.coords.latitude;
        const lon = pos.coords.longitude;

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
      (err) => {
        console.log("GPS не активовано");
      },
      { enableHighAccuracy: true, maximumAge: 10000, timeout: 5000 }
    );
  }
}

// Кнопка центрування на GPS
document.getElementById('gpsBtn').addEventListener('click', () => {
  if (userMarker) {
    map.setView(userMarker.getLatLng(), 15);
  } else {
    alert("Очікуємо сигнал GPS... Перевірте, чи увімкнено геолокацію.");
  }
});

// Керування висувною шторкою
const sheet = document.getElementById('bottomSheet');
const handle = document.getElementById('sheetHandle');

handle.addEventListener('click', () => {
  sheet.classList.toggle('expanded');
});

// Парсинг Excel
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

      // Адреса
      if (!address && (
        lower.includes('ул.') || lower.includes('вул.') || lower.includes('просп') || 
        lower.includes('пер.') || lower.includes('пров.') || lower.includes('одесс') || 
        lower.includes('одес') || lower.includes('таирова') || lower.includes('черемушки') || 
        lower.includes('поскот') || lower.includes('ш.') || lower.includes('шоссе')
      )) {
        address = str;
      }
      // Клієнт / ФОП / ТОВ
      else if (!clientName && (
        lower.includes('фоп') || lower.includes('тов') || lower.includes('пп') || 
        lower.includes('маг') || lower.includes('копійка') || lower.includes('кофе') || lower.includes('магазин')
      )) {
        clientName = str;
      }
      // Вага
      else if (!weight && (lower.includes('кг') || lower.includes('вес') || lower.includes('вага'))) {
        weight = str;
      }
      // Кількість накладних
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

    // Картка в списку
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

    geocodeAndAddMarker(fullAddress, count, clientName, weight, invoices, card);
  }

  sheet.classList.remove('expanded');
}

function geocodeAndAddMarker(address, number, clientName, weight, invoices, cardElement) {
  // Використовуємо запит з затримкою, щоб не блокувався сервіс геолокації
  setTimeout(() => {
    fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}`)
      .then(res => res.json())
      .then(data => {
        if (data && data.length > 0) {
          const lat = parseFloat(data[0].lat);
          const lon = parseFloat(data[0].lon);

          // Вміст віконця при натисканні на точку
          const popupContent = `
            <div style="font-size: 13px; color: #111; line-height: 1.4;">
              <b style="font-size: 15px; color: #d32f2f;">📍 Точка №${number}</b><br>
              <b>Клієнт:</b> ${clientName}<br>
              <b>Адреса:</b> ${address}<br>
              ${weight ? `<b>Вага:</b> ${weight}<br>` : ''}
              ${invoices ? `<b>Кількість накладних:</b> ${invoices}<br>` : ''}
            </div>
          `;

          const marker = L.marker([lat, lon]).addTo(markersGroup);
          marker.bindPopup(popupContent);

          // Клік по картці внизу фокусує карту на точці
          cardElement.addEventListener('click', () => {
            map.setView([lat, lon], 16);
            marker.openPopup();
            sheet.classList.remove('expanded');
          });

          // Масштабуємо карту під усі точки разом із базою
          const allLatLngs = markersGroup.getLayers().map(m => m.getLatLng());
          allLatLngs.push(L.latLng(BASE_COORDS[0], BASE_COORDS[1]));
          map.fitBounds(L.latLngBounds(allLatLngs), { padding: [30, 30] });
        }
      })
      .catch(err => console.log('Помилка геолокації:', err));
  }, number * 300); // затримка для стабільного завантаження
}

// Старт карти
initMap();
