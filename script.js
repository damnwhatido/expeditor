let map = null;
let markersGroup = null;
let userMarker = null;

// Ініціалізація карти
function initMap() {
  if (!map) {
    map = L.map('map', { zoomControl: false }).setView([46.4825, 30.7233], 12);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap'
    }).addTo(map);

    markersGroup = L.layerGroup().addTo(map);
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
        alert("Будь ласка, увімкніть GPS (геолокацію) на телефоні для визначення вашої позиції!");
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
    alert("Очікуємо сигнал GPS...");
  }
});

// Керування висувною шторкою (BottomSheet)
const sheet = document.getElementById('bottomSheet');
const handle = document.getElementById('sheetHandle');

handle.addEventListener('click', () => {
  sheet.classList.toggle('expanded');
});

// Читання Excel файлу
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
    let isHeaderOrNote = false;

    row.forEach(cell => {
      if (!cell) return;
      const str = String(cell).trim();
      const lower = str.toLowerCase();

      // Відсіюємо службові рядки типу "Примечание: Фонтан+Черемушки"
      if (lower.includes('примечание') || lower.includes('примітка') || lower.includes('маршрут')) {
        isHeaderOrNote = true;
      }

      // Шукаємо дійсну адресу магазину/клієнта
      if (!address && (
        lower.includes('ул.') || lower.includes('вул.') || lower.includes('просп') || 
        lower.includes('пер.') || lower.includes('пров.') || lower.includes('одесс') || 
        lower.includes('одес') || lower.includes('таирова') || lower.includes('черемушки') || 
        lower.includes('поскот') || lower.includes('ш.') || lower.includes('шоссе')
      )) {
        address = str;
      }
      // Шукаємо ФОП / ТОВ / Назву магазину
      else if (!clientName && (
        lower.includes('фоп') || lower.includes('тов') || lower.includes('пп') || 
        lower.includes('маг') || lower.includes('копійка') || lower.includes('кофе') || lower.includes('магазин')
      )) {
        clientName = str;
      }
      // Шукаємо вагу
      else if (!weight && (lower.includes('кг') || lower.includes('вес') || lower.includes('вага'))) {
        weight = str;
      }
    });

    // Пропускаємо, якщо це була примітка або не знайшли адресу
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

    // Картка у висувній шторці
    const card = document.createElement('div');
    card.className = 'point-card';
    card.innerHTML = `
      <div class="point-title">📍 Точка №${count}: ${clientName}</div>
      <div class="point-address"><b>Адреса:</b> ${fullAddress}</div>
      ${weight ? `<div class="point-address"><b>Вага:</b> ${weight}</div>` : ''}
    `;
    if (listContainer) listContainer.appendChild(card);

    // Додаємо мітку на карту
    geocodeAndAddMarker(fullAddress, count, clientName, weight);
  }

  // Автоматично згортаємо шторку після завантаження, щоб бачити карту
  sheet.classList.remove('expanded');
}

function geocodeAndAddMarker(address, number, clientName, weight) {
  fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}`)
    .then(res => res.json())
    .then(data => {
      if (data && data.length > 0) {
        const lat = parseFloat(data[0].lat);
        const lon = parseFloat(data[0].lon);

        const popupContent = `
          <div style="font-size: 14px; color: #000;">
            <b style="font-size: 15px; color: #d32f2f;">Точка №${number}</b><br>
            <b>Клієнт:</b> ${clientName}<br>
            <b>Адреса:</b> ${address}<br>
            ${weight ? `<b>Вага:</b> ${weight}<br>` : ''}
          </div>
        `;

        const marker = L.marker([lat, lon]).addTo(markersGroup);
        marker.bindPopup(popupContent);

        const bounds = markersGroup.getLayers().map(m => m.getLatLng());
        if (bounds.length > 0) {
          map.fitBounds(L.latLngBounds(bounds));
        }
      }
    })
    .catch(err => console.log('Помилка геокодування:', err));
}

// Запуск карти при відкритті сайту
initMap();
                                                 
