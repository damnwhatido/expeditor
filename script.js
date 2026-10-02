let extractedAddresses = [];
let map = null;
let markersGroup = null;

// Ініціалізація карти
function initMap() {
  if (!map) {
    map = L.map('map').setView([46.4825, 30.7233], 11); // Одеса
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '© OpenStreetMap'
    }).addTo(map);
    markersGroup = L.layerGroup().addTo(map);
  }
}

document.getElementById('excelFile').addEventListener('change', function(e) {
  const target = e.target;
  const file = target.files ? target.files[0] : null;
  if (!file) return;

  document.getElementById('fileName').innerText = "Файл: " + file.name;

  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    parseAddresses(jsonData);
  };
  reader.readAsArrayBuffer(file);
});

async function parseAddresses(rows) {
  extractedAddresses = [];
  const listContainer = document.getElementById('pointsList');
  const mapDiv = document.getElementById('map');
  
  if (listContainer) listContainer.innerHTML = '';
  mapDiv.style.display = 'block';
  
  initMap();
  markersGroup.clearLayers();

  let count = 0;

  for (const row of rows) {
    if (!row || !Array.isArray(row) || row.length === 0) continue;

    let addressCell = row.find((cell) => {
      if (typeof cell !== 'string') return false;
      const lower = cell.toLowerCase();
      return (
        lower.includes('ул.') || lower.includes('улица') || 
        lower.includes('вул.') || lower.includes('вулиця') || 
        lower.includes('просп') || lower.includes('пер.') || lower.includes('пров.') || 
        lower.includes('обл.') || lower.includes('г.') || lower.includes('м.') || 
        lower.includes('одесс') || lower.includes('одес') || lower.includes('таирова') || 
        lower.includes('черемушки') || lower.includes('поскот')
      );
    });

    if (addressCell && typeof addressCell === 'string') {
      let fullAddress = addressCell.trim();
      const lowerAddr = fullAddress.toLowerCase();
      if (!lowerAddr.includes('одесс') && !lowerAddr.includes('одес')) {
        fullAddress = 'Одеса, ' + fullAddress;
      }

      count++;
      extractedAddresses.push(fullAddress);

      // Карточка точки
      const card = document.createElement('div');
      card.className = 'point-card';
      const encodedAddr = encodeURIComponent(fullAddress);
      const singleMapUrl = `https://www.google.com/maps/search/?api=1&query=${encodedAddr}`;

      card.innerHTML = `
        <div class="point-title">Точка №${count}</div>
        <div class="point-address">${fullAddress}</div>
        <a href="${singleMapUrl}" target="_blank" class="single-btn">📍 Перевірити в Google Maps</a>
      `;
      if (listContainer) listContainer.appendChild(card);

      // Знаходимо координати та ставимо мітку на карті
      geocodeAndMark(fullAddress, count);
    }
  }

  const btn = document.getElementById('buildRouteBtn');
  if (extractedAddresses.length > 0 && btn) {
    btn.disabled = false;
  } else {
    alert('Не вдалося знайти адреси у файлі.');
  }
}

// Запит координат для точки та встановлення маркера
function geocodeAndMark(address, number) {
  fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}`)
    .then(res => res.json())
    .then(data => {
      if (data && data.length > 0) {
        const lat = parseFloat(data[0].lat);
        const lon = parseFloat(data[0].lon);
        
        const marker = L.marker([lat, lon]).addTo(markersGroup);
        marker.bindPopup(`<b>Точка №${number}</b><br>${address}`);
        
        // Масштабуємо карту по всіх маркерах
        const bounds = markersGroup.getLayers().map(m => m.getLatLng());
        if (bounds.length > 0) {
          map.fitBounds(L.latLngBounds(bounds));
        }
      }
    })
    .catch(err => console.log('Помилка геолокації точки:', err));
}

// Запуск повноцінного маршруту в Google Навігаторі
document.getElementById('buildRouteBtn').addEventListener('click', function() {
  if (extractedAddresses.length === 0) return;

  const origin = "Current+Location";
  const destination = encodeURIComponent(extractedAddresses[extractedAddresses.length - 1]);
  const waypointsArray = extractedAddresses.slice(0, extractedAddresses.length - 1);
  const waypoints = waypointsArray.map(addr => encodeURIComponent(addr)).join('|');

  const routeUrl = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}&waypoints=${waypoints}&travelmode=driving`;
  window.open(routeUrl, '_blank');
});
