let extractedAddresses = [];

// Відстежуємо вибір файлу Excel
document.getElementById('excelFile').addEventListener('change', function(e) {
  const target = e.target;
  const file = target.files ? target.files[0] : null;
  if (!file) return;

  document.getElementById('fileName').innerText = "Файл: " + file.name;

  const reader = new FileReader();

  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    
    // Парсимо Excel
    const workbook = XLSX.read(data, { type: 'array' });
    
    // Беремо перший аркуш
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Перетворюємо на масив рядків
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    parseAddresses(jsonData);
  };

  reader.readAsArrayBuffer(file);
});

// Розбір адрес (підтримка укр/рос мов)
function parseAddresses(rows) {
  extractedAddresses = [];
  const listContainer = document.getElementById('pointsList');
  if (listContainer) listContainer.innerHTML = '';

  rows.forEach((row) => {
    if (!row || !Array.isArray(row) || row.length === 0) return;

    // Шукаємо комірку з адресою за ключовими словами
    let addressCell = row.find((cell) => {
      if (typeof cell !== 'string') return false;
      const lower = cell.toLowerCase();
      return (
        lower.includes('ул.') || 
        lower.includes('улица') || 
        lower.includes('вул.') || 
        lower.includes('вулиця') || 
        lower.includes('просп') || 
        lower.includes('пер.') || 
        lower.includes('пров.') || 
        lower.includes('шоссе') || 
        lower.includes('ш.') || 
        lower.includes('обл.') || 
        lower.includes('область') || 
        lower.includes('г.') || 
        lower.includes('город') || 
        lower.includes('м.') || 
        lower.includes('місто') || 
        lower.includes('одесс') || 
        lower.includes('одес') || 
        lower.includes('аэропортовская') || 
        lower.includes('черемушки') || 
        lower.includes('таирова') || 
        lower.includes('поскот') || 
        lower.includes('пгт') ||
        lower.includes('смт')
      );
    });

    if (addressCell && typeof addressCell === 'string') {
      let fullAddress = addressCell.trim();
      
      // Якщо міста немає в рядку, за замовчуванням додаємо Одесу
      const lowerAddr = fullAddress.toLowerCase();
      if (!lowerAddr.includes('одесс') && !lowerAddr.includes('одес')) {
        fullAddress = 'Одеса, ' + fullAddress;
      }

      extractedAddresses.push(fullAddress);

      // Малюємо картку точки
      const card = document.createElement('div');
      card.className = 'point-card';
      
      const encodedAddr = encodeURIComponent(fullAddress);
      const singleMapUrl = `https://www.google.com/maps/search/?api=1&query=${encodedAddr}`;

      card.innerHTML = `
        <div class="point-title">Точка №${extractedAddresses.length}</div>
        <div class="point-address">${fullAddress}</div>
        <a href="${singleMapUrl}" target="_blank" class="single-btn">📍 Перевірити в Google Maps</a>
      `;
      
      if (listContainer) listContainer.appendChild(card);
    }
  });

  const btn = document.getElementById('buildRouteBtn');
  if (extractedAddresses.length > 0 && btn) {
    btn.disabled = false;
  } else {
    alert('Не вдалося знайти адреси. Перевірте, чи є у файлі позначки "вул.", "ул.", "м.", "г." тощо.');
  }
}

// Кнопка побудови маршруту для всіх точок
document.getElementById('buildRouteBtn').addEventListener('click', function() {
  if (extractedAddresses.length === 0) return;

  const origin = "Current+Location"; // Старт від твоєї геопозиції
  const destination = encodeURIComponent(extractedAddresses[extractedAddresses.length - 1]); // Фініш — остання точка

  // Усі точки між першою та останньою
  const waypointsArray = extractedAddresses.slice(0, extractedAddresses.length - 1);
  const waypoints = waypointsArray.map(addr => encodeURIComponent(addr)).join('|');

  // Формуємо посилання для Google Maps
  const routeUrl = `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${destination}&waypoints=${waypoints}&travelmode=driving`;

  window.open(routeUrl, '_blank');
});
