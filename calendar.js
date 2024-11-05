// Función para leer el archivo Excel y generar el calendario
function loadExcel() {
    const calendarContainer = document.getElementById('calendarContainer');
    calendarContainer.innerHTML = '';  // Limpiar contenedor

    // Realizar una solicitud HTTP para obtener el archivo Excel
    fetch('src/calendar_25.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Error al cargar el archivo Excel');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });

            // Procesar la primera hoja
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            generateCalendar(jsonData);
        })
        .catch(error => {
            console.error(error);
            alert("No se pudo cargar el archivo Excel.");
        });
}

// Llama a la función loadExcel cuando se cargue la página
window.onload = loadExcel;

// Función para convertir el valor de fecha de Excel al formato "dd-mm-yy"
function excelDateToJSDate(excelDate) {
    const jsDate = new Date((excelDate - 25569) * 86400 * 1000); // Convertir sin suma de días extra
    const day = String(jsDate.getUTCDate()).padStart(2, '0');
    const month = String(jsDate.getUTCMonth() + 1).padStart(2, '0');
    const year = String(jsDate.getUTCFullYear()).slice(-2); // Solo los últimos dos dígitos del año
    return `${day}-${month}-${year}`;
}

// Función para convertir el valor de hora de Excel en decimal al formato "hh:mm"
function excelTimeToJSClock(excelTime) {
    const totalMinutes = Math.round(excelTime * 24 * 60);
    const hours = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const minutes = String(totalMinutes % 60).padStart(2, '0');
    return `${hours}:${minutes}`;
}

// Función para generar el calendario basado en los datos de Excel
function generateCalendar(data) {
    const calendarContainer = document.getElementById('calendarContainer');
    const events = {};
    const monthSet = new Set();

    // Procesar las filas del archivo Excel
    data.slice(1).forEach(row => {
        const [storeName, dateExcel, timeExcel] = row;

        // Convertir fecha y hora al formato adecuado
        const dateStr = excelDateToJSDate(dateExcel);
        const timeStr = excelTimeToJSClock(timeExcel);

        if (!dateStr || !timeStr) {
            console.warn(`Fila con datos incompletos: ${row}`);
            return;
        }

        // Parsear la fecha en formato "dd-mm-yy"
        const [day, month, year] = dateStr.split('-').map(Number);

        if (!year || !month || !day) {
            console.warn(`Fecha inválida en la fila: ${row}`);
            return;
        }

        const fullYear = 2000 + year; // Ajustar el año al siglo XXI
        monthSet.add(`${fullYear}-${month - 1}`);

        if (!events[fullYear]) events[fullYear] = {};
        if (!events[fullYear][month - 1]) events[fullYear][month - 1] = {};
        if (!events[fullYear][month - 1][day]) {
            events[fullYear][month - 1][day] = [];
        }

        events[fullYear][month - 1][day].push(`${storeName} (${timeStr})`);
    });

    // Ordenar `monthSet` y crear el calendario de cada mes en orden cronológico
    Array.from(monthSet)
        .sort((a, b) => {
            const [yearA, monthA] = a.split('-').map(Number);
            const [yearB, monthB] = b.split('-').map(Number);
            return yearA === yearB ? monthA - monthB : yearA - yearB;
        })
        .forEach(monthKey => {
            const [year, month] = monthKey.split('-').map(Number);
            createMonthCalendar(year, month, events[year][month]);
        });

    document.getElementById('calendarTitle').textContent = "Calendario de Eventos Competencias";
    document.getElementById('downloadPdfBtn').style.display = 'inline-block';
}

// Crear un calendario visual para el mes correspondiente
function createMonthCalendar(year, month, days) {
    const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    const calendarContainer = document.getElementById('calendarContainer');

    const monthDiv = document.createElement('div');
    monthDiv.className = 'calendar-month';

    const title = document.createElement('h2');
    title.textContent = `${monthNames[month]} ${year}`;
    monthDiv.appendChild(title);

    const daysOfWeek = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
    const daysRow = document.createElement('div');
    daysRow.className = 'calendar-container';
    daysOfWeek.forEach(day => {
        const dayDiv = document.createElement('div');
        dayDiv.textContent = day;
        dayDiv.className = 'calendar-day';
        daysRow.appendChild(dayDiv);
    });
    monthDiv.appendChild(daysRow);

    const firstDay = (new Date(year, month, 1).getDay() + 6) % 7;
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    let currentDay = 1;
    for (let i = 0; i < 6; i++) {
        const weekRow = document.createElement('div');
        weekRow.className = 'calendar-container';
        for (let j = 0; j < 7; j++) {
            const dayDiv = document.createElement('div');
            dayDiv.className = 'calendar-day';

            if (i === 0 && j < firstDay) {
                dayDiv.innerHTML = '';
            } else if (currentDay > daysInMonth) {
                dayDiv.innerHTML = '';
            } else {
                const dayNumber = document.createElement('div');
                dayNumber.className = 'day-number';
                dayNumber.textContent = currentDay;

                const eventsDiv = document.createElement('div');
                eventsDiv.className = 'events';

                if (days && days[currentDay]) {
                    days[currentDay].forEach(event => {
                        eventsDiv.innerHTML += `<span class="event-point"></span>${event}<br>`;
                    });
                }

                dayDiv.appendChild(dayNumber);
                dayDiv.appendChild(eventsDiv);
                currentDay++;
            }

            if (j === 5 || j === 6) {
                dayDiv.classList.add('weekend');
            }

            weekRow.appendChild(dayDiv);
        }
        monthDiv.appendChild(weekRow);
    }

    calendarContainer.appendChild(monthDiv);
}

// Función para descargar el calendario como PDF
function downloadPDF() {
    const element = document.getElementById('calendarContainer');
    const title = document.getElementById('calendarTitle').textContent;

    const opt = {
        margin: [0, 0, 0, 0],
        filename: `calendario_eventos.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 1.5 },
        jsPDF: { unit: 'in', format: 'letter', orientation: 'landscape' }
    };

    html2pdf().from(element).set(opt).save();
}
