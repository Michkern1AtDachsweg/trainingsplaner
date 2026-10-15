// Globale Variablen
let allPlans = {};
let currentPlan = null;
let activeInterval = null;
const boxingBell = new Audio('boxing-bell.mp3');

// Event-Listener beim Laden der Seite
window.addEventListener('load', () => {
    const saved = localStorage.getItem('trainingsplaene');
    if (saved) {
        allPlans = JSON.parse(saved);
        updatePlanSelector();
    }
    initEventListeners();
});

function initEventListeners() {
    document.getElementById('excel-input').addEventListener('change', handleExcelImport);
    document.getElementById('plan-selector').addEventListener('change', (e) => {
        currentPlan = allPlans[e.target.value] || null;
        renderPlanTable();
    });
    document.getElementById('btn-delete-plan').addEventListener('click', () => {
        if (!currentPlan) return;
        if (confirm(`Soll der Plan "${currentPlan.name}" wirklich gelöscht werden?`)) {
            delete allPlans[currentPlan.name];
            saveToStorage();
            updatePlanSelector();
            currentPlan = null;
            renderPlanTable();
        }
    });
    document.getElementById('btn-start-training').addEventListener('click', startTraining);
}

function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            const newPlan = { name: sheetName, steps: [] };

            // **FINALE KORREKTUR: Robuste Schleife, die leere Zeilen ignoriert**
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];

                // Überspringe, wenn die Zeile leer ist oder in Spalte A nichts steht
                if (!row || row.length < 1 || !row[0]) {
                    continue;
                }

                // Überspringe die Header-Zeile
                if (row[0].toString().toLowerCase().includes('zeit')) {
                    continue;
                }

                const baseName = row[0].toString().trim();
                const workMin = parseFloat(String(row[1]).replace(',', '.')) || 0;
                const repeats = parseInt(row[2]) || 1;
                const pauseMin = parseFloat(String(row[3]).replace(',', '.')) || 0;
                const note = row[5] ? row[5].toString().trim() : "";

                for (let r = 1; r <= repeats; r++) {
                    const stepName = repeats > 1 ? `${baseName} (Runde ${r}/${repeats})` : baseName;
                    newPlan.steps.push({ name: stepName, duration: workMin * 60, isPause: false, note: note });

                    if (pauseMin > 0) {
                        const pauseNote = (r < repeats)
                            ? `Nächste Übung: ${baseName} (Runde ${r + 1}/${repeats})`
                            : "Block beendet. Nächste Übung vorbereiten...";
                        newPlan.steps.push({ name: "Pause", duration: pauseMin * 60, isPause: true, note: pauseNote });
                    }
                }
            }

            if (newPlan.steps.length > 0) {
                allPlans[sheetName] = newPlan;
            }
        });

        saveToStorage();
        updatePlanSelector();
        // Wähle den zuletzt importierten Plan aus
        const lastPlanName = workbook.SheetNames[workbook.SheetNames.length - 1];
        planSelector.value = lastPlanName;
        currentPlan = allPlans[lastPlanName];
        renderPlanTable();
        alert("Excel-Datei erfolgreich importiert!");
    };
    reader.readAsArrayBuffer(file);
}

function updatePlanSelector() {
    planSelector.innerHTML = '<option value="">-- Wähle einen Plan --</option>';
    Object.keys(allPlans).forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        planSelector.appendChild(option);
    });
}

function renderPlanTable() {
    const tableBody = document.getElementById('plan-table').querySelector('tbody');
    const footer = document.querySelector('footer');
    tableBody.innerHTML = "";

    if (!currentPlan || currentPlan.steps.length === 0) {
        lblTotalTime.textContent = "Gesamtdauer: 00:00";
        lblEndTime.textContent = "Endzeit: --:-- Uhr";
        footer.style.display = 'none'; // Footer ausblenden, wenn kein Plan da ist
        return;
    }

    footer.style.display = 'flex'; // Footer einblenden

    let totalSeconds = 0;
    currentPlan.steps.forEach(step => {
        totalSeconds += step.duration;
        const tr = document.createElement('tr');
        if (step.isPause) tr.style.backgroundColor = "rgba(41, 128, 185, 0.15)";

        tr.innerHTML = `
            <td>${step.name}</td>
            <td>${formatTime(step.duration)}</td>
            <td>${step.isPause ? "✔️" : "❌"}</td>
            <td style="font-size: 13px; color: #a4b0be;">${step.note}</td>
        `;
        tableBody.appendChild(tr);
    });

    lblTotalTime.textContent = `Gesamtdauer: ${formatTime(totalSeconds)}`;
    const end = new Date(Date.now() + totalSeconds * 1000);
    lblEndTime.textContent = `Endzeit: ${end.getHours().toString().padStart(2, '0')}:${end.getMinutes().toString().padStart(2, '0')} Uhr`;
}

function formatTime(seconds) {
    if (isNaN(seconds) || seconds < 0) seconds = 0;
    const m = Math.floor(seconds / 60).toString().padStart(2, '0');
    const s = (seconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
}

function saveToStorage() {
    localStorage.setItem('trainingsplaene', JSON.stringify(allPlans));
}

function startTraining() {
    // ... (Der Code für den Trainings-Modus bleibt unverändert)
}
