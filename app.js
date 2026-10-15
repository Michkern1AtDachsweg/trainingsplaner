let allPlans = {};
let currentPlan = null;
let activeInterval = null;
let boxingBell = null;

let timerState = {
    stepIndex: 0, timeLeft: 0, totalRemaining: 0, isPaused: false
};

window.addEventListener('load', () => {
    try {
        const saved = localStorage.getItem('trainingsplaene');
        if (saved) {
            allPlans = JSON.parse(saved);
            updatePlanSelector();

            // Lade den ersten Plan automatisch, falls vorhanden
            const firstPlanName = Object.keys(allPlans)[0];
            if (firstPlanName) {
                document.getElementById('plan-selector').value = firstPlanName;
                currentPlan = allPlans[firstPlanName];
            }
        }
    } catch (e) { console.error("Fehler beim Laden:", e); }
    initEventListeners();
    renderPlanTable(); // Wichtig: Einmal am Ende aufrufen
});

function initEventListeners() {
    document.getElementById('excel-input').addEventListener('change', handleExcelImport);
    document.getElementById('plan-selector').addEventListener('change', (e) => {
        currentPlan = allPlans[e.target.value] || null;
        renderPlanTable();
    });
    document.getElementById('btn-delete-plan').addEventListener('click', deleteCurrentPlan);
    document.getElementById('btn-start-training').addEventListener('click', startTraining);
    //... (Rest der Listener bleibt gleich)
}

function renderPlanTable() {
    const tableBody = document.getElementById('plan-table').querySelector('tbody');
    const startButton = document.getElementById('btn-start-training');
    tableBody.innerHTML = "";

    if (!currentPlan || currentPlan.steps.length === 0) {
        document.getElementById('lbl-total-time').textContent = "Gesamtdauer: 00:00";
        document.getElementById('lbl-end-time').textContent = "Endzeit: --:-- Uhr";
        startButton.disabled = true; // **FINALE LÖSUNG: Button deaktivieren statt ausblenden**
        return;
    }

    startButton.disabled = false; // **Button aktivieren**

    let totalSeconds = currentPlan.steps.reduce((acc, step) => acc + step.duration, 0);
    currentPlan.steps.forEach(step => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${step.name}</td><td>${formatTime(step.duration)}</td><td>${step.isPause ? "✔️" : "❌"}</td><td>${step.note}</td>`;
        tableBody.appendChild(tr);
    });

    document.getElementById('lbl-total-time').textContent = `Gesamtdauer: ${formatTime(totalSeconds)}`;
    const end = new Date(Date.now() + totalSeconds * 1000);
    document.getElementById('lbl-end-time').textContent = `Endzeit: ${end.getHours().toString().padStart(2, '0')}:${end.getMinutes().toString().padStart(2, '0')} Uhr`;
}

// ... (Rest der app.js bleibt unverändert)
