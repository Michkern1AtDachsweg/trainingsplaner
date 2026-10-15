// Globale Datenstruktur
let allPlans = {};
let currentPlan = null;
let activeInterval = null;

// Audio-Initialisierung
const boxingBell = new Audio('boxing-bell.mp3');

// Beim Laden der Seite gespeicherte Pläne aus dem Handyspeicher laden
window.addEventListener('load', () => {
    const saved = localStorage.getItem('trainingsplaene');
    if (saved) {
        allPlans = JSON.parse(saved);
        updatePlanSelector();
    }
    initEvents();
});

// UI Elemente
const planSelector = document.getElementById('plan-selector');
const planTable = document.getElementById('plan-table').querySelector('tbody');
const lblTotalTime = document.getElementById('lbl-total-time');
const lblEndTime = document.getElementById('lbl-end-time');

function initEvents() {
    // Excel-Datei hochladen & einlesen
    document.getElementById('excel-input').addEventListener('change', handleExcelImport);

    // Plan wechseln
    planSelector.addEventListener('change', (e) => {
        currentPlan = allPlans[e.target.value] || null;
        renderPlanTable();
    });

    // Plan löschen
    document.getElementById('btn-delete-plan').addEventListener('click', () => {
        if (!currentPlan) return;
        if (confirm(`Soll der Plan "${currentPlan.name}" wirklich gelöscht werden?`)) {
            delete allPlans[currentPlan.name];
            saveToStorage();
            updatePlanSelector();
        }
    });

    // Training starten
    document.getElementById('btn-start-training').addEventListener('click', startTraining);
}

// Speichert die Pläne dauerhaft auf dem Smartphone
function saveToStorage() {
    localStorage.setItem('trainingsplaene', JSON.stringify(allPlans));
}

// Updatet das Dropdown-Menü
function updatePlanSelector() {
    planSelector.innerHTML = '<option value="">-- Wähle einen Plan --</option>';
    Object.keys(allPlans).forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        planSelector.appendChild(option);
    });
    if (currentPlan) {
        planSelector.value = currentPlan.name;
    }
}

// Spaltet Excel-Blätter auf und importiert sie als getrennte Pläne
function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            // Konvertiert das Tabellenblatt in JSON-Format
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonData.length < 3) return; // Überspringe leere Blätter

            const newPlan = { name: sheetName, steps: [] };

            // Datenzeilen verarbeiten (Überspringt die ersten 2 Zeilen für Titel und Header)
            for (let i = 2; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row[0]) continue;

                const baseName = row[0].toString().trim();
                const workMin = parseFloat(row[1]) || 1.0;
                const repeats = parseInt(row[2]) || 1;
                const pauseMin = parseFloat(row[3]) || 0.0;
                const note = row[5] ? row[5].toString().trim() : "";

                // Intervall-Kette generieren
                for (let r = 1; r <= repeats; r++) {
                    const stepName = repeats > 1 ? `${baseName} (Runde ${r}/${repeats})` : baseName;
                    newPlan.steps.push({ name: stepName, duration: workMin * 60, isPause: false, note: note });

                    if (pauseMin > 0) {
                        const pauseNote = r < repeats ? `Nächste Übung: ${baseName} (Runde ${r + 1}/${repeats})` : "Block beendet. Nächste Übung vorbereiten...";
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
        planSelector.value = workbook.SheetNames[workbook.SheetNames.length - 1];
        planSelector.dispatchEvent(new Event('change'));
        alert("Excel-Datei erfolgreich importiert!");
    };
    reader.readAsArrayBuffer(file);
}

// Zeichnet die Tabelle im Hauptbildschirm
function renderPlanTable() {
    planTable.innerHTML = "";
    if (!currentPlan) {
        lblTotalTime.textContent = "Gesamtdauer: 00:00";
        lblEndTime.textContent = "Endzeit: --:-- Uhr";
        return;
    }

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
        planTable.appendChild(tr);
    });

    lblTotalTime.textContent = `Gesamtdauer: ${formatTime(totalSeconds)} Min.`;
    const end = new Date(Date.now() + totalSeconds * 1000);
    lblEndTime.textContent = `Endzeit: ${end.getHours().toString().padStart(2, '0')}:${end.getMinutes().toString().padStart(2, '0')} Uhr`;
}

function formatTime(seconds) {
    const m = Math.floor(seconds / 60).toString().padStart(2, '0');
    const s = (seconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
}

// --- INTERVALLTIMER ENGINE ---
let timerState = {
    stepIndex: 0,
    timeLeft: 0,
    totalRemaining: 0,
    isPaused: false
};

function startTraining() {
    if (!currentPlan || currentPlan.steps.length === 0) return;

    // Wechsel in Vollbild-Ansicht
    document.getElementById('main-view').classList.add('hidden');
    document.getElementById('training-view').classList.remove('hidden');

    timerState.stepIndex = 0;
    timerState.isPaused = false;
    timerState.totalRemaining = currentPlan.steps.reduce((acc, s) => acc + s.duration, 0);

    // Initialen 3-fach Gong abspielen
    playStartBellSequence();

    setTimeout(() => {
        runStep();
    }, 1500);
}

function playStartBellSequence() {
    boxingBell.play();
    setTimeout(() => { boxingBell.play(); }, 400);
    setTimeout(() => { boxingBell.play(); }, 800);
}

function runStep() {
    if (timerState.stepIndex >= currentPlan.steps.length) {
        clearInterval(activeInterval);
        boxingBell.play();
        alert("Training beendet! Starke Leistung!");
        quitTraining();
        return;
    }

    const currentStep = currentPlan.steps[timerState.stepIndex];
    timerState.timeLeft = currentStep.duration;

    // UI aktualisieren
    document.getElementById('timer-step-name').textContent = currentStep.name.toUpperCase();
    document.getElementById('timer-note').textContent = currentStep.note || "";
    document.getElementById('training-view').style.backgroundColor = currentStep.isPause ? "#2980b9" : "#c0392b";

    updateNextExercise();
    boxingBell.play(); // Echte Boxglocke bei jedem Rundenwechsel abspielen!

    clearInterval(activeInterval);
    activeInterval = setInterval(() => {
        if (timerState.isPaused) return;

        // Akustischer 3-Sekunden-Countdown
        if (timerState.timeLeft <= 4 && timerState.timeLeft > 1) {
            playSystemBeep(1800, 80);
        }

        if (timerState.timeLeft <= 0) {
            timerState.stepIndex++;
            runStep();
        } else {
            timerState.timeLeft--;
            timerState.totalRemaining--;
            updateTimerDisplay();
        }
    }, 1000);
}

// Web-Audio-API Erzeugung für flüssigen und unblockierten Countdown-Beep
function playSystemBeep(freq, duration) {
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = "sine";
    osc.frequency.value = freq;
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.start();
    setTimeout(() => { osc.stop(); ctx.close(); }, duration);
}

function updateTimerDisplay() {
    document.getElementById('timer-countdown').textContent = formatTime(timerState.timeLeft);
    document.getElementById('timer-total-remaining').textContent = formatTime(timerState.totalRemaining);
    document.getElementById('timer-rounds').textContent = `${timerState.stepIndex + 1} / ${currentPlan.steps.length}`;

    const end = new Date(Date.now() + timerState.totalRemaining * 1000);
    document.getElementById('timer-end-time').textContent = `${end.getHours().toString().padStart(2, '0')}:${end.getMinutes().toString().padStart(2, '0')}`;
}

function updateNextExercise() {
    const next = currentPlan.steps.slice(timerState.stepIndex + 1).find(s => !s.isPause);
    document.getElementById('timer-next-exercise').textContent = next ? `NÄCHSTE ÜBUNG: ${next.name}` : "LETZTE RUNDE! GIB ALLES!";
}

// Steuerungs-Buttons im Timer
document.getElementById('btn-pause').addEventListener('click', (e) => {
    timerState.isPaused = !timerState.isPaused;
    e.target.textContent = timerState.isPaused ? "▶️ FORTSETZEN" : "⏸️ PAUSE";
    document.getElementById('training-view').style.backgroundColor = timerState.isPaused ? "#7f8c8d" : (currentPlan.steps[timerState.stepIndex].isPause ? "#2980b9" : "#c0392b");
});

document.getElementById('btn-skip').addEventListener('click', () => {
    timerState.totalRemaining -= timerState.timeLeft;
    timerState.stepIndex++;
    runStep();
});

document.getElementById('btn-stop').addEventListener('click', () => {
    if (confirm("Möchtest du das Training wirklich abbrechen?")) {
        quitTraining();
    }
});

function quitTraining() {
    clearInterval(activeInterval);
    document.getElementById('training-view').classList.add('hidden');
    document.getElementById('main-view').classList.remove('hidden');
    renderPlanTable();
}
