// Globale Variablen für die App
let allPlans = {};
let currentPlan = null;
let activeInterval = null;

// Audio-Datei initialisieren (muss im gleichen Ordner wie index.html liegen)
const boxingBell = new Audio('boxing-bell.mp3');

// Timer-Statusobjekt
let timerState = {
    stepIndex: 0,
    timeLeft: 0,
    totalRemaining: 0,
    isPaused: false
};

// Automatisches Laden beim Starten der Seite
window.addEventListener('load', () => {
    try {
        const saved = localStorage.getItem('trainingsplaene');
        if (saved) {
            allPlans = JSON.parse(saved);
            updatePlanSelector();
        }
    } catch (e) {
        console.error("Fehler beim Laden aus dem Speicher:", e);
    }
    initEventListeners();
});

// Registriert alle Button-Klicks und Auswahlen
function initEventListeners() {
    document.getElementById('excel-input').addEventListener('change', handleExcelImport);

    document.getElementById('plan-selector').addEventListener('change', (e) => {
        currentPlan = allPlans[e.target.value] || null;
        renderPlanTable();
    });

    document.getElementById('btn-delete-plan').addEventListener('click', deleteCurrentPlan);
    document.getElementById('btn-start-training').addEventListener('click', startTraining);
    document.getElementById('btn-pause').addEventListener('click', togglePause);
    document.getElementById('btn-skip').addEventListener('click', skipStep);
    document.getElementById('btn-stop').addEventListener('click', () => quitTraining(false));
}

// Speichert den aktuellen Stand im Smartphone-Speicher
function saveToStorage() {
    try {
        localStorage.setItem('trainingsplaene', JSON.stringify(allPlans));
    } catch (e) {
        console.error("Fehler beim Speichern:", e);
    }
}

// Löscht den aktuell ausgewählten Trainingsplan
function deleteCurrentPlan() {
    if (!currentPlan) return;
    if (confirm(`Soll der Plan "${currentPlan.name}" wirklich gelöscht werden?`)) {
        delete allPlans[currentPlan.name];
        saveToStorage();
        updatePlanSelector();
        currentPlan = null;
        renderPlanTable();
    }
}

// Aktualisiert das Dropdown-Auswahlmenü im Header
function updatePlanSelector() {
    const planSelector = document.getElementById('plan-selector');
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

// Liest deine .xlsx-Excel-Datei direkt im Webbrowser ein
function handleExcelImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length === 0) return;

                const newPlan = { name: sheetName, steps: [] };

                // Wir gehen jede Zeile der Tabelle durch und überspringen Kopfzeilen
                for (let i = 0; i < jsonData.length; i++) {
                    const row = jsonData[i];

                    // Zeile überspringen, wenn leer oder Überschrift
                    if (!row || row.length < 1 || !row[0]) continue;
                    if (row[0].toString().toLowerCase().includes('zeit') || row[0].toString().toLowerCase().includes('plan')) {
                        continue;
                    }

                    const baseName = row[0].toString().trim();
                    const workMin = parseFloat(String(row[1]).replace(',', '.')) || 0;
                    const repeats = parseInt(row[2]) || 1;
                    const pauseMin = parseFloat(String(row[3]).replace(',', '.')) || 0;
                    const note = row[5] ? row[5].toString().trim() : "";

                    // Intervallkette generieren
                    for (let r = 1; r <= repeats; r++) {
                        const stepName = repeats > 1 ? `${baseName} (Runde ${r}/${repeats})` : baseName;
                        newPlan.steps.push({ name: stepName, duration: workMin * 60, isPause: false, note: note });

                        // Füge eine Pause hinzu, wenn definiert
                        if (pauseMin > 0) {
                            const pauseNote = (r < repeats) ? `Nächste Übung: ${baseName} (Runde ${r + 1}/${repeats})` : "Block beendet. Nächste Übung vorbereiten...";
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

            // Wähle den letzten importierten Reiter aus
            const lastPlanName = workbook.SheetNames[workbook.SheetNames.length - 1];
            if (allPlans[lastPlanName]) {
                planSelector.value = lastPlanName;
                currentPlan = allPlans[lastPlanName];
                renderPlanTable();
            }
            alert("Excel-Datei erfolgreich eingelesen!");
        } catch (err) {
            alert("Fehler beim Lesen der Excel-Datei: " + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Zeichnet die tabellarische Übersicht des Trainingsplans
function renderPlanTable() {
    const tableBody = document.getElementById('plan-table').querySelector('tbody');
    const footer = document.querySelector('footer');
    tableBody.innerHTML = "";

    if (!currentPlan || currentPlan.steps.length === 0) {
        document.getElementById('lbl-total-time').textContent = "Gesamtdauer: 00:00";
        document.getElementById('lbl-end-time').textContent = "Endzeit: --:-- Uhr";
        footer.style.display = 'none'; // Verhindert leere Starts
        return;
    }

    footer.style.display = 'flex'; // Zeigt den Start-Button

    let totalSeconds = currentPlan.steps.reduce((acc, step) => acc + step.duration, 0);

    currentPlan.steps.forEach(step => {
        const tr = document.createElement('tr');
        if (step.isPause) {
            tr.style.backgroundColor = "rgba(41, 128, 185, 0.15)";
        }

        tr.innerHTML = `
            <td>${step.name}</td>
            <td>${formatTime(step.duration)}</td>
            <td>${step.isPause ? "✔️" : "❌"}</td>
            <td style="font-size: 13px; color: #a4b0be;">${step.note}</td>
        `;
        tableBody.appendChild(tr);
    });

    document.getElementById('lbl-total-time').textContent = `Gesamtdauer: ${formatTime(totalSeconds)} Min.`;
    const end = new Date(Date.now() + totalSeconds * 1000);
    document.getElementById('lbl-end-time').textContent = `Endzeit: ${end.getHours().toString().padStart(2, '0')}:${end.getMinutes().toString().padStart(2, '0')} Uhr`;
}

// Hilfsfunktion zur Formatierung von Sekunden in mm:ss
function formatTime(seconds) {
    if (isNaN(seconds) || seconds < 0) seconds = 0;
    const m = Math.floor(seconds / 60).toString().padStart(2, '0');
    const s = Math.round(seconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
}

// --- TIMER-ENGINE (Ablaufsteuerung im Vollbild-Modus) ---

function startTraining() {
    if (!currentPlan || currentPlan.steps.length === 0) return;

    // Wechsle die sichtbaren Container
    document.getElementById('main-view').classList.add('hidden');
    document.getElementById('training-view').classList.remove('hidden');

    timerState.stepIndex = 0;
    timerState.isPaused = false;
    timerState.totalRemaining = currentPlan.steps.reduce((acc, s) => acc + s.duration, 0);

    // Initialer 3-fach Gong vor dem Start
    playStartBellSequence();

    setTimeout(() => {
        runStep();
    }, 1500);
}

function playStartBellSequence() {
    try {
        boxingBell.play();
        setTimeout(() => { boxingBell.play(); }, 400);
        setTimeout(() => { boxingBell.play(); }, 800);
    } catch (e) {
        console.warn("Audio-Wiedergabe blockiert durch Browser-Sicherheitsrichtlinie.");
    }
}

function runStep() {
    if (timerState.stepIndex >= currentPlan.steps.length) {
        clearInterval(activeInterval);
        try { boxingBell.play(); } catch (e) { }
        alert("Training beendet! Hervorragende Leistung!");
        quitTraining(true);
        return;
    }

    const currentStep = currentPlan.steps[timerState.stepIndex];
    timerState.timeLeft = currentStep.duration;
    timerState.countdownBeepPlayed = false;

    // UI-Texte und Farben setzen
    document.getElementById('timer-step-name').textContent = currentStep.name.toUpperCase();
    document.getElementById('timer-note').textContent = currentStep.note || "";
    document.getElementById('training-view').style.backgroundColor = currentStep.isPause ? "#2980b9" : "#c0392b";

    updateNextExercise();

    try {
        boxingBell.play(); // Rundenwechselgong abspielen
    } catch (e) { }

    clearInterval(activeInterval);
    activeInterval = setInterval(timerTick, 1000);
    updateTimerDisplay();
}

function timerTick() {
    if (timerState.isPaused) return;

    // Countdown bei 3, 2, 1 Sekunden vor Schluss piepen lassen
    if (timerState.timeLeft <= 4 && timerState.timeLeft > 1) {
        playSystemBeep(1800, 80);
    }

    if (timerState.timeLeft <= 1) {
        timerState.stepIndex++;
        runStep();
    } else {
        timerState.timeLeft--;
        timerState.totalRemaining--;
        updateTimerDisplay();
    }
}

// Native, asynchrone Tonerzeugung über den Browser (blockiert niemals den Timer)
function playSystemBeep(freq, duration) {
    try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = "sine";
        osc.frequency.value = freq;
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        setTimeout(() => { osc.stop(); ctx.close(); }, duration);
    } catch (e) {
        console.warn("Web-Audio-Beep blockiert.");
    }
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

function togglePause() {
    timerState.isPaused = !timerState.isPaused;
    const btn = document.getElementById('btn-pause');
    btn.textContent = timerState.isPaused ? "▶️ FORTSETZEN" : "⏸️ PAUSE";

    // Hintergrundfarbe im Pausemodus abdunkeln (grau)
    document.getElementById('training-view').style.backgroundColor = timerState.isPaused ? "#7f8c8d" : (currentPlan.steps[timerState.stepIndex].isPause ? "#2980b9" : "#c0392b");
}

function skipStep() {
    timerState.totalRemaining -= timerState.timeLeft;
    timerState.stepIndex++;
    runStep();
}

function quitTraining(force = false) {
    if (force || confirm("Möchtest du das Training wirklich beenden?")) {
        clearInterval(activeInterval);
        document.getElementById('training-view').classList.add('hidden');
        document.getElementById('main-view').classList.remove('hidden');
        renderPlanTable();
    }
}
