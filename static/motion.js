/* ============================================================
   Profiling System — Motion Studio 9:16 Engine
   ============================================================ */

(function () {
  "use strict";

  // Scene Configuration (Duration in seconds)
  const SCENES = [
    { id: 1, duration: 7, title: "Problem Hook" },
    { id: 2, duration: 7, title: "Web Solution" },
    { id: 3, duration: 8, title: "Smart Multi-Sheet" },
    { id: 4, duration: 10, title: "Instant Table Automation" },
    { id: 5, duration: 8, title: "Download & CTA" }
  ];

  const TOTAL_DURATION = SCENES.reduce((acc, s) => acc + s.duration, 0); // 40 seconds

  // DOM Elements
  const btnTogglePlay = document.getElementById("btn-toggle-play");
  const playIcon = document.getElementById("play-icon");
  const playText = document.getElementById("play-text");
  const btnRestart = document.getElementById("btn-restart");
  const btnSound = document.getElementById("btn-sound");
  const btnRecordMode = document.getElementById("btn-record-mode");
  const timelineItems = document.querySelectorAll(".timeline-item");
  const mockTableRows = document.getElementById("mock-table-rows");

  let isPlaying = false;
  let isSoundOn = true;
  let currentTime = 0; // in seconds
  let currentSceneIndex = 0;
  let animationFrameId = null;
  let lastTimestamp = null;

  // Real Project Data for Scene 4
  const SAMPLE_DEALERS = [
    { name: "Berkah Saudara Solok", leads: 493, kacab: "0.00%", flp: "0.00%", cr: "55.58%" },
    { name: "Hayati Pasaman", leads: 444, kacab: "0.00%", flp: "0.45%", cr: "11.49%" },
    { name: "Sahabat Motor Solok", leads: 327, kacab: "0.00%", flp: "0.92%", cr: "45.87%" },
    { name: "Mina SPN Motor", leads: 202, kacab: "0.00%", flp: "0.50%", cr: "76.24%" },
    { name: "Cempaka Muaro", leads: 189, kacab: "0.00%", flp: "0.00%", cr: "69.31%" }
  ];

  // ============================================================
  // Web Audio API Synthesizer (No external MP3 files needed)
  // ============================================================
  let audioCtx = null;

  function initAudio() {
    if (!audioCtx) {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      audioCtx = new AudioContext();
    }
    if (audioCtx.state === "suspended") {
      audioCtx.resume();
    }
  }

  function playTone(freq, type = "sine", duration = 0.2, gainVal = 0.1) {
    if (!isSoundOn || !audioCtx) return;
    try {
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.type = type;
      osc.frequency.setValueAtTime(freq, audioCtx.currentTime);
      gain.gain.setValueAtTime(gainVal, audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.0001, audioCtx.currentTime + duration);
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.start();
      osc.stop(audioCtx.currentTime + duration);
    } catch (e) {
      console.warn(e);
    }
  }

  function playSfx(type) {
    if (!isSoundOn) return;
    initAudio();
    if (type === "transition") {
      playTone(440, "sine", 0.3, 0.08);
      setTimeout(() => playTone(660, "sine", 0.4, 0.08), 80);
    } else if (type === "scan") {
      playTone(880, "triangle", 0.15, 0.05);
    } else if (type === "row") {
      playTone(520, "sine", 0.1, 0.06);
    } else if (type === "success") {
      playTone(523.25, "sine", 0.2, 0.1); // C5
      setTimeout(() => playTone(659.25, "sine", 0.2, 0.1), 100); // E5
      setTimeout(() => playTone(783.99, "sine", 0.4, 0.12), 200); // G5
    }
  }

  // ============================================================
  // Scene Timeline Engine
  // ============================================================
  function getSceneAtTime(t) {
    let accumulated = 0;
    for (let i = 0; i < SCENES.length; i++) {
      accumulated += SCENES[i].duration;
      if (t < accumulated) {
        const sceneStartTime = accumulated - SCENES[i].duration;
        const progressInScene = (t - sceneStartTime) / SCENES[i].duration;
        return { index: i, scene: SCENES[i], progressInScene, sceneStartTime };
      }
    }
    return { index: SCENES.length - 1, scene: SCENES[SCENES.length - 1], progressInScene: 1, sceneStartTime: TOTAL_DURATION - SCENES[SCENES.length - 1].duration };
  }

  function updateVisuals(time) {
    const info = getSceneAtTime(time);

    // Switch scene visibility if changed
    if (info.index !== currentSceneIndex) {
      switchScene(info.index);
    }

    // Update Progress Bars (1 to 5)
    SCENES.forEach((s, idx) => {
      const fillEl = document.getElementById(`fill-${idx + 1}`);
      if (!fillEl) return;
      if (idx < info.index) {
        fillEl.style.width = "100%";
      } else if (idx === info.index) {
        fillEl.style.width = `${Math.min(100, Math.max(0, info.progressInScene * 100))}%`;
      } else {
        fillEl.style.width = "0%";
      }
    });

    // Special animations based on scene progress
    if (info.index === 3) {
      // Scene 4: Animate Table Rows populating
      renderAnimatedRows(info.progressInScene);
    }
  }

  function switchScene(newIndex) {
    // Hide all scenes
    SCENES.forEach((s) => {
      const el = document.getElementById(`scene-${s.id}`);
      if (el) el.classList.add("hidden");
    });

    // Show current scene
    const targetEl = document.getElementById(`scene-${SCENES[newIndex].id}`);
    if (targetEl) targetEl.classList.remove("hidden");

    // Update Sidebar
    timelineItems.forEach((item, idx) => {
      item.classList.toggle("active", idx === newIndex);
    });

    currentSceneIndex = newIndex;

    // Trigger Sound
    if (newIndex === 2) {
      playSfx("scan");
    } else if (newIndex === 4) {
      playSfx("success");
    } else {
      playSfx("transition");
    }
  }

  function renderAnimatedRows(progress) {
    if (!mockTableRows) return;
    const totalRows = SAMPLE_DEALERS.length;
    const rowsToShow = Math.floor(progress * (totalRows + 1));

    let html = "";
    for (let i = 0; i < Math.min(totalRows, rowsToShow); i++) {
      const d = SAMPLE_DEALERS[i];
      html += `
        <div class="mock-row-data">
          <span>${d.name}</span>
          <span>${d.leads}</span>
          <span>${d.kacab}</span>
          <span>${d.flp}</span>
          <span style="font-weight:700; color:#0284c7;">${d.cr}</span>
        </div>
      `;
    }
    mockTableRows.innerHTML = html;
  }

  // Animation Loop
  function tick(timestamp) {
    if (!isPlaying) return;

    if (!lastTimestamp) lastTimestamp = timestamp;
    const delta = (timestamp - lastTimestamp) / 1000;
    lastTimestamp = timestamp;

    currentTime += delta;

    if (currentTime >= TOTAL_DURATION) {
      currentTime = TOTAL_DURATION;
      updateVisuals(currentTime);
      pause();
      return;
    }

    updateVisuals(currentTime);
    animationFrameId = requestAnimationFrame(tick);
  }

  function play() {
    initAudio();
    if (currentTime >= TOTAL_DURATION) {
      currentTime = 0;
    }
    isPlaying = true;
    lastTimestamp = null;
    playIcon.textContent = "⏸";
    playText.textContent = "Pause";
    btnTogglePlay.classList.add("playing");
    animationFrameId = requestAnimationFrame(tick);
  }

  function pause() {
    isPlaying = false;
    lastTimestamp = null;
    playIcon.textContent = "▶";
    playText.textContent = "Play";
    btnTogglePlay.classList.remove("playing");
    if (animationFrameId) {
      cancelAnimationFrame(animationFrameId);
      animationFrameId = null;
    }
  }

  function togglePlay() {
    if (isPlaying) {
      pause();
    } else {
      play();
    }
  }

  function restart() {
    pause();
    currentTime = 0;
    switchScene(0);
    updateVisuals(0);
    play();
  }

  function jumpToScene(sceneIdx) {
    let targetTime = 0;
    for (let i = 0; i < sceneIdx; i++) {
      targetTime += SCENES[i].duration;
    }
    currentTime = targetTime;
    updateVisuals(currentTime);
    if (!isPlaying) {
      play();
    }
  }

  // ============================================================
  // Event Listeners
  // ============================================================
  btnTogglePlay.addEventListener("click", togglePlay);
  btnRestart.addEventListener("click", restart);

  btnSound.addEventListener("click", () => {
    isSoundOn = !isSoundOn;
    btnSound.textContent = isSoundOn ? "🔊 Sound: ON" : "🔇 Sound: OFF";
    if (isSoundOn) playTone(520, "sine", 0.1);
  });

  // Toggle Screen Recording Mode
  btnRecordMode.addEventListener("click", () => {
    document.body.classList.toggle("recording-mode");
    if (document.body.classList.contains("recording-mode")) {
      btnRecordMode.textContent = "❌ Keluar Mode Rekam (Esc)";
    } else {
      btnRecordMode.textContent = "🎥 Siap Rekam (Fokus)";
    }
  });

  // Keyboard Shortcuts
  window.addEventListener("keydown", (e) => {
    if (e.code === "Space") {
      e.preventDefault();
      togglePlay();
    } else if (e.code === "KeyR") {
      restart();
    } else if (e.code === "Escape") {
      document.body.classList.remove("recording-mode");
      btnRecordMode.textContent = "🎥 Siap Rekam (Fokus)";
    }
  });

  // Sidebar scene clicks
  timelineItems.forEach((item) => {
    item.addEventListener("click", () => {
      const idx = parseInt(item.getAttribute("data-scene"), 10) - 1;
      jumpToScene(idx);
    });
  });

  // Initial State Setup
  switchScene(0);
  updateVisuals(0);

  // Auto-play preview after 1 second on load
  setTimeout(() => {
    play();
  }, 800);

})();
