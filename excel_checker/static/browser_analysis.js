/* Steuert den Analyse-Worker: Lebenszyklus, Vorwärmen, Laufzeitschätzung.
 *
 * Die Datei verlässt den Browser nicht. Es gibt keinen Upload und keinen
 * Report-Endpunkt — deshalb auch keine Session-IDs, keine Server-Zustände
 * und nichts, was nach dem Seitenaufbau noch zum Server ginge.
 */
(function (global) {
  'use strict';

  const VENDOR_BASE = '/static/vendor';
  const WORKER_URL = '/static/analysis_worker.js';

  let warmWorker = null;      // vorgewärmt, wartet auf die erste Analyse
  let warmReady = null;       // Promise, das mit dem vorgewärmten Worker auflöst

  function spawn() {
    // type:'module' ist Pflicht — Pyodide laeuft nicht im klassischen Worker.
    const worker = new Worker(WORKER_URL, { type: 'module' });
    const ready = new Promise((resolve, reject) => {
      worker.addEventListener('message', function onMsg(e) {
        if (e.data.type === 'ready') {
          worker.removeEventListener('message', onMsg);
          resolve(worker);
        } else if (e.data.type === 'error') {
          worker.removeEventListener('message', onMsg);
          reject(new Error(e.data.message));
        }
      });
      worker.addEventListener('error', (e) => reject(new Error(e.message || 'Worker-Fehler')));
    });
    worker.postMessage({ type: 'init', base: VENDOR_BASE });
    return { worker: worker, ready: ready };
  }

  /** Lädt Pyodide schon beim Seitenaufbau, damit der erste Klick nicht wartet. */
  function warmUp(onBoot) {
    if (warmReady) return warmReady;
    const spawned = spawn();
    warmWorker = spawned.worker;
    if (onBoot) {
      warmWorker.addEventListener('message', (e) => {
        if (e.data.type === 'boot') onBoot(e.data.step);
      });
    }
    warmReady = spawned.ready.catch((err) => {
      // Fehlgeschlagenes Vorwärmen darf die Seite nicht blockieren — beim
      // nächsten Versuch wird ein frischer Worker gestartet.
      warmWorker = null;
      warmReady = null;
      throw err;
    });
    return warmReady;
  }

  /**
   * Analysiert eine Datei. Nimmt den vorgewärmten Worker, wenn es einen gibt,
   * und beendet ihn danach immer — der WASM-Heap gibt Speicher nie zurück.
   */
  async function analyseFile(file, handlers) {
    handlers = handlers || {};
    let worker;
    if (warmWorker && warmReady) {
      const pending = warmReady;
      worker = warmWorker;
      warmWorker = null;
      warmReady = null;
      await pending;
    } else {
      const spawned = spawn();
      worker = spawned.worker;
      if (handlers.onBoot) {
        worker.addEventListener('message', (e) => {
          if (e.data.type === 'boot') handlers.onBoot(e.data.step);
        });
      }
      await spawned.ready;
    }

    const buffer = await file.arrayBuffer();

    try {
      return await new Promise((resolve, reject) => {
        worker.addEventListener('message', (e) => {
          const msg = e.data;
          if (msg.type === 'progress') {
            if (handlers.onProgress) handlers.onProgress(msg.event);
          } else if (msg.type === 'report') {
            resolve(msg);
          } else if (msg.type === 'error') {
            reject(new Error(msg.message + (msg.detail ? '\n' + msg.detail : '')));
          }
        });
        worker.addEventListener('error', (e) => reject(new Error(e.message || 'Worker-Fehler')));
        // Der Puffer wird übergeben, nicht kopiert — bei 40 MB spart das
        // eine zweite Kopie im Speicher.
        worker.postMessage({ type: 'analyze', buffer: buffer, filename: file.name }, [buffer]);
      });
    } finally {
      worker.terminate();
      // Nächste Analyse soll wieder auf einen warmen Worker treffen — aber
      // ohne onBoot: dessen Meldungen würden sonst die Ergebniszeile der
      // gerade fertigen Analyse überschreiben.
      warmUp();
    }
  }

  /**
   * Ehrliche Laufzeitschätzung aus der Dateigröße.
   *
   * Stützpunkte aus bench/README.md (Pyodide, aus BytesIO): 11 KB → unter
   * 1 s, 44,4 MB → 166 s. Das ist grob linear in der Dateigröße; die
   * Konstante deckt Boot und Report-Aufbau ab. Lieber etwas zu großzügig
   * schätzen als den Nutzer überraschen.
   */
  function estimateSeconds(bytes) {
    const mb = bytes / 1048576;
    return Math.max(2, Math.round(2 + mb * 3.7));
  }

  function formatEstimate(bytes) {
    const secs = estimateSeconds(bytes);
    if (secs < 60) return 'etwa ' + secs + ' Sekunden';
    const mins = Math.round(secs / 60);
    return 'etwa ' + mins + (mins === 1 ? ' Minute' : ' Minuten');
  }

  global.BrowserAnalysis = {
    warmUp: warmUp,
    analyseFile: analyseFile,
    estimateSeconds: estimateSeconds,
    formatEstimate: formatEstimate,
  };
})(window);
