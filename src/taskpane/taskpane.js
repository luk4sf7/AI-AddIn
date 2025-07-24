// taskpane.js
// Runs in the task-pane (browser) context.

/* ==============================
   🔧  Configuration
   ============================== */
const BASE_URL = 'https://aissociate.at'; // Don't use process.env in browser unless you inject it at build-time
const API_KEY  = 'ck:4ea46438-b20e-438d-8127-55cd026e794b:c4207d7b-6885-432c-85b5-9ba938410992';


const QUESTION   = 'Wann haftet der Geschäftsführer einer GmbH?';
const LEGAL_AREA = 'zivilrechtogh';
const SCOPE      = null;

/* ==============================
   🛫  Office initialization
   ============================== */
Office.onReady(() => {
  const form           = document.getElementById('chat-form');
  const textarea       = document.getElementById('question');
  const askButton      = document.getElementById('ask-button');
  const buttonText     = document.getElementById('button-text');
  const spinner        = document.getElementById('spinner');
  const responseBox    = document.getElementById('response');

  // Optional: hooks for history (if you have these elements in HTML)
  const historyList    = document.getElementById('history-list');
  const clearHistoryBt = document.getElementById('clear-history');

  /* ---------- Form submit handler ---------- */
  form.addEventListener('submit', async (evt) => {
    evt.preventDefault();
    const question = textarea.value.trim();
    if (!question) return;

    toggleBusy(true);
    responseBox.textContent = '';
    responseBox.classList.remove('error');

    try {
      const answer = await streamAsk(question, { responseEl: responseBox });
      // Optional: Add to history & save
      if (historyList) addHistoryEntry({ q: question, a: answer });
      if (typeof saveHistory === 'function') saveHistory();
      // Optional: Insert answer into Word document
      insertToDocument(answer);
    } catch (err) {
      console.error(err);
      responseBox.textContent = `Fehler: ${err.message}`;
      responseBox.classList.add('error');
    } finally {
      toggleBusy(false);
      textarea.value = '';
      textarea.focus();
    }
  });

  /** Show / hide spinner & disable button */
  function toggleBusy(isBusy) {
    askButton.disabled = isBusy;
    spinner.style.display = isBusy ? 'inline-block' : 'none';
    buttonText.textContent = isBusy ? 'Wird generiert …' : 'Antwort generieren';
  }

  /**
   * Streams an “ask” request and writes to responseEl progressively.
   * Returns the full text at the end.
   */
  async function streamAsk(question, { legalArea = null, scope = null, files = [], responseEl } = {}) {
    if (!API_KEY) {
      throw new Error('AISSOCIATE_API_KEY fehlt oder ist leer.');
    }
    console.log(question);

    const endpoint = `${BASE_URL}/api/public/v1/chat/ask`;
    const payload  = {
      question,
      law:      legalArea,
      sub_law:  scope,
      file_context: files,
      file_query_type: 'general',
    };
    console.log(payload);

    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': API_KEY,
        },
        body: JSON.stringify(payload),
        });
    } catch (err) {
      // This block triggers when the browser can't even start or finish the request
      // (CORS blocked, DNS, SSL, network offline, mixed content, etc.)
      if (err instanceof TypeError && /Failed to fetch/i.test(err.message)) {
        throw new Error(
          [
            'Netzwerkfehler: „Failed to fetch“.',
            'Ursachen können sein:',
            '- CORS blockiert (API sendet keinen Access-Control-Allow-Origin Header)',
            '- Domain nicht in AppDomains in der Office-Manifest-Datei',
            '- HTTPS/SSL-Problem (Zertifikat nicht vertraut)',
            '- Falsche URL/Endpoint',
            '- Firmen-Firewall/Proxy blockiert',
            '',
            'Prüfe die Browser-Netzwerk-Konsole (DevTools → Network) & die Manifest-Datei.'
          ].join('\n')
        );
      }
      // Otherwise rethrow
      throw err;
    }


    if (!res.ok || !res.body) {
      const t = await res.text().catch(() => '');
      throw new Error(`Request failed (${res.status}): ${t}`);
    }

    const reader  = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer    = '';
    let fullText  = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      // Split by blank line delimiter (end of SSE block)
      const parts = buffer.split(/\r?\n\r?\n/);
      buffer = parts.pop(); // keep remainder

      for (const rawBlock of parts) {
        if (!rawBlock.trim()) continue;

        const parsed = parseSSEBlock(rawBlock);
        if (!parsed) continue;

        if (parsed.type === 'ERROR') {
          throw new Error(parsed.text || 'Unbekannter API-Fehler');
        }

        if (parsed.type === 'message' && typeof parsed.text === 'string') {
          fullText += parsed.text;
          if (responseEl) {
            responseEl.textContent += parsed.text;
            responseEl.scrollTop = responseEl.scrollHeight;
          }
        }
      }
    }
    console.log('---');
    console.log(question);
    console.log(legalArea);
    console.log('------');

    return fullText.trim();
  }

  /**
   * Parse a single SSE block (delimited by blank lines).
   */
  function parseSSEBlock(block) {
    const lines = block.split(/\r?\n/);
    let dataStr = '';
    for (const line of lines) {
      if (line.startsWith('data:')) {
        dataStr += (dataStr ? '\n' : '') + line.slice(5).trim();
      }
    }
    if (!dataStr) return null;
    try {
      return JSON.parse(dataStr);
    } catch {
      return null;
    }
  }

  /* ------- Optional helpers for Office insertion & history ------- */

  function insertToDocument(text) {
    if (!Office.context || !Office.context.document) return;
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn('Failed to insert into doc:', result.error);
      }
    });
  }

  function addHistoryEntry(entry) {
    if (!historyList) return;
    const li = document.createElement('li');
    li.innerHTML = `<strong>Q:</strong> ${entry.q}<br><strong>A:</strong> ${entry.a}`;
    historyList.prepend(li);
  }

  function saveHistory() {
    if (!historyList) return;
    const items = Array.from(historyList.children).map(li => li.innerText);
    localStorage.setItem('chatHistory', JSON.stringify(items));
  }

  if (clearHistoryBt) {
    clearHistoryBt.addEventListener('click', () => {
      if (historyList) historyList.innerHTML = '';
      localStorage.removeItem('chatHistory');
    });
  }

}); // <-- Important: close Office.onReady
