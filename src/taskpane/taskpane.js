Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const chatForm = document.getElementById("chat-form");
    const questionInput = document.getElementById("question");
    const responseDiv = document.getElementById("response");
    const button = document.getElementById("ask-button");
    const buttonText = document.getElementById("button-text");
    const spinner = document.getElementById("spinner");
    const historyList = document.getElementById("history-list");
    const clearHistoryButton = document.getElementById("clear-history");
    const logo = document.querySelector(".logo");

    if (logo) {
      logo.addEventListener("error", () => {
        logo.style.display = "none"; 
      });
    }

    // Formular-Submit-Handler
    chatForm.addEventListener("submit", async (event) => {
      event.preventDefault();

      const question = questionInput.value.trim();
      if (!question) {
        showError("Bitte gib eine Frage ein.");
        return;
      }

      
      button.disabled = true;
      buttonText.style.opacity = "0.6";
      spinner.style.display = "inline-block";
      responseDiv.innerText = "Antwort wird generiert…";

      try {
        
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 100000);

        // AI:ssociate API-Anfrage
        const res = await fetch("https://aissociate.at/api/public/v1/chat/ask", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "x-api-key": "ck:4ea46438-b20e-438d-8127-55cd026e794b:c4207d7b-6885-432c-85b5-9ba938410992",
          },
          body: JSON.stringify({
            question: question,
            law: "zivilrechtogh",
            file_context: [],
            file_query_type: "general"

          }),
          signal: controller.signal, // Für Timeout
        });

        clearTimeout(timeoutId); 

        if (!res.ok) {
          const errorText = await res.text(); 
          throw new Error(`API-Fehler: ${res.status} - ${res.statusText} - ${errorText}`);
        }

        // Streaming-Antwort verarbeiten
        const reader = res.body.getReader();
        let answer = "";
        responseDiv.innerText = ""; 

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          // Text aus dem Stream dekodieren
          const chunk = new TextDecoder().decode(value);
          try {
            // Angenommen, die API sendet JSON-Objekte mit einem `text`-Feld pro Event
            const event = JSON.parse(chunk);
            if (event.text) {
              answer += event.text;
              responseDiv.innerText = answer; 
            }
          } catch (e) {
            answer += chunk;
            responseDiv.innerText = answer;
          }
        }

        // Antwort ins Word-Dokument einfügen
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          range.insertText(`Antwort von AI:SSOCIATE:\n${answer}\n`, Word.InsertLocation.end);
          range.font.color = "#00798C";
          await context.sync();
        });

        addToHistory(question, answer);

        questionInput.value = "";
        questionInput.focus();
      } catch (err) {
        // Detaillierte Fehlerbehandlung
        let errorMessage = "Fehler bei der Anfrage: ";
        if (err.name === "AbortError") {
          errorMessage += "Die Anfrage hat zu lange gedauert (Timeout). Bitte überprüfe deine Internetverbindung.";
        } else if (err.message.includes("Failed to fetch")) {
          errorMessage += "Verbindung zur API fehlgeschlagen. Mögliche Ursachen:\n" +
                         "- API-Server nicht erreichbar (überprüfe https://aissociate.at).\n" +
                         "- CORS-Problem (Server erlaubt Anfragen von localhost nicht).\n" +
                         "- Ungültiger API-Schlüssel.\n" +
                         "Details: " + err.message;
        } else {
          errorMessage += err.message;
        }
        showError(errorMessage);
        console.error("API Error:", err);
      } finally {
        // Button und Spinner zurücksetzen
        button.disabled = false;
        buttonText.style.opacity = "1";
        spinner.style.display = "none";
      }
    });

    clearHistoryButton.addEventListener("click", () => {
      historyList.innerHTML = "";
      responseDiv.innerText = "Verlauf gelöscht.";
      setTimeout(() => {
        responseDiv.innerText = "Noch keine Antwort.";
      }, 2000);
    });

    function showError(message) {
      responseDiv.innerText = message;
      responseDiv.classList.add("error");
      setTimeout(() => {
        responseDiv.classList.remove("error");
        responseDiv.innerText = "Noch keine Antwort.";
      }, 5000); 
    }

    function addToHistory(question, answer) {
      const entry = document.createElement("li");
      entry.innerHTML = `<strong>Frage:</strong> ${sanitizeHTML(question)}<br><strong>Antwort:</strong> ${sanitizeHTML(answer)}`;
      historyList.prepend(entry);
      if (historyList.children.length > 10) {
        historyList.removeChild(historyList.lastChild);
      }
    }

    // HTML-Injection verhindern
    function sanitizeHTML(str) {
      const div = document.createElement("div");
      div.textContent = str;
      return div.innerHTML;
    }
  }
});
          