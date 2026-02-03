Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    console.log("Mekoros Option A loaded");
    initAutocomplete();
  }
});

let typingTimer;
let lastText = "";

function initAutocomplete() {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    // Listen for typing events in Word
    setInterval(checkTyping, 500); // Poll every 0.5s
  });
}

async function checkTyping() {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    const typedText = range.text.trim();

    if (!typedText || !isHebrew(typedText) || typedText === lastText) return;

    lastText = typedText;

    const suggestion = await getSefariaSuggestion(typedText);
    if (!suggestion) return;

    insertGhostLetters(suggestion);
  });
}

// Simple check for Hebrew characters
function isHebrew(text) {
  return /[\u0590-\u05FF]/.test(text);
}

// Insert ghost letters inline
function insertGhostLetters(suggestion) {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    const ghostText = suggestion.text + " (" + suggestion.sefer + ")";
    range.insertText(ghostText, Word.InsertLocation.replace);
    range.font.color = "#A0A0A0"; // light gray
    await context.sync();
  });
}

// Query Sefaria search API
async function getSefariaSuggestion(query) {
  try {
    const searchUrl = `https://www.sefaria.org/api/search?query=${encodeURIComponent(query)}&size=1`;
    const searchResp = await fetch(searchUrl);
    const searchData = await searchResp.json();

    if (!searchData?.hits || searchData.hits.length === 0) return null;

    const firstHit = searchData.hits[0];
    const ref = firstHit.ref;
    const seferName = ref.split(":")[0] || "ספֵר";

    // Fetch the actual text
    const textUrl = `https://www.sefaria.org/api/texts/${encodeURIComponent(ref)}?context=0`;
    const textResp = await fetch(textUrl);
    const textData = await textResp.json();

    if (!textData?.text) return null;

    // Join first 50 words as suggestion
    const suggestionText = textData.text.join(" ").split(" ").slice(0, 50).join(" ");

    return { text: suggestionText, sefer: seferName };
  } catch (e) {
    console.error("Sefaria fetch error:", e);
    return null;
  }
}
