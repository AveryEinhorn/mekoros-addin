// Placeholder JS for Option A ghost letters
// In practice this connects to Sefaria API for suggestions
Office.onReady(() => {
  console.log("Mekoros Add-in loaded.");
});

// Example function: insert ghost text into Word
function insertGhostText(text) {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

// This is just a placeholder; real Sefaria connection requires hosting
