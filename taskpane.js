Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("suggestButton").onclick = insertSource;
        console.log("Mekoros Autocomplete ready (Option 1)");
    }
});

async function insertSource() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();

            const typedText = range.text.trim();
            if (!typedText) return;

            const suggestion = await getSefariaSuggestion(typedText);
            if (!suggestion) {
                console.log("No source found for:", typedText);
                return;
            }

            // Insert the source in Word
            const fullText = suggestion.text + " (" + suggestion.sefer + ")";
            range.insertText(fullText, Word.InsertLocation.replace);

            await context.sync();
        });
    } catch (e) {
        console.error(e);
    }
}

// Simple check for Hebrew letters (not strictly necessary here)
function isHebrew(text) {
    return /[\u0590-\u05FF]/.test(text);
}

// Fetch suggestion from Sefaria API
async function getSefariaSuggestion(query) {
    try {
        const searchUrl = `https://www.sefaria.org/api/search?query=${encodeURIComponent(query)}&size=1`;
        const searchResp = await fetch(searchUrl);
        const searchData = await searchResp.json();

        if (!searchData.hits || searchData.hits.length === 0) return null;

        const firstHit = searchData.hits[0];
        const ref = firstHit.ref;
        const seferName = ref.split(":")[0] || "ספֵר";

        // Fetch the actual text
        const textUrl = `https://www.sefaria.org/api/texts/${encodeURIComponent(ref)}?context=0`;
        const textResp = await fetch(textUrl);
        const textData = await textResp.json();

        if (!textData.text) return null;

        // Join first 50 words
        const suggestionText = textData.text.join(" ").split(" ").slice(0, 50).join(" ");

        return { text: suggestionText, sefer: seferName };
    } catch (e) {
        console.error("Sefaria fetch error:", e);
        return null;
    }
}
