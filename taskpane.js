Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Mekoros Option 1 ready");
        const btn = document.getElementById("suggestButton");
        btn.addEventListener("click", insertSefariaSource);
    }
});

async function insertSefariaSource() {
    try {
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load("text");
            await context.sync();

            const typedText = range.text.trim();
            if (!typedText) return alert("Please select or type Hebrew words first.");

            const suggestion = await getSefariaSuggestion(typedText);
            if (!suggestion) return alert("No source found.");

            const fullText = suggestion.text + " (" + suggestion.sefer + ")";
            range.insertText(fullText, Word.InsertLocation.replace);

            await context.sync();
        });
    } catch (e) {
        console.error(e);
        alert("Error inserting source: " + e.message);
    }
}

async function getSefariaSuggestion(query) {
    try {
        const searchUrl = `https://www.sefaria.org/api/search?query=${encodeURIComponent(query)}&size=1`;
        const searchResp = await fetch(searchUrl);
        const searchData = await searchResp.json();

        if (!searchData.hits || searchData.hits.length === 0) return null;

        const firstHit = searchData.hits[0];
        const ref = firstHit.ref;
        const seferName = ref.split(":")[0] || "ספֵר";

        const textUrl = `https://www.sefaria.org/api/texts/${encodeURIComponent(ref)}?context=0`;
        const textResp = await fetch(textUrl);
        const textData = await textResp.json();

        if (!textData.text) return null;

        const suggestionText = textData.text.join(" ").split(" ").slice(0, 50).join(" ");
        return { text: suggestionText, sefer: seferName };
    } catch (e) {
        console.error("Sefaria fetch error:", e);
        return null;
    }
}
