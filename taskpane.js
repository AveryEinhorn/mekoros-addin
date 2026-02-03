Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Mekoros Autocomplete loaded");
        initAutocomplete();
    }
});

function initAutocomplete() {
    // Attach event to selection change in Word
    Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        // Start listening for changes
        context.document.onSelectionChanged.add(onSelectionChange);
        await context.sync();
    });
}

async function onSelectionChange() {
    Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();

        const typedText = range.text.trim();

        if (!typedText || !isHebrew(typedText)) return;

        const suggestion = await getSefariaSuggestion(typedText);
        if (!suggestion) return;

        // Insert ghost letters in light gray
        insertGhostLetters(suggestion);
    });
}

//
