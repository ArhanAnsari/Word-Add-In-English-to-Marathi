Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
	document.getElementById('convertText').onclick = convertText;
    }
});

async function convertText() {
    await Word.run(async context => {
	const document = context.document;
	const selection = document.getSelection();
	selection.load('text');
	await context.sync();

	const translatedText = await translateText(selection.text);
	selection.insertText(translatedText, Word.InsertLocation.replace);
	await context.sync();
    });
}

async function translateText(text) {
    const response = await fetch('https://libretranslate.com/translate', {
	method: 'POST',
	headers: {
	    'Content-Type': 'application/json'
	},
	body: JSON.stringify({
	    q: text,
	    source: 'en', // English
	    target: 'mr', // Marathi
	    format: 'text'
	})
    });

    if (!response.ok) {
	return 'Error: Failed to fetch translation.';
    }

    const data = await response.json();
    return data.translatedText;
}
