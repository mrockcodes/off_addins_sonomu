// taskpane.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Perform initialization tasks.
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const tone = document.getElementById("tone").value;
    const transformedText = await transformText(selection.text, tone);
    selection.insertText(transformedText, Word.InsertLocation.replace);
  });
}

async function transformText(text, tone) {
  const apiKey = 'your-openai-api-key-here';
  const prompt = `Transform the following text into a ${tone} tone:\n\n${text}`;

  try {
    const response = await axios.post('https://api.openai.com/v1/completions', {
      model: 'text-davinci-003',
      prompt: prompt,
      max_tokens: 150,
    }, {
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    });

    return response.data.choices[0].text.trim();
  } catch (error) {
    console.error(error);
    return "Error transforming text.";
  }
}
