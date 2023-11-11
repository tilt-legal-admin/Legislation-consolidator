/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Replace or add the new button's event listener
    document.getElementById("processButton").onclick = processText;
  }
});

async function getTextFromWord(): Promise<string> {
  return Word.run(async (context) => {
      // Get the current selection
      const range = context.document.getSelection();
      range.load("text");

      // Synchronize the document state
      await context.sync();

      return range.text;
  });
}

async function processTextWithOpenAI(text: string, instructions: string): Promise<string> {
  const apiKey = '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'; 
  const url = "https://api.openai.com/v1/chat/completions";

  try {
    const systemPrompt = "You are an AI assistant that has been integrated as a Microsoft Word add-in. You will receive instructions from the user on how to amend a section of text, followed by the text chunk you need to amend. Respond with only the amended text without any additional commentary or guidance.";
    const userPrompt = `Instructions: ${instructions}\n\n Text: ${text}`;  
    const response = await fetch(url, {
          method: "POST",
          headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${apiKey}`,
          },
          body: JSON.stringify({
              messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }],
              temperature: 0.6,
              model: "gpt-4-1106-preview",
              max_tokens: 1000,
          }),
      });

      const responseData = await response.json();
      // Process the response to extract the text you need
      const content = responseData.choices[0]?.message?.content.trim();
      return content || ''; // Return the processed text or an empty string if no content
  } catch (error) {
      console.error('Error processing text with OpenAI:', error);
      throw error;
  }
}


async function applyChangesToWord(newText: string): Promise<void> {
  return Word.run(async (context) => {
      // Replace the selected text with the new text
      const range = context.document.getSelection();
      range.insertText(newText, "Replace");

      await context.sync();
  });
}


async function processText() {
    try {
        const instructions = (document.getElementById("instructionField") as HTMLTextAreaElement).value;
        const textFromWord = await getTextFromWord();

        const processedText = await processTextWithOpenAI(textFromWord, instructions);
        await applyChangesToWord(processedText);
    } catch (error) {
        console.error("Error processing text: ", error);
    }
}
