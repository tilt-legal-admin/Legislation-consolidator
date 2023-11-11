/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Replace or add the new button's event listener
    document.getElementById("processButton").onclick = processText;
  }
});

interface TextWithStyle {
  text: string;
  style: Word.Font;
}

async function getTextFromWord(): Promise<TextWithStyle> {
  return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.font.load('name, size, color, bold, italic, underline');

      await context.sync();

      return {
          text: range.text,
          style: range.font
      };
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


async function applyChangesToWord(newText: string, originalStyle: Word.Font): Promise<void> {
  return Word.run(async (context) => {
      const range = context.document.getSelection();

      range.clear(); // Clears the selected content
      const insertedRange = range.insertText(newText, "Replace");

      // Apply the original style
      insertedRange.font.name = originalStyle.name;
      insertedRange.font.size = originalStyle.size;
      insertedRange.font.color = originalStyle.color;
      insertedRange.font.bold = originalStyle.bold;
      insertedRange.font.italic = originalStyle.italic;
      insertedRange.font.underline = originalStyle.underline;

      await context.sync();
  });
}

async function processText() {
  try {
      const instructions = (document.getElementById("instructionField") as HTMLTextAreaElement).value;
      const { text, style } = await getTextFromWord();

      const processedText = await processTextWithOpenAI(text, instructions);
      await applyChangesToWord(processedText, style);
  } catch (error) {
      console.error("Error processing text: ", error);
  }
}

