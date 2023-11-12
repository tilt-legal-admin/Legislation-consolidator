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
  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      const rangeXml = range.getOoxml();
      await context.sync();
      return rangeXml.value;
    });
  } catch (error) {
    throw new Error(`Error in getTextFromWord: ${error.message}`);
  }
}

async function processTextWithOpenAI(text: string, instructions: string): Promise<string> {
  const apiKey = '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';
  const url = "https://api.openai.com/v1/chat/completions";

  try {
    const systemPrompt = "You are an AI assistant that has been integrated as a Microsoft Word add-in. You will receive instructions from the user on how to amend a section of text, followed by the text chunk you need to amend in OOXML. Respond with only the amended OOXML without any additional commentary or guidance.";
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
        model: "gpt-3.5-turbo",
        max_tokens: 5000,
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


async function applyChangesToWord(ooxml: string): Promise<void> {
  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertOoxml(ooxml, "Replace");
      await context.sync();
    });
  } catch (error) {
    throw new Error(`Error in applyChangesToWord: ${error.message}`);
  }
}

async function processText() {
  try {
      console.log("processText started"); // Log when the function starts
      const instructions = (document.getElementById("instructionField") as HTMLTextAreaElement).value;
      console.log("Instructions:", instructions); // Log the instructions

      console.log("Calling getTextFromWord");
      const ooxml = await getTextFromWord();
      console.log("Received OOXML:", ooxml); // Log the received OOXML

      console.log("Calling processTextWithOpenAI with OOXML and instructions");
      const processedText = await processTextWithOpenAI(ooxml, instructions);
      console.log("Processed Text:", processedText); // Log the processed text

      console.log("Calling applyChangesToWord with processed text");
      await applyChangesToWord(processedText);
      console.log("Changes applied successfully");
  } catch (error) {
      console.error("Error processing text: ", error); // Log any errors
  }
}

