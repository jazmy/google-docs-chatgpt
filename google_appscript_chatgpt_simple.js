// Google Docs ChatGPT Integration 
// This Google Apps Script code adds a custom menu item to a Google Doc called "ChatGPT Menu" 
// which allows the user to prompt a GPT-3.5-turbo model by entering text into a dialog box.
// The prompt is passed to the OpenAI API. 
// The API generates a response, which is then added as a paragraph to the document.
// Author: Jasmine Robinson
// Date: 2023-03-19

//--------------------------------------
// API Key
//-------------------------------------- 
const apiKey = "Your API Key";
const apiUrl = "https://api.openai.com/v1/chat/completions";

//--------------------------------------
// Create menu items in the Google Doc
//--------------------------------------
function onOpen() {
    DocumentApp.getUi()
        .createMenu("ChatGPT Menu")
        .addItem("New Prompt", "gptgetInput")
        .addToUi();
}

//--------------------------------------
// GPT Prompts Functions
//-------------------------------------- 
async function gptgetInput() {
    const html = HtmlService.createHtmlOutput(`
    <style>
      .input-box {
        height: 150px;
        width: 450px;
      }
    </style>
    <div>
      <textarea id="input" class="input-box" oninput="updateCount()"></textarea>
      <br>
      <span id="count"></span>
      <span id="message"></span>
      <br>
      <button id="submitButton" onclick="submitInput()">Submit</button>
    </div>
    <div id="loading" style="display:none; text-align:center">
      <p>Please wait while the content is generated, may take a few minutes...</p>
      <img src="https://www.google.com/images/spin-32.gif">
    </div>
    <script>
      function submitInput() {
        const userInput = document.getElementById("input").value;
        const loadingDiv = document.getElementById("loading");
        loadingDiv.style.display = "block";
        google.script.run.withSuccessHandler(closeDialog).processgptInput(userInput);
      }

      function closeDialog() {
        google.script.host.close();
      }

      function updateCount() {
        const input = document.getElementById("input");
        const count = input.value.length;
        const countSpan = document.getElementById("count");
        const message = document.getElementById("message");
        const button = document.getElementById("submitButton");
        if (count > 4000) {
          countSpan.style.color = "red";
          message.innerHTML = "Character count exceeds limit of 4000";
          button.style.display = "none";
        } else {
          countSpan.style.color = "black";
          message.innerHTML = "";
          button.style.display = "block";
        }
        countSpan.innerHTML = count + "/4000";
      }
    </script>
  `);

    DocumentApp.getUi().showModalDialog(
        html,
        "Type in your prompt"
    );
}

// Take the value from the modal dialog box and pass it to the OpenAI API
async function processgptInput(userInput) {
    console.log("User Input:", userInput);

    // Get the active document and body
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    // We are using the chatgpt-3.5-turbo model
    let requestBody = {
        model: "gpt-3.5-turbo",
        messages: [
            { role: "user", content: userInput },
        ],
    };

    let options = {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + apiKey,
        },
        payload: JSON.stringify(requestBody),
    };

    try {
        let response = await UrlFetchApp.fetch(apiUrl, options);
        console.log("API Response:", response.getContentText());

        let data = JSON.parse(response.getContentText());
        let chatResponseText = data.choices[0].message.content;
        console.log("chatResponseText:", chatResponseText);
        
    // Create custom headings for the prompt
    var textStyle = {};
    textStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    textStyle[DocumentApp.Attribute.BOLD] = true;
    textStyle[DocumentApp.Attribute.ITALIC] = false;
    textStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#212121";

    var boldPrompt = body.appendParagraph(userInput);
    boldPrompt.setAttributes(textStyle);
      // Append the server response paragraph and remove the bold font
      var responseParagraph = body.appendParagraph(chatResponseText.trimStart());
      textStyle[DocumentApp.Attribute.BOLD] = false;
      responseParagraph.setAttributes(textStyle);
    } catch (error) {
        console.log("Error:", error);
        return false;
    }
    return true;
}
