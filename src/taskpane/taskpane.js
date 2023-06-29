/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {

  // Thid is customer API endpoint. Not tested due to CORS
  async function makeAPICall(emailBody) {
    const apiUrl = 'https://email-parsing.api.dev.dfp.ai/suggestion';
    const requestData = { message: emailBody };
  
    fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestData),
    })
      .then(response => response.json())
      .then(data => {
        console.log(data); // Logging the response data to console
      })
      .catch(error => {
        console.error(error); // Logging the error to console
      });
  }
  
  // These two functions below is for debug and replaces customer API backend
  
  // This first function is for plain text OpenAI API caller. Backend CORS free
  async function runWebcall(emailBody) {
    try {
      const encodedPrompt = encodeURIComponent(emailBody);
      const url = `https://openaicaller1.azurewebsites.net/api/openaicaller1?prompt=${encodedPrompt}`;
  
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
  
      const data = await response.text();
      return data;
    } catch (error) {
      console.error('An error occurred while making the web call:', error);
      throw error;
    }
  }

  // This second function is for JSON OpenAI API caller. Backend CORS free.
  async function makeAzureAPICall(emailBody) {
    const apiUrl = 'https://openaicaller1.azurewebsites.net/api/json_openapicaller2?clientId=apim-openaicaller1-apim';
    const requestData = { prompt: emailBody };
  
    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestData),
      });
  
      const data = await response.text(); // Convert the response to plain text
  
      return data; // Return the response as plain text
    } catch (error) {
      console.error(error); // Logging the error to console
      throw error; // Rethrow the error to be handled by the caller
    }
  }
  
  // We get message body and push it to the call function

  Office.context.mailbox.item.body.getAsync("text", async function(result) {
    if (result.status === "succeeded") {
      const emailBody = result.value;
      const updatedBody = await makeAzureAPICall(emailBody)
      // Here we trigger new reply form and paste new body
      Office.context.mailbox.item.displayReplyAllForm({
        'htmlBody': updatedBody
      });
    } else {
      console.error("Error retrieving email body: " + result.error);
    }
  });
}