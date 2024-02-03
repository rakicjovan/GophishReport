let testP = document.getElementById("testP");
let ridRegex = /https:\/\/[A-Za-z]+\.[A-Za-z][A-Za-z]\/\?rid=[A-Za-z0-9]+/i;

Office.onReady((info) => {
    // Office is ready
    if (info.host === Office.HostType.Outlook) {
        //testP.appendChild(document.createElement("p").appendChild(document.createTextNode("This is a new paragraph!")));
        // Assign event handler to the button click
        document.getElementById("reportButton").onclick = parseMessage;
    }
});

// Event handler for the button click
function parseMessage() {
    //testP.appendChild(document.createElement("p").appendChild(document.createTextNode("This is a new paragraph!")));
    // Get the current item
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Parse the message content
                var messageContent = result.value;

                // Display the parsed message content in a new window
                //console.log(messageContent);
                console.log(typeof messageContent);

                if (ridRegex.test(messageContent)) {
                    testP.innerHTML = "This mail is reported!";
                    let ridUrl = messageContent.match(ridRegex)[0];
                    console.log(ridUrl);
                    let transformedURL = ridUrl.replace(/\?rid=\d+/, "report" + ridUrl.substring(ridUrl.indexOf('?')));
                    console.log(transformedURL);
                    webReport(transformedURL);

                } else {
                    testP.innerHTML = "This mail is not reported!";
                };

            } else {
                console.error("Error getting item body: " + result.error.message);
            }
        }
    );
}

async function webReport(reportUrl) {     
    try {
        const response = await fetch(reportUrl);

        if (!response.ok) {
            throw new Error(response.status);
        }
        else {
            testP.innerHTML = response.status;
        }
    }
    catch (error) {
        testP.innerHTML = "Error: " + error;
    }
};