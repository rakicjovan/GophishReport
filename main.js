const ridAlternative = "keyname";

let headP = document.getElementById("headP");
let ridRegex = /https:\/\/[A-Za-z]+\.[A-Za-z][A-Za-z]\/\?rid=[A-Za-z0-9]+/i;
let reportButton = document.getElementById("reportButton");

if (ridAlternative) {
    ridRegex = new RegExp('https://[A-Za-z]+\\.[A-Za-z][A-Za-z]/\\?' + ridAlternative + '=[A-Za-z0-9]+', 'i');
    console.log(ridRegex);
}

Office.onReady((info) => {
    // Office is ready
    if (info.host === Office.HostType.Outlook) {
        //headP.appendChild(document.createElement("p").appendChild(document.createTextNode("This is a new paragraph!")));
        // Assign event handler to the button click
        document.getElementById("reportButton").onclick = parseMessage;
    }
});

// Event handler for the button click
function parseMessage() {
    // Get the current item
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Parse the message content
                var messageContent = result.value;

                if (ridRegex.test(messageContent)) {
                    headP.innerHTML = "This mail is reported!";
                    let ridUrl = messageContent.match(ridRegex)[0];
                    console.log(ridUrl);
                    let transformedURL = ridUrl.replace(/\?rid=\d+/, "report" + ridUrl.substring(ridUrl.indexOf('?')));
                    console.log(transformedURL);
                    webReport(transformedURL);

                } else {
                    headP.innerHTML = "This mail is not reported!";
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
            console.log("test");
            reportButton.disabled = true;
            throw new Error(response.status);
        }
        else {
            headP.innerHTML = "Successfully reported the mail, you can delete it now!" + response.status;
            reportButton.disabled = true;
        }
    }
    catch (error) {
        console.log("test");
        headP.innerHTML = "Error: " + error;
        document.getElementById("reportButton").disabled = true;
    }
};