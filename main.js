let alternateRid = "keyname";
let headP = document.getElementById("headP");
let reportButton = document.getElementById("reportButton");
let checkmarkHTML = '<svg class="checkmark" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 52 52"><circle class="checkmark__circle" cx="26" cy="26" r="25" fill="none"/><path class="checkmark__check" fill="none" d="M14.1 27.2l7.1 7.2 16.7-16.8"/></svg>';



let ridRegex;
if (alternateRid !== "") {
    ridRegex = new RegExp(`https:\/\/.*${alternateRid}=.{7}`);
} else {
    ridRegex = /https:\/\/.*rid=.{7}/;
}

Office.onReady((info) => {
    // Office is ready
    if (info.host === Office.HostType.Outlook) {
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
                let messageContent = result.value;
                // Check if the message contains a report link
                if (ridRegex.test(messageContent)) {
                    headP.innerHTML = "This mail is reported!";
                    let indexOfEqualSign = messageContent.match(ridRegex)[0].indexOf('=');
                    let ridUrl = messageContent.match(ridRegex)[0].substr(0, indexOfEqualSign + 8);
                    let transformedURL = addReportPrefix(ridUrl);
                    webReport(transformedURL);

                }
                // If the message does not contain a report link, still display a checkmark but no further action is taken
                else {
                    headP.innerHTML = "[Dev no rid] Successfully reported the mail, you can delete it now!";
                    document.getElementById("reportButton").disabled = true;
                    document.getElementById("wrapperId").insertAdjacentHTML("beforeend", checkmarkHTML);
                    console.log(document.getElementById("wrapperId").innerHTML);
                };

            }
            else {
                console.error("Error getting item body: " + result.error.message);
            }
        }
    );
}

function addReportPrefix(url) {
    // Check if the URL contains a query string
    if (url.includes('?')) {
        // Split the URL into two parts: the base URL and the query string
        const [baseUrl, queryString] = url.split('?');

        // Add "report" in front of the query string
        const newUrl = `${baseUrl}/report?${queryString}`;

        return newUrl;
    }

    // If there is no query string, simply add "/report" to the end of the URL
    return `${url}/report`;
};

async function webReport(reportUrl) {
    // Fetch the report URL     
    try {
        const response = await fetch(reportUrl);

        if (!response.ok && response.status !== 204) {
            console.log("test");
            document.getElementById("reportButton").disabled = true;
            throw new Error(response.status);
        }
        else {
            headP.innerHTML = "Successfully reported the mail, you can delete it now!";
            document.getElementById("reportButton").disabled = true;
            document.getElementById("wrapperId").insertAdjacentHTML("beforeend", checkmarkHTML);
            console.log(document.getElementById("wrapperId").innerHTML);
        }
    }
    catch (error) {
        console.log("test");
        headP.innerHTML = "Error: " + error;
        document.getElementById("reportButton").disabled = true;

        document.getElementById("wrapperId").insertAdjacentHTML("beforeend", checkmarkHTML);
        console.log(document.getElementById("wrapperId").innerHTML);
    }
};

function showInfoBox() {
    let infoModal = document.getElementById("infoModal");
    infoModal.style.display = "block";
    setTimeout(function() {
        infoModal.style.opacity = 1;
    }, 10);
};

function closeInfoBox() {
    document.getElementById("infoModal").style.display = "none";
};

function openGitHub() {
    window.open("https://github.com/jole583/Goreport", "_blank");
};

function addCheckmark() {

};
