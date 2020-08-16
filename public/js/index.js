let models = window["powerbi-client"].models;
let reportContainer = $("#report-container").get(0);

// Initialize iframe for embedding report
powerbi.bootstrap(reportContainer, {type: "report"});

var value1 = "";
var value2 = "";
var slicers = {
    Parameter1: 10,
    Parameter2: 10,
};

// AJAX request to get the report details from the API and pass it to the UI
$.ajax({
    type: "GET",
    url: "/getEmbedToken",
    dataType: "json",
    success: function (embedData) {

        // Create a config object with type of the object, Embed details and Token Type
        let reportLoadConfig = {
            type: "report",
            tokenType: models.TokenType.Embed,
            accessToken: embedData.accessToken,
            embedUrl: embedData.embedUrl,
            /*
            // Enable this setting to remove gray shoulders from embedded report
            settings: {
                background: models.BackgroundType.Transparent
            }
            */
        };

        // Use the token expiry to regenerate Embed token for seamless end user experience
        // Refer https://aka.ms/RefreshEmbedToken
        tokenExpiry = embedData.expiry;

        // Embed Power BI report when Access token and Embed URL are available
        let report = powerbi.embed(reportContainer, reportLoadConfig);

        // Clear any other loaded handler events
        report.off("loaded");

        // Triggers when a report schema is successfully loaded
        report.on("loaded", function () {
            console.log("Report load successful");
        });

        // Clear any other rendered handler events
        report.off("rendered");

        // Triggers when a report is successfully embedded in UI
        report.on("rendered", function () {
            console.log("Report render successful");
        });

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // report.refresh()
        //     .then(function (result) {
        //         Log.logText("Refreshed");
        //     })
        //     .catch(function (errors) {
        //         Log.log(errors);
        //     });

        report.on("dataSelected", function (event) {
            // Log.logText("Event - dataSelected:");
            var data = event.detail;
            console.log("data Selected event", data);
            console.log(data.dataPoints[0].identity[0].equals);
            console.log(data.dataPoints[0].identity[0].target.column);
            try {
                slicers[data.dataPoints[0].identity[0].target.column] = data.dataPoints[0].identity[0].equals;
            } catch (e) {

            }
            console.log(slicers);
        });

        // Report.on will add an event listener.
        report.on("buttonClicked", function (event) {
            var data = event.detail;
            console.log("btn event", data);
            console.log("btn id", event.detail['id']);

            if (event.detail['id'] === '749549557d5cb398db1c') {
                runAnimation(slicers.Parameter1, slicers.Parameter2);
            } else if (event.detail['id'] === 'd464a66bc3028d40e1d9') {
                stopAnimation();
            } else if (event.detail['id'] === 'ce2dffe249ace0000212') {
                console.log("Refresh Dataset");
            }

            report.getPages().then(function (pages) {
                // console.log(pages);
            });

        });

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // Clear any other error handler events
        report.off("error");

        // Handle embed errors
        report.on("error", function (event) {
            let errorMsg = event.detail;
            console.error(errorMsg);
            return;
        });
    },

    error: function (err) {

        // Show error container
        let errorContainer = $(".error-container");
        $(".embed-container").hide();
        errorContainer.show();

        // Get the error message from err object
        let errMsg = JSON.parse(err.responseText)['error'];

        // Split the message with \r\n delimiter to get the errors from the error message
        let errorLines = errMsg.split("\r\n");

        // Create error header
        let errHeader = document.createElement("p");
        let strong = document.createElement("strong");
        let node = document.createTextNode("Error Details:");

        // Get the error container
        let errContainer = errorContainer.get(0);

        // Add the error header in the container
        strong.appendChild(node);
        errHeader.appendChild(strong);
        errContainer.appendChild(errHeader);

        // Create <p> as per the length of the array and append them to the container
        errorLines.forEach(element => {
            let errorContent = document.createElement("p");
            let node = document.createTextNode(element);
            errorContent.appendChild(node);
            errContainer.appendChild(errorContent);
        });
    }
});