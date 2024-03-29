$(document).ready(function () { 
    ExecuteOrDelayUntilScriptLoaded(AttachClickEventToButton, 'SP.js'); 
});
 
//This method attaches click event to input button. 
function AttachClickEventToButton() { 
    try { 
        $("#btnTriggerMicrosoftFlow").click(function () { 
            StartMicrosoftFlowTriggerOperations(); 
        });
     }
     catch (e) { 
        console.log("Error occurred in AttachClickEventToButton " + e.message); 
    } 
}
 
//This method triggers the microsoft flow 
function StartMicrosoftFlowTriggerOperations() { 
    try { 
        
		//var dataTemplate = "{\r\n    \"emailaddress\":\"{0}\",\r\n    \"emailSubject\": \"{1}\",\r\n    \"emailBody\": \"{2}\"\r\n}"
		var dataTemplate = "{\r\n    \"PersonID\":\"{0}\",\r\n    \"FirstName\": \"{1}\",\r\n    \"LastName\": \"{2}\",\r\n  \"JobTitle\": \"{3}\",\r\n \"EmailAddress\": \"{4}\"\r\n}"; 
        var httpPostUrl = "HOSTURLFROMMICROSOFTFLOW"; // Copy Url from microsoft flow.
        //Call FormatRow function and replace with the values supplied in input controls. 
        dataTemplate = dataTemplate.FormatRow($("#txtPersonID").val(), $("#txtFirstName").val(), $("#txtLastName").val(),$("#txtJobTitle").val(),$("#txtEmailAddress").val());
         var settings = { 
            "async": true, 
            "crossDomain": true, 
            "url": httpPostUrl, 
            "method": "POST", 
            "headers": { 
                "content-type": "application/json", 
                "cache-control": "no-cache" 
            }, 
            "processData": false, 
            "data": dataTemplate 
        } 
        $.ajax(settings).done(function (response) { 
            console.log("Successfully triggered the Microsoft Flow. "); 
			$("#txtPersonID").val("");
			$("#txtFirstName").val("");
			$("#txtLastName").val("");
			$("#txtJobTitle").val("");
			$("#txtEmailAddress").val("");
        }); 
    } 
    catch (e) { 
        console.log("Error occurred in StartMicrosoftFlowTriggerOperations " + e.message); 
    } 
}
 
//This method formats the rowTemplate by replacing the placeholders based on the arguments passed. 
String.prototype.FormatRow = function () { 
    try { 
        var content = this; 
        for (var i = 0; i < arguments.length; i++) { 
            var replacement = '{' + i + '}'; 
            content = content.replace(replacement, arguments[i]); 
        } 
        return content; 
    } 
    catch (e) { 
        console.log("Error occurred in FormatRow " + e.message); 
    }
 }
