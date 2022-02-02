// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        //Debug
        console.log("Initialize was Successful");
    });
};

//Office.onReady(function () {
//    var UserName = Office.context.mailbox.userProfile.displayName;
//    //Debug
//    console.log("onReady was Successful");
//});

function addCUILabel(event) {
    //Debug
    console.log("Add CUI Label Function Called");
    //Using callback to run these functions in order
    setMarkings("CUI", function(){
         addTextToSubject(function(){
            addInternetHeader(function(){
                event.completed();
                console.log("Finished Marking Command");
            })
        });  
    });
}

function addPIILabel(event) {
    //Debug
    console.log("Add PII Label Function Called");
    //Using callback to run these functions in order
    setMarkings("PII", function(){
        addTextToSubject(function(){
           addInternetHeader(function(){
                event.completed();
                console.log("Finished Marking Command");
            });
               
        })
    });
}

// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
};

//Determine if the Message is HTML or Text Formatted
//Set the Header and Footer appropriately
function setMarkings(marking , callback) {
    Office.context.mailbox.item.body.getTypeAsync(
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                var dataValue = result.value; // Get selected data.
                console.log('Selected data is ' + dataValue);

                //Set newheader and newsig to either HTML or Text and CUI or PII
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    if (marking == "CUI") {
                        var newheader = "<p><b>CUI</b></p><br>"
                        var newsig = "<p><b>Controlled by: <br>Controlled by: <br>CUI Category: <br>Distribution/Dissemination Controls: <br>POC: </p><br><p>CUI</p></b>"
                    }
                    else if (marking == "PII") {
                        var newheader = "<p><b>CUI</b></p><br>"
                        var newsig = "<b><p>Controlled by: <br>Controlled by: <br>CUI Category: <br>Distribution/Dissemination Controls: <br>POC: </p><p><center>This e-mail contains Controlled Unclassified Information (CUI) information which must be protected under \
the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). Unauthorized disclosure \
or misuse of this PERSONAL INFORMATION may result in disciplinary action, criminal and/or civil penalties. Further \
distribution is prohibited without the approval of the author of this message unless the recipient has a need to know \
in the performance of official duties. If you have received this message in error, please notify the sender and delete all \
copies of this message.</center></p><p></p><br><p>CUI</p></b>"
                    }
                }
                else {
                    //Body is of text type.
                    if (marking == "CUI") {
                        var newheader = "CUI \n\n"
                        var newsig = "\n\n Controlled by: \n Controlled by: \n CUI Category: \n Distribution/Dissemination Controls: \n POC: \n\n CUI"
                    }
                    else if (marking == "PII") {
                        var newheader = "CUI \n\n"
                        var newsig = "\n\n Controlled by: \n Controlled by: \n CUI Category: \n Distribution/Dissemination Controls: \n POC: \n\n \
This e-mail contains Controlled Unclassified Information (CUI) information which must be protected under \
the Freedom of Information Act (5 U.S.C. 552) and/or the Privacy Act of 1974 (5 U.S.C. 552a). Unauthorized disclosure \
or misuse of this PERSONAL INFORMATION may result in disciplinary action, criminal and/or civil penalties. Further \
distribution is prohibited without the approval of the author of this message unless the recipient has a need to know \
in the performance of official duties. If you have received this message in error, please notify the sender and delete all \
copies of this message. \n\n CUI"
                    }
                };

                //Send the new header and footer to the Body
                addTextToBody(result.value, newheader, newsig);
            }
            //Error Handler
            else {
                //Debug
                console.log(result.error);
            }

        }
    );
    //Calls SetSubject Function Next
    callback();
};

// Determine if setSignatureAsync is supported in the version of Outlook
// If setSignatureAsync is not supported rewrite entire body of message
function addTextToBody(msgType, newheader, newsig) {

    //Debug
    console.log(msgType + " msgType was passed");

    //Get some User Data to setup new Signature
    var UserName = Office.context.mailbox.userProfile.displayName;
    var UserEmail = Office.context.mailbox.userProfile.emailAddress;

    //Setup the newSignature
    var str = "";
    str += "<table>";
    str += "<tr>";
    str += "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img alt='United_States_Department_of_Defense_Seal' width='100' height='100' src='data:image/png;base64,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'></td>";
    str += "<td style='padding-left: 5px;'>";
    str += "//SIGNED//<br/>";
    str += "<strong>" + UserName + "</strong>";
    str += "<br/>";
    str += "Title: <br/>";
    str += "Email: " + UserEmail + "<br/>";
    str += "Phone: <br/>";
    str += "</td>";
    str += "</tr>";
    str += "</table>";


    //Create new Signature for User
    var htmlSigHeader = str;
    var textSigHeader = "\n //Signed// \n " + UserName + " \n Email: " + UserEmail;


    if (msgType === Office.MailboxEnums.BodyType.Html) {
        newsig = htmlSigHeader + newsig;
        msgType = "html";
    }
    else {
        newsig = textSigHeader + newsig;
        msgType = "text"
    };

    //Debug
    console.log("This is right before setSignatureAsync is called");

    //SET SIGNATURE or Footer
    try {
        Office.context.mailbox.item.body.setSignatureAsync(
            newsig,
            {
                coercionType: msgType,
                asyncContext: "setSignature"
            },
            //Debug
            setCallback
        );

        //SET Header
        Office.context.mailbox.item.body.prependAsync(
            newheader,
            {
                coercionType: msgType,
                asyncContext: "prepend Body"
            },
            //Debug
            setCallback
        );
    }
    catch (err) {
        //If setSignatureAsync is not supported change the body of the message
        //or just add everything to the top and have user copy & paste
        
        console.log(err.message);

        //Set Body of Message instead of Header and Footer
        Office.context.mailbox.item.body.getAsync(
            msgType,
            function (oldBody) {
                Office.context.mailbox.item.body.setAsync(
                    newheader + oldBody.value + newsig,
                    {
                        coercionType: msgType,
                        asyncContext: "Body"
                    },
                    //Debug
                    setCallback
                )
            }
        );
    };
};


//statusUpdate(icon, marking + " Markings in text inserted successfully.");

// Gets the subject of the item and adds (U) in front of it.
function addTextToSubject(callback) {
    //var subject = Office.context.mailbox.item.subject;

    Office.context.mailbox.item.subject.getAsync(
        function (asyncResult) {
            //Debug
            console.log(asyncResult.value + " Subject was found");
            //Set the Subject
            Office.context.mailbox.item.subject.setAsync(
                "(U) " + asyncResult.value,
                { asyncContext: "Subject"},
                //Debug
                setCallback
            );
        }
    );
    //Calls addInternetHeader Function Next
    callback();
};

// Set custom internet headers.
function addInternetHeader(callback) {
    try{
        Office.context.mailbox.item.internetHeaders.setAsync(
        { "x-preferred-fruit": "orange"},
        setCallback(asyncResult, "Internet Header")
        );
        //Calls killEvent Next
        callback();
    }
    catch (err) {
        console.log("Internet Header =" + err.message);
        //Calls killEvent Next
        callback();
    }
    
};

//Debug Success/Error Handler
function setCallback() {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Successfully set Async on " + asyncResult.asyncContext);
    } else {
      console.log("Error setting " + asyncResult.asyncContext + ": " + JSON.stringify(asyncResult.error));
    }
};

