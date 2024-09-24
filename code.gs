var heading = CardService.newTextParagraph().setText("<b>Cybernut Reporting Tool</b>")
var alreadyCLickedHeading = CardService.newTextParagraph().setText("<b>WAIT - Did you accidentally click on something in this email?</b>")




// async function region(domainNameTo) {
//   try {
//     let res = UrlFetchApp.fetch(`https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com/userregion?domain=${domainNameTo}`, {
//       method: "get",
//       headers: { "content-Type": "application/json", }
//     });




//     const content = res.getContentText();
//     const jsonResponse = JSON.parse(content);
   
//     return jsonResponse.aws_region;
//   } catch (error) {
//     // If there's an error, return the default region
//     return 'us-east-1';
//   }
// }
async function region(domainNameTo) {
  try {
    let res = UrlFetchApp.fetch(`https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com/userregion?domain=${domainNameTo}`, {
      method: "get",
      headers: { "content-Type": "application/json" }
    });




    const statusCode = res.getResponseCode();
    const content = res.getContentText();
    const jsonResponse = JSON.parse(content);




    return {
      aws_region: jsonResponse.aws_region,
      status_code: statusCode
    };
  } catch (error) {
    return {
      aws_region: 'us-east-1',
      status_code: error.responseCode || 'unknown' // handle cases where error responseCode might not be available
    };
  }
}






















async function verifyDomain(fromDomain,messageid) {
 let reg = region(fromDomain)


if (reg === "ap-southeast-1") {
     
     var globalUrl = "vsqdkxcc8d"
    } else if (reg === "eu-central-1") {
     
     var globalUrl ="telmnzu55i"
    } else {
     
     var globalUrl = "44dgkpf1cb"
    }






console.log("message in conding ==============",encodeURIComponent(messageid),"from domain",fromDomain)


 








  let res = UrlFetchApp.fetch(
    `https://${globalUrl}.execute-api.us-east-1.amazonaws.com/admindomainsgoogle?domain=${fromDomain}&messageId=${encodeURIComponent(messageid)}`,
    {
      method: "get",
      headers: { "content-Type": "application/json" },
    }
  );
  // const content = res.getContentText();
  // const jsonResponse = JSON.parse(content);
  // return jsonResponse.exists;
  const content = res.getContentText();
  const jsonResponse = JSON.parse(content);
  return jsonResponse.messageExists
}
















let defaultMessageForThirdStep = 'Thank you, you will hear back from IT if you need to take any further action.'
let adminMessageForThirdStep = ""
















function extractDomainFromEmail(email) {
  var atIndex = email.indexOf("@");
  if (atIndex !== -1) {
    var domain = email.substring(atIndex + 1);
    domain = domain.replace(/[<>]/g, '');
    return domain;
  } else {
    return null; // Invalid email format
  }
}
















async function HomePage(e) {
  var reportButton = CardService.newTextButton().setText('Report Email')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setBackgroundColor("#D83025")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStep1"))
















  // Step 0 code
  var builder = CardService.newCardBuilder();
  builder.addSection(CardService.newCardSection()
    .setCollapsible(false)
    .setNumUncollapsibleWidgets(1)
    .addWidget(heading)
















    .addWidget(CardService.newTextParagraph()
      .setText('Suspicious content or sender? Report it for further analysis.'))
  )
  if (e) {
    builder.addSection(CardService.newCardSection()
      .addWidget(reportButton))
     
  }
  if (e.messageMetadata) {
  builder.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Onboarding Tutorial')
      .setDisabled(false)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("openLearnAddonLink"))));




  }
 
 




  var card = builder.build();
 
 
  return card;
}
















async function handleStep1(e) {
 
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken)
  var checkboxGroup = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName('selectedItems')
    .addItem("I replied to the email", "I replied to the email", false)
    .addItem('I downloaded a file', 'I downloaded a file', false)
    .addItem('I opened an attachment', 'I opened an attachment', false)
    .addItem('I visited a link', 'I visited a link', false)
    .addItem('I entered my password', 'I entered my password', false)
    .addItem('I forwarded the email', 'I forwarded the email', false)
    .addItem('I logged into a page', 'I logged into a page', false)
    .addItem('None of the above', 'None of the above', false);
































  var reportButton = CardService.newTextButton()
    .setText('Report Email')
    .setOnClickAction(CardService.newAction()
      .setFunctionName("handleStep2"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED).setBackgroundColor("#D83025");




    if (!e.messageMetadata.messageId) {
        // Create a card builder
        var cardBuilder = CardService.newCardBuilder();








        // Add a section with a text message
        var section = CardService.newCardSection();
        var textWidget = CardService.newTextParagraph()
            .setText('Please open the email and look for the button in the top left corner. Click on it to go back and find the report button.');








        section.addWidget(textWidget);
        cardBuilder.addSection(section);








        // Build and return the card
        var card = cardBuilder.build();
        return card;
    } else {
      var mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
  var sender = mailMessage.getFrom();
  var to = mailMessage.getTo();
  var timestamp = new Date();
  timestamp = timestamp.getTime()
    }




 
















  // To domain
  function getDomainFromEmail(email) {
    // Regular expression to match the domain part of the email
    const domainPattern = /@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/;
    const match = email.match(domainPattern);
    // Return the matched domain or null if no match
    return match ? match[1] : null;
  }
















  // let domainNameToIndexAtTheRate = to.indexOf("@");
  // let domainNameTo = to.slice(domainNameToIndexAtTheRate + 1);
  // domainNameTo = domainNameTo.replace(">", "");
  const domainNameTo = getDomainFromEmail(to);
 
console.log("domain to",domainNameTo)












  // from domain
  let domainNameFromSenderIndexAtTheRate = sender.indexOf("@");
  let domainNameFromSender = sender.slice(domainNameFromSenderIndexAtTheRate + 1);
  domainNameFromSender = sender.replace(">", "");
  var fromEmailAddress;
  if (e.messageMetadata) {  //if email is open set access token and get message id
    var shortMessageId = e.messageMetadata.messageId;
















    // Use the Gmail API to get the specific email's details
    var emailData = GmailApp.getMessageById(shortMessageId);
















    // Extract headers from the email data
    var headers = emailData.getRawContent().match(/^Message-ID: (.+)$/mi);
    // var headers = emailData.getRawContent().match(/^Message-ID:\s*<([^>]+)>/mi)
   
    var messageId = extractIdFromHeader(headers[1]);
   
    console.log("messageid",messageId)
   
















    if (sender.includes("<")) {
      var regex = /<([^>]+)>/;
      var match = regex.exec(sender);
      if (match && match.length > 1) {
        fromEmailAddress = match[1];
      }
    }
    else {
      fromEmailAddress = domainNameFromSender
    }
    var fromDomain = extractDomainFromEmail(fromEmailAddress)












  const awsRegion = await region(domainNameTo)
    const reg = awsRegion.aws_region
    console.log("this is region",reg,"domain",domainNameTo)
 








    try {
      // STEP 1 CODE
      const isVerifiedDomain = await verifyDomain(fromDomain,messageId)
      console.log("this is region",reg,"domain",domainNameTo,"message id",messageId,"verifay domain",isVerifiedDomain,"from domain",fromDomain)
     
      if (isVerifiedDomain) {
        // try {
        // await senDataToDb(to, fromEmailAddress, messageId, timestamp, fromDomain, isVerifiedDomain)
        // } catch (error) {
        // Logger.log(error)
        // }
        // Do the redirect
        var encodedMessageId = encodeURIComponent(messageId);
       
        var redirectUrl = `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${reg ? reg : "us-east-1"}`;
        return CardService.newActionResponseBuilder()
          .setOpenLink(CardService.newOpenLink()
            .setUrl(redirectUrl))
          .build();
      }
      else {




         const thread = GmailApp.getMessageById(e.messageMetadata.messageId).getThread();
         const labels = thread.isInInbox()
       
        if (labels) {
            var builder = CardService.newCardBuilder()
        builder.addSection(CardService.newCardSection()
          .setCollapsible(false)
          .setNumUncollapsibleWidgets(1)
          .setNumUncollapsibleWidgets(1)
          .addWidget(alreadyCLickedHeading)
          .addWidget(CardService.newTextParagraph()
            .setText(
              '<b>You will not get in trouble by telling us.</b><br/><br/>By sharing this information, it will help your IT department monitor and catch potential cyber attacks in your school district.<br/><br/>'
            ))
          .addWidget(CardService.newTextParagraph().setText("Thank you for your cooperation and transparency.<br/><br/><b>Please select from the list below if applicable:</b> "))
          .addWidget(checkboxGroup)
          .addWidget(reportButton)
        )
















        builder.setFixedFooter(CardService.newFixedFooter()
          .setPrimaryButton(CardService.newTextButton()
            .setText('Onboarding Tutorial')
            .setDisabled(false)
            .setOnClickAction(CardService.newAction()
              .setFunctionName("openLearnAddonLink"))));
















        var card = builder.build();
        return card;
        }
        else {
         
         const res_value = await handleStep2(e)
         
         return res_value ;




         
         
        }
       
      }
    }
    catch (error) {
      Logger.log(error)
    }
  }
}




async function handleStep2(e) {




var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);




  var fromDomain = e.parameters.fromDomain;
  var selectedItemsValues = e.formInputs.selectedItems; // Get the selected items from the form input
  var selectedItems = [];
 
  // Loop the selected items and add to array
  if (selectedItemsValues) {
    for (var i = 0; i < selectedItemsValues.length; i++) {
      selectedItems.push(selectedItemsValues[i]);
    }
  }




  var messageId = e.messageMetadata.messageId;
  var mailMessage = GmailApp.getMessageById(messageId);
  var subject = mailMessage.getSubject();
  var sender = mailMessage.getFrom();
  var body = mailMessage.getPlainBody();
  var bodyHtml = mailMessage.getBody();
  const checkedValues = selectedItems.join(', ');
  var editedBody = checkedValues;
  var to = mailMessage.getTo();




function extractFirstEmail(to) {
  // Regular expression to match email addresses
  const emailPattern = /[\w.-]+@[\w.-]+\.\w+/;
  let match = emailPattern.exec(to);
 
  // Return the first matched email address
  return match ? match[0] : null;
}




 
 
 
  // let domainNameToIndexAtTheRate = emailAddressTo.indexOf("@");
  // let domainNameTo = emailAddressTo.slice(domainNameToIndexAtTheRate + 1).replace(/[^a-zA-Z0-9.-]/g, "");
  function extractFirstDomain(to) {
  // Split the input by commas to handle multiple email addresses
  var emails = to.split(',');




  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    var emailAddress = email.match(/<(.+)>/) ? email.match(/<(.+)>/)[1] : email.trim();
    var domainNameIndex = emailAddress.indexOf("@");
    if (domainNameIndex !== -1) {
      var domainName = emailAddress.slice(domainNameIndex + 1).replace(/[^a-zA-Z0-9.-]/g, "");
      if (domainName) {
        return domainName;
      }
    }
  }
  return '';
}
  let domainNameTo = extractFirstDomain(to)
  // Extract domain from 'sender' field
  let emailAddressSender = sender.match(/<(.+)>/) ? sender.match(/<(.+)>/)[1] : sender;
  let domainNameFromSenderIndexAtTheRate = emailAddressSender.indexOf("@");
  let fromEmailAddress = emailAddressSender.slice(domainNameFromSenderIndexAtTheRate + 1).replace(/[^a-zA-Z0-9.-]/g, "");
  let domainNameFromSender = fromEmailAddress




  // Extract headers from the email data
  var headers = mailMessage.getRawContent().match(/^Message-ID: (.+)$/mi);
  var messageIdOrg = headers ? extractIdFromHeader(headers[1]) : null;




  // Further processing with extracted domains and headers
 
 
  fromEmailAddress
 
console.log("this is domain",domainNameTo)








const awsRegion = await region(domainNameTo)
  const reg = awsRegion.aws_region




console.log("region response code",reg, domainNameTo)
// Extract domain from 'to' field
 
  console.log("code",awsRegion.status_code,"email",Session.getActiveUser().getEmail())
  if (awsRegion.status_code === 200){
  var emailAddressTo = extractFirstEmail(to)
  }else {
   
   
    var emailAddressTo = Session.getActiveUser().getEmail()
  }
 












  try {
    let adminUrl;
    let serviceUrl;
    if (reg === "ap-southeast-1") {
      adminUrl = "b4nzi83qm2";
      serviceUrl = "rmlq7e7vh7"
    } else if (reg === "eu-central-1") {
      adminUrl = "dej7cfclm9";
      serviceUrl = "skzb4w2nje"
    } else {
      adminUrl = "k3g591je54";
      serviceUrl = "560ef3pt4j"
    }








// function getDomainFromActiveUser() {
//   // Get the active user's email
//   var email = Session.getActiveUser().getEmail();
 
//   // Extract the domain from the email
//   var domain = email.split('@')[1];
 
//   return domain;
// }
// const currentDomain = getDomainFromActiveUser()
















//     console.log("this is domian",domainNameTo,"current domain",currentDomain)
//     let suspeciousEmail = UrlFetchApp.fetch(`https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`, {
//       method: "post",
//       headers: {
//         "content-Type": "application/json",
//       },
//       payload: JSON.stringify({
//         domain: domainNameTo
//       })
//     });




function getDomainOrFallback(domainNameTo, adminUrl, reg) {
  try {
    // Try to fetch the domain from the API
    let response = UrlFetchApp.fetch(`https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`, {
      method: "post",
      headers: {
        "content-Type": "application/json",
      },
      payload: JSON.stringify({
        domain: domainNameTo
      })
    });
    return response
  } catch (error) {
    // If there's an error, construct a fallback response
    var email = Session.getActiveUser().getEmail();
    var currentDomain = email.split('@')[1];
    console.log("current domian",currentDomain,adminUrl,reg)
    let response = UrlFetchApp.fetch(`https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`, {
      method: "post",
      headers: {
        "content-Type": "application/json",
      },
      payload: JSON.stringify({
        domain: currentDomain
      })
    });
    return response
  }
}
    suspeciousEmail = JSON.parse(getDomainOrFallback(domainNameTo, adminUrl, reg))
    console.log("forward suspeciousemail",suspeciousEmail,"hello")
   
   
    adminMessageForThirdStep = suspeciousEmail.CONFIRMATION_MESSAGE












console.log("hello ",suspeciousEmail.FORWARD_SUSPICIOUS_EMAIL,fromEmailAddress,emailAddressTo,domainNameTo,messageIdOrg,subject,editedBody)


    // https://pyslc6a88h.execute-api.us-east-1.amazonaws.com/eventdispatcher
    let resss = UrlFetchApp.fetch(`https://${serviceUrl}.execute-api.${reg}.amazonaws.com/eventdispatcher`, {
      method: "post",
      headers: {
        "content-Type": "application/json",
      },
      payload: JSON.stringify({
        domain: domainNameTo,
        fromAddress: fromEmailAddress,
        destination: emailAddressTo,
        action: "FORWARD_SUSPICIOUS_EMAIL",
        message_id: messageIdOrg,
        emailtemplate: bodyHtml,
        provider: "google",
        triggerBoth: true,
        email: suspeciousEmail.FORWARD_SUSPICIOUS_EMAIL,
        subject: subject,
        body: editedBody,
        source: "gmail"
      })
    });




 




 
  } catch (error) {
    Logger.log(error)
  }
  var builder = CardService.newCardBuilder()
  builder.addSection(CardService.newCardSection()
    .setCollapsible(false)
    .setNumUncollapsibleWidgets(1)
    .addWidget(heading)
    .addWidget(CardService.newTextParagraph()
      .setText(adminMessageForThirdStep.length > 0 ? adminMessageForThirdStep : defaultMessageForThirdStep))
  )
 
  const isVerifiedDomain = await verifyDomain(domainNameFromSender, reg,messageId)
  console.log("is varified",isVerifiedDomain,"messageid",messageId)




  const threads = GmailApp.getMessageById(e.messageMetadata.messageId).getThread();
  const check_Inbox = threads.isInInbox()
  if (check_Inbox) {
  builder.setFixedFooter(CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Onboarding Tutorial')
      .setDisabled(false)
      .setOnClickAction(CardService.newAction()
        .setFunctionName("openLearnAddonLink"))));
  }
 
   




 








  var card = builder.build();
 




  return card;
}
















function redirect() {
  try {
    let redirectApi = UrlFetchApp.fetch("https://pyslc6a88h.execute-api.us-east-1.amazonaws.com/email", {
      method: "post",
      headers: {
        "content-Type": "application/json",
      },
      payload: JSON.stringify({
        "source": "linked"
      })
    })
  }
  catch (error) {
    Logger.log(error)
  }
}


function generateUUID() {
  var template = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
  return template.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    var v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}
async function openLearnAddonLink() {
  let email = Session.getActiveUser().getEmail();
  let currentDomain = email.split('@')[1];
  var reg = await region(currentDomain)
  console.log("region",reg.aws_region,"current domain",currentDomain)
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(`https://www.cybernut-k12.com/onboardingreport?partitionkey=campaign-8d16cb87-e16e-400a-a288-14e55a99a1bb&sortkey=${generateUUID()}&region=${reg.aws_region}&email=${email}&tracker=demo`))
   
    .build();
}








function extractIdFromHeader(header) {
  var matches = header.match(/<([^>]+)@/);
  if (matches && matches.length > 1) {
    return matches[1];
  }
}
















