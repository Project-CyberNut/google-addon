var version = "v 2.3.8"
var heading = CardService.newTextParagraph().setText(
  `<b>Cybernut Reporting Tool   </b>  ${version}`
);
var alreadyClickedHeading = CardService.newTextParagraph().setText(
  "<b>WAIT - Did you accidentally click on something in this email?</b>"
);
async function callErrorReportingApi(error, htmlbody) {
  var now = new Date();
  console.log(`Add on version Cybernut Reporting Tool  ${version}`)




  // console.log("event time ",now.toLocaleString(),"html body",htmlbody)
  try {
    const url = `https://560ef3pt4j.execute-api.us-east-1.amazonaws.com/microsoftaddinactivitynew?timestamp=${now.toLocaleString()}`;
    const payload = {
      id: Session.getActiveUser().getEmail(),
      body: String(error) + ` Add-On Version: ${version}`,
      htmlbody: htmlbody || "",

    };
    const options = {
      method: "post",
      headers: { "content-Type": "application/json" },
      payload: JSON.stringify(payload),
    };
    let res = UrlFetchApp.fetch(url, options);
    console.log("Error reporting API called successfully.", res);
  } catch (apiError) {
    Logger.log("Failed to call error reporting API: " + apiError.message);
    return new Error(apiError.message);
  }
}




async function region(domainNameTo) {
  try {
    const res = UrlFetchApp.fetch(
      `https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com/userregion?domain=${domainNameTo}`,
      {
        method: "get",
        headers: { "Content-Type": "application/json" },
        muteHttpExceptions: false, // Default behavior (throws on non-2xx)
      }
    );




    // Explicit status code check for 200
    const statusCode = res.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`API request failed with status ${statusCode}`);
    }




    const content = res.getContentText();
    const jsonResponse = JSON.parse(content);




    return {
      aws_region: jsonResponse.aws_region,
      status_code: statusCode,
    };
  } catch (error) {
    await callErrorReportingApi(error, " ");
    // Extract status code from error message if available
    const errorStatusCode =
      error.message.match(/status (\d+)/)?.[1] ||
      error.responseCode ||
      "unknown";




    return {
      aws_region: "us-east-1",
      status_code: errorStatusCode,
    };
  }
}

function foundReportUrl(e) {
  const message = GmailApp.getMessageById(e.gmail.messageId);
  const emailBody = message.getBody();

  // The encoded version of "https://www.cybernut-k12.com/report"
  // We only need a key part of it to find the link.
  const encodedTarget = 'www.cybernut-k12.com';
  console.log(encodedTarget, "this is the cybernut url to find", emailBody, "this is email body")

  // The String.includes() method is the simplest way to find this text.
  if (emailBody.includes(encodedTarget)) {
    return true; // Found the encoded link.
  }

  return false; // Did not find it.
}



function cybernutDomains(senderDomain) {
  if (!senderDomain) {
    return false;
  }

  // Convert to lowercase for case-insensitive comparison
  const domain = senderDomain.toLowerCase();

  // List of suspicious domains
  const suspiciousDomains = [
    'k12districtnotification.com',
    'google-notice-alert.com',
    'google-k12-support.com',
    'google-k12-alert.com',
    'seesaw-services.com',
    'zoom-securelogin.com',
    'google-signin-support.com',
    'adobe-platform.com',
    'google-signinsupport.com',
    'ixl-services.com',
    'google-k12.com',
    'nearpodportal.com',
    'google-report.com',
    'brainpop-mail.com',
    'download-onedrive.com',
    'account-teams.com',
    'kamiportal.com',
    'outlook365-office.com',
    'learninga-z-updates.com',
    'amaz0n-securelogin.com',
    'dropboxsecure-login.com',
    'amazingdeals-shopping.com',
    'amazon-ordertrack.com',
    'khanacademy-app.com',
    'box-fileaccess.com',
    'naviance-mail.com',
    'box-securelogin.com',
    'schooltool-services.com',
    'connect-meetingshub.com',
    'classlink-online.com',
    'connect-socials.com',
    'frontlinealerts.com',
    'dhl-expressship.com',
    'newslea-updates.com',
    'dhl-globaltrack.com',
    'quizizz-services.com',
    'docu-signin.com',
    'docus1gn-verify.com',
    'canvas-portal.com',
    'drop-boxlogin.com',
    'nvisionconnect.com',
    'powerschool-updates.com',
    'fastfilesharenow.com',
    'fastshipping-hub.com',
    'fed-extracking.com',
    'upsdelivery-hub.com',
    'fedex-deliverynow.com',
    'gimkit-services.com',
    'g00gle-verify.com',
    'discoveryedu-platform.com',
    'global-shiptrack.com',
    'kahoot-connect.com',
    'media-socialhub.com',
    'deltamath-services.com',
    'micosoft-secure.com',
    'edpuzzle-online.com',
    'ms-updatecenter.com',
    'castlelearning-platform.com',
    'secure-meetingportal.com',
    'sharefiles-secure.com',
    'malwarebytes-notifications.com',
    'shopnow-discounts.com',
    'sophos-connect.com',
    'ups-packagealert.com',
    'usps-securepackage.com',
    'usps-trackingportal.com',
    'shopify-updates.com',
    'zoom-meetingconnect.com',
    'walmart-communications.com',
    'meta-alerts.com',
    'google-notice.com',
    'instagram-services.com',
    'cybernut-k12.com',
    'tiktok-teams.com',
    'cybernut.com',
    'hulu-connect.com',
    'netflix-updates.com',
    'schoology-communications.com'
  ];

  return suspiciousDomains.includes(domain);
}

function getAttachmentIds(messageId) {
  const attachmentIds = [];
  try {
    // 1. Get the message using the standard service
    const message = GmailApp.getMessageById(messageId);


    // 2. Get all attachments from the message
    const attachments = message.getAttachments();


    // 3. Process each attachment
    attachments.forEach(attachment => {
      attachmentIds.push({
        filename: attachment.getName(),
        mimeType: attachment.getContentType(),
        // 4. Get the file content and encode it in Base64
        // content_base64: Utilities.base64Encode(attachment.getBytes())
      });
    });


  } catch (e) {
    console.log('Error fetching attachments with GmailApp for messageId %s: %s', messageId, e.toString());
  }
  return attachmentIds;
}


async function verifyDomain(sourceid, messageid, region, activeuser, moveToTrash) {
  try {
    const globalUrl = getGlobalUrl(region); // Assuming this is a helper function you have
    const apiUrl = `https://${globalUrl}.execute-api.${region}.amazonaws.com/admindomainsgoogle?gmailId=${sourceid}&user_email=${activeuser}&messageId=${encodeURIComponent(
      messageid
    )}`;


    console.log(
      `Verifying with Gmail ID: ${sourceid}`,
      `Message ID: ${messageid}`,
      `API URL: ${apiUrl}`
    );


    // Directly attempt to fetch the data once
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "GET",
      headers: { "Content-Type": "application/json" },
      muteHttpExceptions: false, // Throw an error on non-2xx responses
    });


    // Explicitly check for a successful status code
    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      throw new Error(`API returned status ${statusCode}`);
    }


    const jsonResponse = JSON.parse(response.getContentText());
    if (moveToTrash === true) {
      return jsonResponse.trashEmail
    } else {
      return jsonResponse.messageExists
    }





  } catch (error) {
    // If the single attempt fails, report the error and stop
    console.error(`Domain verification failed: ${error.message}`);
    await callErrorReportingApi(error, " "); // Your custom error reporting
    throw new Error(`Domain verification failed: ${error.message}`);
  }
}




// Helper function to get the global URL based on region
function getGlobalUrl(region) {
  let mapping = {
    "ap-southeast-1": "vsqdkxcc8d",
    "eu-central-1": "telmnzu55i",
  };
  return mapping[region] || "44dgkpf1cb"; // Default URL
}




async function EventDispatcherApi(payload, serviceUrl, reg) {
  const url = `https://${serviceUrl}.execute-api.${reg}.amazonaws.com/eventdispatcher`;




  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };




  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();




  if (code === 200) {
    return response.getContentText();
  } else {
    return `Error: Received HTTP ${code} - ${response.getContentText()}`;
  }
}




async function getDomainOrFallback(domainNameTo, adminUrl, reg) {
  const url = `https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`;




  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: { "Content-Type": "application/json" },
    payload: JSON.stringify({ domain: domainNameTo }),
    muteHttpExceptions: true,
  });




  if (response.getResponseCode() === 200) {
    return JSON.parse(response);
  }




  return new Error(
    `API failed. Status: ${response.getResponseCode()} - ${response.getContentText()}`
  );
}




let defaultMessageForThirdStep =
  "Thank you, you will hear back from IT if you need to take any further action.";
let adminMessageForThirdStep = "";

function buildErrorCard() {
  var cardBuilder = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  var textWidget = CardService.newTextParagraph().setText(
    "There was an error in completing your action, for escalation / faster resolution you can contact us at support@cybertnut.com"
  );

  // Add a close button that navigates back to home
  var closeButton = CardService.newTextButton()
    .setText("Close")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#4285F4")
    .setOnClickAction(CardService.newAction().setFunctionName("HomePage"));

  section.addWidget(textWidget);
  section.addWidget(closeButton);
  cardBuilder.addSection(section);
  var card = cardBuilder.build();
  return card;
}




async function HomePage(e) {
  console.log('HomePage|called!');
  try {
    var reportButton = CardService.newTextButton()
      .setText("Report Email")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#D83025")
      .setOnClickAction(CardService.newAction().setFunctionName("handleStep1"));
    console.log('HomePage|reportButton|created!');



    var builder = CardService.newCardBuilder();
    console.log('HomePage|CardService|builder|section|created!');
    builder.addSection(
      CardService.newCardSection()
        .setCollapsible(false)
        .setNumUncollapsibleWidgets(1)
        .addWidget(heading)
        .addWidget(
          CardService.newTextParagraph().setText(
            "Suspicious content or sender? Report it for further analysis."
          )
        )
    );
    console.log('HomePage|CardService|builder|section|added!');
    if (e.gmail) {
      console.log('HomePage|CardService|builder|section|e.gmail|true!');
      let mailMessage = GmailApp.getMessageById(e.gmail.messageId);
      console.log('HomePage|CardService|builder|section|mailMessage|found!');
      let bodyHtml = mailMessage.getBody();
      console.log('HomePage|CardService|builder|section|bodyHtml|found!');
      // console.log("html body",bodyHtml)
      const errorReportingApiResponse = await callErrorReportingApi("Home function run perfectly", bodyHtml);
      console.log('HomePage|CardService|builder|section|Inside If|errorReportingApiResponse:', errorReportingApiResponse);
    } else {
      const errorReportingApiResponse = await callErrorReportingApi(
        "Home function run perfectly in inbox folder",
        "none"
      );
      console.log('HomePage|CardService|builder|section|Inside Else|errorReportingApiResponse:', errorReportingApiResponse);
    }

    if (e) {
      builder.addSection(CardService.newCardSection().addWidget(reportButton));
      console.log('HomePage|CardService|builder|section|reportButton|added!');
    }




    builder.setFixedFooter(
      CardService.newFixedFooter().setPrimaryButton(
        CardService.newTextButton()
          .setText("Onboarding Tutorial")
          .setDisabled(false)
          .setOnClickAction(
            CardService.newAction().setFunctionName("openLearnAddonLink")
          )
      )
    );
    console.log('HomePage|CardService|builder|section|fixedFooter|added!');

    var card = builder.build();
    console.log('HomePage|CardService|builder|section|card|built!');
    return card;
  } catch (error) {
    console.log('HomePage|error|caught!', error);
    await callErrorReportingApi(error.stack, " ");
    return buildErrorCard();
  }
}

async function handleStep1(e) {
  console.log('handleStep1|called!');
  let bodyHtml = "";
  if (e?.messageMetadata?.messageId) {
    console.log('handleStep1|messageMetadata.messageId|found!', e.messageMetadata.messageId);
    let mail = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mail ? mail.getBody() : " ";
    console.log('handleStep1|mail|found!');
    console.log('handleStep1|bodyHtml|found!');
  }
  console.log('handleStep1|bodyHtml|found!', {bodyHtml});
  try {
    var checkboxGroup = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setFieldName("selectedItems")
      .addItem("I replied to the email", "I replied to the email", false)
      .addItem("I downloaded a file", "I downloaded a file", false)
      .addItem("I opened an attachment", "I opened an attachment", false)
      .addItem("I visited a link", "I visited a link", false)
      .addItem("I entered my password", "I entered my password", false)
      .addItem("I forwarded the email", "I forwarded the email", false)
      .addItem("I logged into a page", "I logged into a page", false)
      .addItem("None of the above", "None of the above", false);

    console.log('handleStep1|checkboxGroup|created!');
    var reportButton = CardService.newTextButton()
      .setText("Report Email")
      .setOnClickAction(CardService.newAction().setFunctionName("handleStep2"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#D83025");
    console.log('handleStep1|reportButton|created!');

    if (!e?.messageMetadata?.messageId) {
      console.log('handleStep1|Inside If messageId: false');
      var cardBuilder = CardService.newCardBuilder();
      console.log('handleStep1|cardBuilder|created!');
      var section = CardService.newCardSection();
      var textWidget = CardService.newTextParagraph().setText(
        "Please open the email and look for the button in the top left corner. Click on it to go back and find the report button."
      );

      console.log('handleStep1|textWidget|created!');


      section.addWidget(textWidget);
      cardBuilder.addSection(section);
      console.log('handleStep1|section|added!');

      var card = cardBuilder.build();
      console.log('handleStep1|card|built!');
      return card;
    }
    else {
      var mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
      console.log('handleStep1|mailMessage|found!');
      var sender = mailMessage.getFrom();
      console.log('handleStep1|sender|found!', {sender});
      var to = Session.getActiveUser().getEmail();
      console.log('handleStep1|to|found!', {to});
      var timestamp = new Date();
      timestamp = timestamp.getTime();
    }
    const domainNameTo = to.split("@")[1];
    console.log("domain to", domainNameTo, "sender", sender.split("@")[1].replace('>', ''));

    if (e.messageMetadata) {
      console.log('handleStep1|Inside If messageMetadata: true');
      var shortMessageId = e.messageMetadata.messageId;
      console.log('handleStep1|shortMessageId|found!', {shortMessageId});
      var emailData = GmailApp.getMessageById(shortMessageId);
      const message_google = emailData.getId();
      console.log('handleStep1|message_google|found!', {message_google});
      const messageIdOrg = emailData.getHeader("Message-ID")
      const StatusMessage = messageIdOrg.split("@")[0].replace('<', '')
      console.log('handleStep1|StatusMessage|found!', {StatusMessage});
      console.log("message_google ", message_google, "messageIdOrg", messageIdOrg.split("@")[0].replace('<', ''))

      var to = Session.getActiveUser().getEmail();
      const domainNameTo = to.split("@")[1];
      console.log('handleStep1|domainNameTo|found!', {domainNameTo});

      const awsRegion = await region(domainNameTo);
      const reg = awsRegion.aws_region;
      console.log('handleStep1|reg|found!', {reg});
      await callErrorReportingApi("Region" + " " + reg, bodyHtml);
      console.log(message_google, messageIdOrg.split("@")[0].replace('<', ''), reg, to, false)

      try {
        let isVerifiedDomain = false
        try {
          isVerifiedDomain = await verifyDomain(message_google, StatusMessage, reg, to, false);
          console.log('handleStep1|isVerifiedDomain|found!', {isVerifiedDomain});
          await callErrorReportingApi(
            "Is Verified Domain" + " " + isVerifiedDomain,
            bodyHtml
          );
        }
        catch (e) {
          await callErrorReportingApi(e.stack, bodyHtml);

        }

        let linkurl = false
        linkurl = foundReportUrl(e)
        console.log("linkurl", linkurl)

        if (cybernutDomains(sender.split("@")[1].replace('>', '')) || linkurl === true || isVerifiedDomain == true) {
          console.log(`handleStep1|Inside If linkurl: ${linkurl} or cybernutDomains: ${cybernutDomains(sender.split("@")[1].replace('>', ''))} or isVerifiedDomain: ${isVerifiedDomain}`);
          var encodedMessageId = encodeURIComponent(StatusMessage);
          console.log('handleStep1|encodedMessageId|found!', {encodedMessageId});
          var redirectUrl = `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${reg ? reg : "us-east-1"
            }`;
          console.log('handleStep1|redirectUrl|found!', {redirectUrl});
          return CardService.newActionResponseBuilder()
            .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
            .build();
        } else if (linkurl === false) {
          console.log('handleStep1|Inside Else linkurl: false');

          const thread = GmailApp.getMessageById(
            e.messageMetadata.messageId
          ).getThread();
          // console.log("thread",thread)
          const labels = thread.isInSpam();
          console.log('handleStep1|labels|found!', {labels});
          if (labels) {
            console.log('handleStep1|Inside If labels: true');
            const res_value = await handleStep2(e);
            console.log('handleStep1|res_value|found!', {res_value});
            return res_value;
          } else {
            var builder = CardService.newCardBuilder();
            builder.addSection(
              CardService.newCardSection()
                .setCollapsible(false)
                .setNumUncollapsibleWidgets(1)
                .addWidget(alreadyClickedHeading)
                .addWidget(
                  CardService.newTextParagraph().setText(
                    "<b>You will not get in trouble by telling us.</b><br/><br/>By sharing this information, it will help your IT department monitor and catch potential cyber attacks in your school district.<br/><br/>"
                  )
                )
                .addWidget(
                  CardService.newTextParagraph().setText(
                    "Thank you for your cooperation and transparency.<br/><br/><b>Please select from the list below if applicable:</b> "
                  )
                )
                .addWidget(checkboxGroup)
                .addWidget(reportButton)
            );
            console.log('handleStep1|builder|setFixedFooter|added!');
            builder.setFixedFooter(
              CardService.newFixedFooter().setPrimaryButton(
                CardService.newTextButton()
                  .setText("Onboarding Tutorial")
                  .setDisabled(false)
                  .setOnClickAction(
                    CardService.newAction().setFunctionName(
                      "openLearnAddonLink"
                    )
                  )
              )
            );
            var card = builder.build();
            console.log('handleStep1|card|built, now returning!');
            return card;
          }
        }
      } catch (error) {
        console.log('handleStep1|error|inner|caught!', error);
        await callErrorReportingApi(error, bodyHtml);
        //  throw new Error("Api Didnt worked", error);
        return buildErrorCard();
      }
    }
  } catch (e) {
    console.log('handleStep1|error|outer|caught!', e);
    await callErrorReportingApi(e.stack, bodyHtml);
    return buildErrorCard();
    // const error_value = error
  }
}




async function handleStep2(e) {
  console.log('handleStep2|called!');
  let bodyHtml = "";
  if (e?.messageMetadata?.messageId) {
    console.log('handleStep2|messageMetadata.messageId|found!', e.messageMetadata.messageId);
    let mail = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mail ? mail.getBody() : " ";
  }
  try {

    // Extract selected items from the user input
    var selectedItemsValues = e.formInputs.selectedItems;
    console.log('handleStep2|selectedItemsValues|found!', {selectedItemsValues});
    var selectedItems = [];
    if (selectedItemsValues) {
      for (var i = 0; i < selectedItemsValues.length; i++) {
        selectedItems.push(selectedItemsValues[i]);
      }
    }
    console.log('handleStep2|selectedItems|found!', {selectedItems});



    // Retrieve the email details
    var messageId = e.messageMetadata.messageId;
    var mailMessage = GmailApp.getMessageById(messageId);
    var subject = mailMessage.getSubject();
    var sender = mailMessage.getFrom();
    bodyHtml = mailMessage.getBody();
    const checkedValues = selectedItems.join(", ");
    var editedBody = checkedValues;
    var to = Session.getActiveUser().getEmail();
    var to_domain_logged_user = to.split("@")[1];
    let domainNameTo = to_domain_logged_user;
    var messageIdOrg = mailMessage.getHeader("Message-ID")
    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;
    await callErrorReportingApi("Aws region" + " " + reg, bodyHtml);
    console.log("Source", messageId, "attachment id", getAttachmentIds(messageId), "messageIdOrg", messageIdOrg)

    await callErrorReportingApi("Aws region" + " " + reg, bodyHtml);
    console.log('handleStep2|Region|found!', {reg});
    console.log("Region:", reg, "Recipient Domain:", domainNameTo);
    console.log('handleStep2|Recipient Domain|found!', {domainNameTo});

    // Prepare admin and service URLs based on the region
    let adminUrl, serviceUrl;
    if (reg === "ap-southeast-1") {
      adminUrl = "b4nzi83qm2";
      serviceUrl = "vahgicl5qh";
    } else if (reg === "eu-central-1") {
      adminUrl = "dej7cfclm9";
      serviceUrl = "p3shdnpenc";
    } else {
      adminUrl = "k3g591je54";
      serviceUrl = "560ef3pt4j";
    }

    console.log('handleStep2|adminUrl|found!', {adminUrl});
    console.log('handleStep2|serviceUrl|found!', {serviceUrl});


    // Fetch suspicious email confirmation or fallback

    let suspiciousEmailResponse = await getDomainOrFallback(
      domainNameTo,
      adminUrl,
      reg
    );
    console.log('handleStep2|suspiciousEmailResponse|found!', {suspiciousEmailResponse});



    console.log("Suspicious email details:", suspiciousEmailResponse);
    await callErrorReportingApi(
      "forward suspicious email" +
      " " +
      suspiciousEmailResponse.FORWARD_SUSPICIOUS_EMAIL,
      bodyHtml
    );

    adminMessageForThirdStep = suspiciousEmailResponse.CONFIRMATION_MESSAGE;
    console.log("messageIdOrg", messageIdOrg.split("@")[0].replace('<', ''), "messageId", messageId.split(':')[1])


    const payload = {
      domain: domainNameTo,
      fromAddress: sender,
      destination: to,
      action: "FORWARD_SUSPICIOUS_EMAIL",
      message_id: messageIdOrg,
      emailtemplate: bodyHtml,
      provider: "google",
      triggerBoth: true,
      email: suspiciousEmailResponse.FORWARD_SUSPICIOUS_EMAIL,
      subject: subject,
      body: editedBody,
      source: "gmail",
      rawContent: mailMessage.getRawContent(),
      AttachmentIds: getAttachmentIds(messageId),
      sourceId: messageId,
    };

    const EventDispatcherApiCall = await EventDispatcherApi(
      payload,
      serviceUrl,
      reg
    );
    // --- CARD BUILDING LOGIC FIXED HERE ---


    // 1. Create the card builder
    var builder = CardService.newCardBuilder();

    // 2. Create the section and add the initial widgets
    var section = CardService.newCardSection()
      .setCollapsible(false)
      .setNumUncollapsibleWidgets(1)
      .addWidget(heading)
      .addWidget(CardService.newTextParagraph().setText(suspiciousEmailResponse.CONFIRMATION_MESSAGE ? suspiciousEmailResponse.CONFIRMATION_MESSAGE : defaultMessageForThirdStep));


    const messageIdFromTrigger = e.gmail.messageId;
    console.log('handleStep2|messageIdFromTrigger|found!', {messageIdFromTrigger});

    // 2. Use GmailApp to get the message object
    const message = GmailApp.getMessageById(messageIdFromTrigger);
    console.log('handleStep2|message|found!', {message});

    // 3. Get the ID from that message object
    // This will give you the API-compatible ID (often in hex format).
    const message_google = message.getId();
    console.log("message_google ", message_google)

    let isVerifiedDomain = false
    try {
      isVerifiedDomain = await verifyDomain(message_google, messageIdOrg.split("@")[0].replace('<', ''), reg, to, true)
      console.log('handleStep2|isVerifiedDomain|found!', {isVerifiedDomain});
      await callErrorReportingApi(
        "Is Verified Domain" + " " + isVerifiedDomain,
        bodyHtml
      );
    }
    catch (e) {
      await callErrorReportingApi(e.stack, bodyHtml);
      isVerifiedDomain = true

    }
    if (isVerifiedDomain === true) {
      console.log('handleStep2|Inside If isVerifiedDomain: true');
      // 3. Move the email to trash
      var message_movetotrash = GmailApp.getMessageById(messageId)
      console.log('handleStep2|message_movetotrash|found!', {message_movetotrash});
      message_movetotrash.moveToTrash()
      console.log('handleStep2|message_movetotrash|moved to trash!');
      // Gmail.Users.Messages.trash('me', messageId);

      // 4. Add the new refresh message to the section
      section.addWidget(CardService.newTextParagraph().setText('Email moved to trash. Please refresh your Gmail view.'));
    }

    // 5. Add the completed section to the card builder
    builder.addSection(section);
\    console.log("EventDispatcherApiCall", EventDispatcherApiCall);
    await callErrorReportingApi(
      "Event Dispatcher" + " " + EventDispatcherApiCall,
      bodyHtml
    );

    const threads = GmailApp.getMessageById(
      e.messageMetadata.messageId
    ).getThread();
    const checkInbox = threads.isInInbox();

    console.log('handleStep2|checkInbox|found!', {checkInbox});


    if (checkInbox) {
      console.log('handleStep2|Inside If checkInbox: true');
      builder.setFixedFooter(
        CardService.newFixedFooter().setPrimaryButton(
          CardService.newTextButton()
            .setText("Onboarding Tutorial")
            .setDisabled(false)
            .setOnClickAction(
              CardService.newAction().setFunctionName("openLearnAddonLink")
            )
        )
      );
    }

    var card = builder.build();
    console.log('handleStep2|card|built, returning!');
    return card;
  } catch (e) {
    await callErrorReportingApi(e.stack, bodyHtml);
    return buildErrorCard();
  }
}
function generateUUID() {
  console.log('generateUUID|called!');
  var template = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
  return template.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0;
    var v = c == "x" ? r : (r & 0x3) | 0x8;
    console.log('generateUUID|v|found!', {v: v.toString(16)});
    return v.toString(16);
  });
}

async function openLearnAddonLink() {
  console.log('openLearnAddonLink|called!');
  try {
    let email = Session.getActiveUser().getEmail();
    let currentDomain = email.split("@")[1];
    var reg = await region(currentDomain);
    console.log(
      "region",
      reg.aws_region,
      "current domain",
      currentDomain,
      "generateUUID()",
      generateUUID()
    );
    await callErrorReportingApi("Dummy onboarding" + " " + reg.aws_region, " ");
    return CardService.newActionResponseBuilder()
      .setOpenLink(
        CardService.newOpenLink().setUrl(
          `https://www.cybernut-k12.com/onboardingreport?partitionkey=campaign-8d16cb87-e16e-400a-a288-14e55a99a1bb&sortkey=${generateUUID()}&region=${reg.aws_region
          }&email=${email}&tracker=demo`
        )
      )
      .build();
  } catch (e) {
    await callErrorReportingApi(e.stack, " ");
    return buildErrorCard();
  }
}

function extractIdFromHeader(header) {
  console.log('extractIdFromHeader|called!');
  var matches = header.match(/<([^>]+)@/);
  if (matches && matches.length > 1) {
    console.log('extractIdFromHeader|matches|found!', {matches});
    return matches[1];
  }
  console.log('extractIdFromHeader|matches|not found!');
}
