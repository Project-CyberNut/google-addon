var heading = CardService.newTextParagraph().setText(
  "<b>Cybernut Reporting Tool   </b> v 2.2.2"
);
var alreadyClickedHeading = CardService.newTextParagraph().setText(
  "<b>WAIT - Did you accidentally click on something in this email?</b>"
);

/**
 * CENTRALIZED CONFIGURATION
 * This function holds all the correct API Gateway endpoints.
 * @param {string} region The AWS region (e.g., "us-east-1").
 * @return {object} An object containing the adminUrl, globalUrl, and serviceUrl.
 */
function getApiEndpoints(region) {
  const endpoints = {
    "us-east-1": {
      adminUrl: "bn1b2bl6xk",   // From admin-api-stack-t-admin-api
      globalUrl: "aqh9osmw28",  // From global-stack-t-api
      serviceUrl: "hhszpv1cxd"   // UPDATED: From reported-threats-t-api
    },
    "ap-southeast-1": {
      adminUrl: "9fkxy40dfk",
      globalUrl: "43v1dfp0n3",
      serviceUrl: "jktofxjij2"   // UPDATED: From reported-threats-t-api
    },
    "eu-central-1": {
      adminUrl: "h9j9uzoo9h",
      globalUrl: "napssgoubc",
      serviceUrl: "5b38mlbr9e"   // UPDATED: From reported-threats-t-api
    }
  };
  // Return the specific region's endpoints, or default to us-east-1
  return endpoints[region] || endpoints["us-east-1"];
}


async function callErrorReportingApi(error, htmlbody) {
  var now = new Date();
  try {
    // NOTE: This URL uses a specific API ID not found in the screenshots.
    // Assuming this is a separate, correctly configured error reporting API.
    const url = `https://hhszpv1cxd.execute-api.us-east-1.amazonaws.com/microsoftaddinactivitynew?timestamp=${now.toLocaleString()}`;
    const payload = {
      id: Session.getActiveUser().getEmail(),
      body: error,
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
    // NOTE: This URL uses a specific API ID not found in the screenshots.
    // Assuming this is a separate, correctly configured region lookup API.
    const res = UrlFetchApp.fetch(
      `https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com/userregion?domain=${domainNameTo}`,
      {
        method: "get",
        headers: { "Content-Type": "application/json" },
        muteHttpExceptions: false,
      }
    );

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

async function verifyDomain(fromDomain, messageid, region, globalUrl) {
  const maxRetries = 3;
  const initialDelay = 1000;
  let lastError;

  try {
    // NOTE: The stage name (e.g., /prod/) might be needed here if not using the $default stage.
    const apiUrl = `https://${globalUrl}.execute-api.${region}.amazonaws.com/admindomainsgoogle?domain=${fromDomain}&messageId=${encodeURIComponent(
      messageid
    )}`;

    console.log(
      `Verifying domain: ${fromDomain}`,
      `Message ID: ${messageid}`,
      `API URL: ${apiUrl}`
    );

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = UrlFetchApp.fetch(apiUrl, {
          method: "GET",
          headers: { "Content-Type": "application/json" },
          muteHttpExceptions: false,
        });

        const statusCode = response.getResponseCode();
        if (statusCode !== 200) {
          throw new Error(`API returned status ${statusCode}`);
        }

        const jsonResponse = JSON.parse(response.getContentText());
        return jsonResponse.messageExists;
      } catch (error) {
        lastError = error;
        console.warn(`Attempt ${attempt} failed: ${error.message}`);
        if (attempt < maxRetries) {
          const delay = initialDelay * Math.pow(2, attempt);
          console.log(`Retrying in ${delay}ms...`);
          Utilities.sleep(delay);
        }
      }
    }
    throw new Error(
      `All ${maxRetries} attempts failed. Last error: ${lastError.message}`
    );
  } catch (error) {
    await callErrorReportingApi(error, " ");
    throw new Error(`Domain verification failed: ${error.message}`);
  }
}


async function EventDispatcherApi(payload, serviceUrl, reg) {
  // NOTE: The stage name (e.g., /prod/) might be needed here if not using the $default stage.
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
  // NOTE: The stage name (e.g., /prod/) might be needed here if not using the $default stage.
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

async function HomePage(e) {
  try {
    var reportButton = CardService.newTextButton()
      .setText("Report Email")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#D83025")
      .setOnClickAction(CardService.newAction().setFunctionName("handleStep1"));

    var builder = CardService.newCardBuilder();
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
    if (e.gmail) {
      let mailMessage = GmailApp.getMessageById(e.gmail.messageId);
      let bodyHtml = mailMessage.getBody();
      await callErrorReportingApi("Home function run perfectly", bodyHtml);
    } else {
      await callErrorReportingApi(
        "Home function run perfectly in inbox folder",
        "none"
      );
    }

    if (e) {
      builder.addSection(CardService.newCardSection().addWidget(reportButton));
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

    var card = builder.build();
    return card;
  } catch (error) {
    await callErrorReportingApi(error.stack, " ");
    throw error;
  }
}

async function handleStep1(e) {
  let bodyHtml = "";
  if (e.messageMetadata.messageId) {
    let mail = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mail ? mail.getBody() : " ";
  }

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

    var reportButton = CardService.newTextButton()
      .setText("Report Email")
      .setOnClickAction(CardService.newAction().setFunctionName("handleStep2"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#D83025");

    if (!e.messageMetadata.messageId) {
      var cardBuilder = CardService.newCardBuilder();
      var section = CardService.newCardSection();
      var textWidget = CardService.newTextParagraph().setText(
        "Please open the email and look for the button in the top left corner. Click on it to go back and find the report button."
      );
      section.addWidget(textWidget);
      cardBuilder.addSection(section);
      return cardBuilder.build();
    }

    var mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
    var sender = mailMessage.getFrom();
    var to = Session.getActiveUser().getEmail();
    const domainNameTo = to.split("@")[1];
    var fromEmailAddress;

    if (e.messageMetadata) {
      var shortMessageId = e.messageMetadata.messageId;
      var emailData = GmailApp.getMessageById(shortMessageId);
      var rawContent = emailData.getRawContent();
      var headers = rawContent.match(/^Message-ID:\s*<?([^<>]+)>?/im);
      var messageId = "";

      if (headers && headers[1]) {
        var fullMessageId = headers[1].trim();
        messageId = fullMessageId.split("@")[0];
        console.log("Extracted Message ID Before '@':", messageId);
      } else {
        console.log("Failed to extract Message-ID from headers.");
      }

      if (sender.includes("<")) {
        var regex = /<([^>]+)>/;
        var match = regex.exec(sender);
        if (match && match.length > 1) {
          fromEmailAddress = match[1];
        }
      } else {
        fromEmailAddress = sender;
      }

      var atIndex = fromEmailAddress.indexOf("@");
      var fromDomain = "";
      if (atIndex !== -1) {
        fromDomain = fromEmailAddress.substring(atIndex + 1).replace(/[<>]/g, "");
      }
      
      const awsRegion = await region(domainNameTo);
      const reg = awsRegion.aws_region;
      const endpoints = getApiEndpoints(reg);
      
      await callErrorReportingApi("Region" + " " + reg, bodyHtml);

      try {
        const isVerifiedDomain = await verifyDomain(fromDomain, messageId, reg, endpoints.globalUrl);
        await callErrorReportingApi(
          "Is Verified Domain" + " " + isVerifiedDomain,
          bodyHtml
        );

        if (isVerifiedDomain === true) {
          var encodedMessageId = encodeURIComponent(messageId);
          var redirectUrl = `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${
            reg ? reg : "us-east-1"
          }`;
          return CardService.newActionResponseBuilder()
            .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
            .build();
        } else if (isVerifiedDomain === false) {
          const thread = GmailApp.getMessageById(e.messageMetadata.messageId).getThread();
          const labels = thread.isInSpam();
          if (labels) {
            return await handleStep2(e);
          } else {
            var builder = CardService.newCardBuilder();
            builder.addSection(
              CardService.newCardSection()
                .setCollapsible(false)
                .setNumUncollapsibleWidgets(1)
                .addWidget(alreadyClickedHeading)
                .addWidget(CardService.newTextParagraph().setText("<b>You will not get in trouble by telling us.</b><br/><br/>By sharing this information, it will help your IT department monitor and catch potential cyber attacks in your school district.<br/><br/>"))
                .addWidget(CardService.newTextParagraph().setText("Thank you for your cooperation and transparency.<br/><br/><b>Please select from the list below if applicable:</b> "))
                .addWidget(checkboxGroup)
                .addWidget(reportButton)
            );
            builder.setFixedFooter(
              CardService.newFixedFooter().setPrimaryButton(
                CardService.newTextButton()
                  .setText("Onboarding Tutorial")
                  .setDisabled(false)
                  .setOnClickAction(CardService.newAction().setFunctionName("openLearnAddonLink"))
              )
            );
            return builder.build();
          }
        }
      } catch (error) {
        await callErrorReportingApi(error, bodyHtml);
        throw new Error("Api Didnt worked", error);
      }
    }
  } catch (e) {
    await callErrorReportingApi(e.stack, bodyHtml);
    throw e;
  }
}

async function handleStep2(e) {
  let bodyHtml = "";
  if (e.messageMetadata.messageId) {
    let mail = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mail ? mail.getBody() : " ";
  }
  try {
    var selectedItemsValues = e.formInputs.selectedItems;
    var selectedItems = [];
    if (selectedItemsValues) {
      for (var i = 0; i < selectedItemsValues.length; i++) {
        selectedItems.push(selectedItemsValues[i]);
      }
    }

    var messageId = e.messageMetadata.messageId;
    var mailMessage = GmailApp.getMessageById(messageId);
    var subject = mailMessage.getSubject();
    var sender = mailMessage.getFrom();
    const checkedValues = selectedItems.join(", ");
    var editedBody = checkedValues;
    var to = Session.getActiveUser().getEmail();
    var domainNameTo = to.split("@")[1];
    var headers = mailMessage.getRawContent().match(/^Message-ID: (.+)$/im);
    var messageIdOrg = headers && headers[1] ? headers[1].trim() : null;

    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;
    const { adminUrl, serviceUrl } = getApiEndpoints(reg);
    
    await callErrorReportingApi("Aws region" + " " + reg, bodyHtml);
    console.log("Region:", reg, "Recipient Domain:", domainNameTo);

    let suspiciousEmailResponse = await getDomainOrFallback(domainNameTo, adminUrl, reg);
    console.log("Suspicious email details:", suspiciousEmailResponse);
    await callErrorReportingApi(
      "forward suspicious email" + " " + suspiciousEmailResponse.FORWARD_SUSPICIOUS_EMAIL,
      bodyHtml
    );

    adminMessageForThirdStep = suspiciousEmailResponse.CONFIRMATION_MESSAGE;

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
    };

    const EventDispatcherApiCall = await EventDispatcherApi(payload, serviceUrl, reg);

    var builder = CardService.newCardBuilder();
    builder.addSection(
      CardService.newCardSection()
        .setCollapsible(false)
        .setNumUncollapsibleWidgets(1)
        .addWidget(heading)
        .addWidget(CardService.newTextParagraph().setText(defaultMessageForThirdStep))
    );
    console.log("EventDispatcherApiCall", EventDispatcherApiCall);
    await callErrorReportingApi(
      "Event Dispatcher" + " " + EventDispatcherApiCall,
      bodyHtml
    );

    const threads = GmailApp.getMessageById(e.messageMetadata.messageId).getThread();
    const checkInbox = threads.isInInbox();

    if (checkInbox) {
      builder.setFixedFooter(
        CardService.newFixedFooter().setPrimaryButton(
          CardService.newTextButton()
            .setText("Onboarding Tutorial")
            .setDisabled(false)
            .setOnClickAction(CardService.newAction().setFunctionName("openLearnAddonLink"))
        )
      );
    }
    return builder.build();
  } catch (e) {
    await callErrorReportingApi(e.stack, bodyHtml);
    throw e;
  }
}

function generateUUID() {
  var template = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
  return template.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0;
    var v = c == "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

async function openLearnAddonLink() {
  try {
    let email = Session.getActiveUser().getEmail();
    let currentDomain = email.split("@")[1];
    var reg = await region(currentDomain);
    await callErrorReportingApi("Dummy onboarding" + " " + reg.aws_region, " ");
    return CardService.newActionResponseBuilder()
      .setOpenLink(
        CardService.newOpenLink().setUrl(
          `https://www.cybernut-k12.com/onboardingreport?partitionkey=campaign-8d16cb87-e16e-400a-a288-14e55a99a1bb&sortkey=${generateUUID()}&region=${
            reg.aws_region
          }&email=${email}&tracker=demo`
        )
      )
      .build();
  } catch (e) {
    await callErrorReportingApi(e.stack, " ");
  }
}

function extractIdFromHeader(header) {
  var matches = header.match(/<([^>]+)@/);
  if (matches && matches.length > 1) {
    return matches[1];
  }
}



