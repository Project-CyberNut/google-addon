var heading = CardService.newTextParagraph().setText(
  "<b>Cybernut Reporting Tool   </b> v 2.2.0"
);
var alreadyClickedHeading = CardService.newTextParagraph().setText(
  "<b>WAIT - Did you accidentally click on something in this email?</b>"
);
async function callErrorReportingApi(error) {
  try {
    const url =
      "https://560ef3pt4j.execute-api.us-east-1.amazonaws.com/microsoftaddinactivity";
    const payload = {
      id: Session.getActiveUser().getEmail(),
      body: error,
    };
    const options = {
      method: "post",
      headers: { "content-Type": "application/json" },
      payload: JSON.stringify(payload),
    };
    let res = UrlFetchApp.fetch(url, options);
    console.log(
      "Error reporting API called successfully.",
      JSON.stringify(payload),
      error
    );
  } catch (apiError) {
    Logger.log("Failed to call error reporting API: " + apiError.message);
  }
}

async function region(domainNameTo) {
  try {
    let res = UrlFetchApp.fetch(
      `https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com/userregion?domain=${domainNameTo}`,
      {
        method: "get",
        headers: { "content-Type": "application/json" },
      }
    );

    const statusCode = res.getResponseCode();
    const content = res.getContentText();
    const jsonResponse = JSON.parse(content);

    return {
      aws_region: jsonResponse.aws_region,
      status_code: statusCode,
    };
  } catch (error) {
    await callErrorReportingApi(error);
    return {
      aws_region: "us-east-1",
      status_code: error.responseCode || "unknown",
    };
  }
}

async function verifyDomain(fromDomain, messageid, region) {
  try {
    let reg = region;

    let globalUrl;
    if (reg === "ap-southeast-1") {
      globalUrl = "vsqdkxcc8d";
    } else if (reg === "eu-central-1") {
      globalUrl = "telmnzu55i";
    } else {
      globalUrl = "44dgkpf1cb";
    }

    console.log(
      "message in encoding =",
      encodeURIComponent(messageid),
      "from domain",
      fromDomain,
      "url dlobal",
      globalUrl,
      "region",
      reg
    );

    let res = UrlFetchApp.fetch(
      `https://${globalUrl}.execute-api.${reg}.amazonaws.com/admindomainsgoogle?domain=${fromDomain}&messageId=${encodeURIComponent(
        messageid
      )}`,
      {
        method: "get",
        headers: { "content-Type": "application/json" },
      }
    );
    const content = res.getContentText();
    const jsonResponse = JSON.parse(content);
    return jsonResponse.messageExists;
  } catch (error) {
    await callErrorReportingApi(error.stack);
  }
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
  } catch (e) {
    console.log(e);
    await callErrorReportingApi(e.stack);
    throw e;
  }
}

async function handleStep1(e) {
  try {
    var accessToken = e.messageMetadata.accessToken;
    GmailApp.setCurrentMessageAccessToken(accessToken);

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

      var card = cardBuilder.build();
      return card;
    } else {
      var mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
      var sender = mailMessage.getFrom();
      var to = Session.getActiveUser().getEmail();
      var timestamp = new Date();
      timestamp = timestamp.getTime();
    }
    const domainNameTo = to.split("@")[1];
    console.log("domain to", domainNameTo);

    let domainNameFromSenderIndexAtTheRate = sender.indexOf("@");
    let domainNameFromSender = sender.slice(
      domainNameFromSenderIndexAtTheRate + 1
    );
    domainNameFromSender = sender.replace(">", "");

    var fromEmailAddress;
    if (e.messageMetadata) {
      var shortMessageId = e.messageMetadata.messageId;
      var emailData = GmailApp.getMessageById(shortMessageId);
      var rawContent = emailData.getRawContent();

      // Updated regex to extract Message-ID
      var headers = rawContent.match(/^Message-ID:\s*<?([^<>]+)>?/im);

      if (headers && headers[1]) {
        var fullMessageId = headers[1].trim(); // Get the full Message-ID
        var messageId = fullMessageId.split("@")[0]; // Extract before '@'

        console.log("Extracted Message ID Before '@':", messageId);
      } else {
        console.error("Failed to extract Message-ID from headers.");
      }

      if (sender.includes("<")) {
        var regex = /<([^>]+)>/;
        var match = regex.exec(sender);
        if (match && match.length > 1) {
          fromEmailAddress = match[1];
        }
      } else {
        fromEmailAddress = domainNameFromSender;
      }

      var atIndex = fromEmailAddress.indexOf("@");
      if (atIndex !== -1) {
        var domain = fromEmailAddress.substring(atIndex + 1);
        domain = domain.replace(/[<>]/g, "");
      }

      var fromDomain = domain;

      const awsRegion = await region(domainNameTo);
      const reg = awsRegion.aws_region;
      console.log(
        "this is region",
        reg,
        "domain",
        domainNameTo,
        "fromEmailAddress",
        fromEmailAddress,
        "fromDomain",
        fromDomain
      );

      try {
        const isVerifiedDomain = await verifyDomain(fromDomain, messageId, reg);
        console.log(
          "this is region",
          reg,
          "domain",
          domainNameTo,
          "message id",
          messageId,
          "verify domain",
          isVerifiedDomain,
          "from domain",
          fromDomain
        );

        if (isVerifiedDomain) {
          var encodedMessageId = encodeURIComponent(messageId);
          var redirectUrl = `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${
            reg ? reg : "us-east-1"
          }`;
          return CardService.newActionResponseBuilder()
            .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
            .build();
        } else {
          const thread = GmailApp.getMessageById(
            e.messageMetadata.messageId
          ).getThread();
          const labels = thread.isInInbox();

          if (labels) {
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
            return card;
          } else {
            const res_value = await handleStep2(e);
            return res_value;
          }
        }
      } catch (error) {
        Logger.log(error);
      }
    }
  } catch (e) {
    console.log(e.stack);
    await callErrorReportingApi(e.stack);
    throw e;
    // const error_value = error
  }
}

async function handleStep2(e) {
  try {
    // Extract the access token to interact with Gmail API
    var accessToken = e.messageMetadata.accessToken;
    GmailApp.setCurrentMessageAccessToken(accessToken);

    // Extract selected items from the user input
    var selectedItemsValues = e.formInputs.selectedItems;
    var selectedItems = [];
    if (selectedItemsValues) {
      for (var i = 0; i < selectedItemsValues.length; i++) {
        selectedItems.push(selectedItemsValues[i]);
      }
    }

    // Retrieve the email details
    var messageId = e.messageMetadata.messageId;
    var mailMessage = GmailApp.getMessageById(messageId);
    var subject = mailMessage.getSubject();
    var sender = mailMessage.getFrom();
    var bodyHtml = mailMessage.getBody();
    const checkedValues = selectedItems.join(", ");
    var editedBody = checkedValues;
    var to = Session.getActiveUser().getEmail();
    var to_domain_logged_user = to.split("@")[1];

    let domainNameTo = to_domain_logged_user;

    // Extract sender's domain
    let emailAddressSender = sender.match(/<(.+)>/)
      ? sender.match(/<(.+)>/)[1]
      : sender;
    let domainNameFromSenderIndexAtTheRate = emailAddressSender.indexOf("@");
    let fromEmailAddress = emailAddressSender
      .slice(domainNameFromSenderIndexAtTheRate + 1)
      .replace(/[^a-zA-Z0-9.-]/g, "");
    let domainNameFromSender = fromEmailAddress;

    // Extract the message ID from headers
    var headers = mailMessage.getRawContent().match(/^Message-ID: (.+)$/im);
    var messageIdOrg = headers ? extractIdFromHeader(headers[1]) : null;

    // Get the AWS region for the recipient domain
    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;

    console.log("Region:", reg, "Recipient Domain:", domainNameTo);

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

    // Fetch suspicious email confirmation or fallback
    function getDomainOrFallback(domainNameTo, adminUrl, reg) {
      try {
        let response = UrlFetchApp.fetch(
          `https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`,
          {
            method: "post",
            headers: {
              "content-Type": "application/json",
            },
            payload: JSON.stringify({
              domain: domainNameTo,
            }),
          }
        );
        return response;
      } catch (error) {
        var email = Session.getActiveUser().getEmail();
        var currentDomain = email.split("@")[1];
        console.log("Using fallback domain:", currentDomain);
        return UrlFetchApp.fetch(
          `https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`,
          {
            method: "post",
            headers: {
              "content-Type": "application/json",
            },
            payload: JSON.stringify({
              domain: currentDomain,
            }),
          }
        );
      }
    }

    let suspiciousEmailResponse = JSON.parse(
      getDomainOrFallback(domainNameTo, adminUrl, reg)
    );
    console.log("Suspicious email details:", suspiciousEmailResponse);

    adminMessageForThirdStep = suspiciousEmailResponse.CONFIRMATION_MESSAGE;

    // Dispatch an event for the suspicious email
    UrlFetchApp.fetch(
      `https://${serviceUrl}.execute-api.${reg}.amazonaws.com/eventdispatcher`,
      {
        method: "post",
        headers: {
          "content-Type": "application/json",
        },
        payload: JSON.stringify({
          domain: domainNameTo,
          fromAddress: emailAddressSender,
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
        }),
      }
    );

    // Prepare the response card
    var builder = CardService.newCardBuilder();
    builder.addSection(
      CardService.newCardSection()
        .setCollapsible(false)
        .setNumUncollapsibleWidgets(1)
        .addWidget(heading)
        .addWidget(
          CardService.newTextParagraph().setText(
            adminMessageForThirdStep.length > 0
              ? adminMessageForThirdStep
              : defaultMessageForThirdStep
          )
        )
    );

    const isVerifiedDomain = await verifyDomain(
      domainNameFromSender,
      messageId,
      reg
    );

    console.log("Is Verified Domain:", isVerifiedDomain);

    const threads = GmailApp.getMessageById(
      e.messageMetadata.messageId
    ).getThread();
    const checkInbox = threads.isInInbox();

    if (checkInbox) {
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
    return card;
  } catch (e) {
    console.log(e);
    await callErrorReportingApi(e.stack);

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
    console.log(
      "region",
      reg.aws_region,
      "current domain",
      currentDomain,
      "generateUUID()",
      generateUUID()
    );
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
    await callErrorReportingApi(e.stack);
  }
}

function extractIdFromHeader(header) {
  var matches = header.match(/<([^>]+)@/);
  if (matches && matches.length > 1) {
    return matches[1];
  }
}
