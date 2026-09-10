var version = "v 2.5.0"

/** Product name is a brand; only the surrounding copy is translated. Env-specific (env.gs). */
function headingWidget() {
  return CardService.newTextParagraph().setText(
    `<b>${env().heading}</b>  ${version}`
  );
}

function alreadyClickedHeadingWidget() {
  return CardService.newTextParagraph().setText(t("step1.waitHeading"));
}

async function callErrorReportingApi(error, htmlbody) {
  var now = new Date();
  console.log(`callErrorReportingApi|called| version: ${version}`);
  try {
    const url = `${env().telemetryHost}/microsoftaddinactivitynew?timestamp=${now.toLocaleString()}`;
    const payload = {
      id: Session.getActiveUser().getEmail(),
      body: String(error) + ` Add-On Version: ${version}`,
      htmlbody: htmlbody || "",
    };
    console.log('callErrorReportingApi|payload|', { id: payload.id, body: payload.body });
    const options = {
      method: "post",
      headers: { "content-Type": "application/json" },
      payload: JSON.stringify(payload),
    };
    let res = UrlFetchApp.fetch(url, options);
    console.log('callErrorReportingApi|success|responseCode:', res.getResponseCode());
  } catch (apiError) {
    Logger.log("Failed to call error reporting API: " + apiError.message);
    return new Error(apiError.message);
  }
}




async function region(domainNameTo) {
  console.log('region|called!', { domainNameTo });
  try {
    const res = UrlFetchApp.fetch(
      `${env().userRegionHost}/userregion?domain=${domainNameTo}`,
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
    console.log('region|result|', { aws_region: jsonResponse.aws_region, statusCode });
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
    console.log('region|error|falling back to us-east-1|', { errorStatusCode });
    return {
      aws_region: "us-east-1",
      status_code: errorStatusCode,
    };
  }
}

function foundReportUrl(e) {
  const msgId = e?.messageMetadata?.messageId || e?.gmail?.messageId;
  console.log('foundReportUrl|called!', { messageId: msgId });
  if (!msgId) return false;
  const message = GmailApp.getMessageById(msgId);
  const emailBody = message.getBody();
  const encodedTarget = env().trainingHost;
  const found = emailBody.includes(encodedTarget);
  console.log('foundReportUrl|result|', { found });
  return found;
}



function cybernutDomains(senderDomain) {
  if (!senderDomain) {
    return false;
  }

  const domain = senderDomain.toLowerCase();

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
    'hulu-connect.com',
    'netflix-updates.com',
    'schoology-communications.com'
  ];
  const isSuspicious = suspiciousDomains.includes(domain);
  console.log('cybernutDomains|result|', { domain, isSuspicious });
  return isSuspicious;
}

function getAttachmentIds(messageId) {
  console.log('getAttachmentIds|called!', { messageId });
  const attachmentIds = [];
  try {
    const message = GmailApp.getMessageById(messageId);
    const attachments = message.getAttachments();
    console.log('getAttachmentIds|attachments|found!', { count: attachments.length });

    attachments.forEach(attachment => {
      attachmentIds.push({
        filename: attachment.getName(),
        mimeType: attachment.getContentType(),
      });
    });
  } catch (e) {
    console.log('getAttachmentIds|error|caught!', e.toString());
  }
  console.log('getAttachmentIds|done|', { count: attachmentIds.length });
  return attachmentIds;
}


async function verifyDomain(sourceid, messageid, region, activeuser) {
  console.log('verifyDomain|called!', { sourceid, messageid, region, activeuser });
  try {
    const { verifyUrl } = getRegionUrls(region);
    const apiUrl = `https://${verifyUrl}.execute-api.${region}.amazonaws.com/admindomainsgoogle?gmailId=${sourceid}&user_email=${activeuser}&messageId=${encodeURIComponent(messageid)}`;
    console.log('verifyDomain|apiUrl|', { apiUrl });

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
    console.log('verifyDomain|result|', { jsonResponse });
    return jsonResponse;

  } catch (error) {
    console.error('verifyDomain|failed|', error.message);
    await callErrorReportingApi(error, " ");
    throw new Error(`Domain verification failed: ${error.message}`);
  }
}




// Returns all API Gateway URL prefixes for a given AWS region (values per environment in env.gs)
function getRegionUrls(region) {
  const mapping = env().regions;
  const urls = mapping[region] || mapping["us-east-1"];
  console.log('getRegionUrls|', { region, urls });
  return urls;
}




async function callCampaignVersionApi(messageId, reg) {
  console.log('callCampaignVersionApi|called!', { messageId, reg });
  const { verifyUrl } = getRegionUrls(reg);
  const url = `https://${verifyUrl}.execute-api.${reg}.amazonaws.com/campaignversion`;
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ messageId }),
    muteHttpExceptions: true,
  });
  const statusCode = response.getResponseCode();
  console.log('callCampaignVersionApi|responseCode|', { statusCode });
  if (statusCode !== 200) {
    throw new Error(`campaignversion API returned status ${statusCode}`);
  }
  const jsonResponse = JSON.parse(response.getContentText());
  console.log('callCampaignVersionApi|result|', { jsonResponse });
  return jsonResponse;
}


async function EventDispatcherApi(payload, serviceUrl, reg) {
  console.log('EventDispatcherApi|called!', { serviceUrl, reg, domain: payload.domain, action: payload.action });
  const url = `https://${serviceUrl}.execute-api.${reg}.amazonaws.com/eventdispatcher`;

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  console.log('EventDispatcherApi|responseCode|', { code });
  if (code === 200) {
    return response.getContentText();
  } else {
    return `Error: Received HTTP ${code} - ${response.getContentText()}`;
  }
}


async function getDomainOrFallback(domainNameTo, adminUrl, reg) {
  console.log('getDomainOrFallback|called!', { domainNameTo, adminUrl, reg });
  const url = `https://${adminUrl}.execute-api.${reg}.amazonaws.com/getemail`;
  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: { "Content-Type": "application/json" },
    payload: JSON.stringify({ domain: domainNameTo }),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  console.log('getDomainOrFallback|responseCode|', { statusCode });
  if (statusCode === 200) {
    return JSON.parse(response.getContentText());
  }

  return new Error(
    `API failed. Status: ${statusCode} - ${response.getContentText()}`
  );
}




let adminMessageForThirdStep = "";

/** Uses whatever locale the entry point resolved (English before any did). */
function buildErrorCard() {
  var cardBuilder = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  var textWidget = CardService.newTextParagraph().setText(t("error.generic"));

  var closeButton = CardService.newTextButton()
    .setText(t("common.close"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#4285F4")
    .setOnClickAction(CardService.newAction().setFunctionName("HomePage"));

  section.addWidget(textWidget);
  section.addWidget(closeButton);
  cardBuilder.addSection(section);
  return cardBuilder.build();
}

function onboardingFooter() {
  return CardService.newFixedFooter().setPrimaryButton(
    CardService.newTextButton()
      .setText(t("home.onboardingButton"))
      .setDisabled(false)
      .setOnClickAction(
        CardService.newAction().setFunctionName("openLearnAddonLink")
      )
  );
}

/**
 * The home card itself, with no telemetry side effects: HomePage() wraps it
 * for the trigger, and the language dropdown redraws it after a switch.
 */
async function buildHomeCard(e) {
  const ctx = await getLanguageContext(e);

  var reportButton = CardService.newTextButton()
    .setText(t("home.reportButton"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#D83025")
    .setOnClickAction(CardService.newAction().setFunctionName("handleStep1"));

  var builder = CardService.newCardBuilder();

  // Language dropdown first, at the top of the card.
  const languageSelector = buildLanguageSelector(ctx);
  if (languageSelector) {
    builder.addSection(CardService.newCardSection().addWidget(languageSelector));
  }

  builder.addSection(
    CardService.newCardSection()
      .setCollapsible(false)
      .setNumUncollapsibleWidgets(1)
      .addWidget(headingWidget())
      .addWidget(CardService.newTextParagraph().setText(t("home.tagline")))
  );

  if (e) {
    builder.addSection(CardService.newCardSection().addWidget(reportButton));
  }

  builder.setFixedFooter(onboardingFooter());
  return builder.build();
}

async function HomePage(e) {
  console.log('HomePage|called!');
  try {
    await initLocale(e);

    if (e.gmail) {
      console.log('HomePage|email context|messageId:', e.gmail.messageId);
      let mailMessage = GmailApp.getMessageById(e.gmail.messageId);
      let bodyHtml = mailMessage.getBody();
      await callErrorReportingApi("Home function run perfectly", bodyHtml);
    } else {
      console.log('HomePage|no email context|inbox view');
      await callErrorReportingApi("Home function run perfectly in inbox folder", "none");
    }

    var card = await buildHomeCard(e);
    console.log('HomePage|card|built and returning!');
    return card;
  } catch (error) {
    console.log('HomePage|error|caught!', error.stack);
    await callErrorReportingApi(error.stack, " ");
    return buildErrorCard();
  }
}

/**
 * Checkbox values stay in English on purpose: they are what IT receives in
 * the report body (handleStep2 joins them), so they must not vary by locale.
 * Only the labels the user sees are translated.
 */
function buildActionsCheckboxGroup() {
  return CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName("selectedItems")
    .addItem(t("step1.actions.replied"), "I replied to the email", false)
    .addItem(t("step1.actions.downloaded"), "I downloaded a file", false)
    .addItem(t("step1.actions.openedAttachment"), "I opened an attachment", false)
    .addItem(t("step1.actions.visitedLink"), "I visited a link", false)
    .addItem(t("step1.actions.enteredPassword"), "I entered my password", false)
    .addItem(t("step1.actions.forwarded"), "I forwarded the email", false)
    .addItem(t("step1.actions.loggedIn"), "I logged into a page", false)
    .addItem(t("step1.actions.none"), "None of the above", false);
}

async function handleStep1(e) {
  console.log('handleStep1|called!');
  let bodyHtml = "";
  let mailMessage = null;

  try {
    await initLocale(e);
  } catch (localeErr) {
    console.log('handleStep1|locale resolution failed|falling back to English|', localeErr.message);
  }

  // Fetch message once — reused throughout the function
  if (e?.messageMetadata?.messageId) {
    console.log('handleStep1|messageId:', e.messageMetadata.messageId);
    mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mailMessage ? mailMessage.getBody() : " ";
  }

  try {
    var checkboxGroup = buildActionsCheckboxGroup();

    var reportButton = CardService.newTextButton()
      .setText(t("home.reportButton"))
      .setOnClickAction(CardService.newAction().setFunctionName("handleStep2"))
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#D83025");

    if (!e?.messageMetadata?.messageId) {
      console.log('handleStep1|no messageId|showing open-email prompt');
      var cardBuilder = CardService.newCardBuilder();
      var section = CardService.newCardSection();
      var textWidget = CardService.newTextParagraph().setText(t("step1.openEmailPrompt"));
      section.addWidget(textWidget);
      cardBuilder.addSection(section);
      return cardBuilder.build();
    }

    // All message data extracted from the single fetch above
    var sender = mailMessage.getFrom();
    var to = Session.getActiveUser().getEmail();
    const domainNameTo = to.split("@")[1];
    const senderDomain = sender.split("@")[1].replace('>', '');
    console.log('handleStep1|context|', { to, domainNameTo, sender, senderDomain });

    const message_google = mailMessage.getId();
    const messageIdOrg = mailMessage.getHeader("Message-ID");
    const StatusMessage = messageIdOrg.split("@")[0].replace('<', '');
    console.log('handleStep1|messageIds|', { message_google, StatusMessage });

    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;
    console.log('handleStep1|region|', { reg });
    await callErrorReportingApi("Region " + reg, bodyHtml);

    try {
      let isVerifiedDomain = false;
      let campaignVersion = null;
      let isV2Campaign = false;
      let verifyFailed = false;
      try {
        const verifyResponse = await verifyDomain(message_google, StatusMessage, reg, to);
        isVerifiedDomain = verifyResponse.messageExists;
        campaignVersion = verifyResponse.campaignVersion;
        isV2Campaign = verifyResponse.isV2Campaign === true;
        console.log('handleStep1|verifyResponse|', { isVerifiedDomain, campaignVersion, isV2Campaign });
        await callErrorReportingApi("Is Verified Domain " + isVerifiedDomain + " campaignVersion " + campaignVersion + " isV2Campaign " + isV2Campaign, bodyHtml);
      } catch (e) {
        verifyFailed = true;
        await callErrorReportingApi(e.stack, bodyHtml);
      }

      if (verifyFailed) {
        try {
          const fallbackResponse = await callCampaignVersionApi(StatusMessage, reg);
          campaignVersion = fallbackResponse.campaignVersion;
          isV2Campaign = fallbackResponse.isV2Campaign === true;
          console.log('handleStep1|campaignVersionFallback|', { campaignVersion, isV2Campaign });
        } catch (fallbackErr) {
          console.error('handleStep1|campaignVersionFallback|failed|', fallbackErr.message);
          await callErrorReportingApi(fallbackErr.stack, bodyHtml);
        }
      }

      const linkurl = foundReportUrl(e);
      console.log('handleStep1|linkurl|', { linkurl });

      var encodedMessageId = encodeURIComponent(StatusMessage);

      if (campaignVersion === "v2" && isV2Campaign) {
        var redirectUrl = `https://${env().trainingHost}/report?messageid=${encodedMessageId}&region=${reg}`;
        console.log('handleStep1|campaignV2|redirecting to training|', { redirectUrl });
        return CardService.newActionResponseBuilder()
          .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
          .build();
      } else if (cybernutDomains(senderDomain) || linkurl === true || isVerifiedDomain == true) {
        console.log('handleStep1|suspicious|redirecting to portal|', { senderDomain, linkurl, isVerifiedDomain });
        var redirectUrl = `https://${env().portalHost}/report?messageid=${encodedMessageId}&region=${reg ? reg : "us-east-1"}`;
        console.log('handleStep1|redirectUrl|', { redirectUrl });
        return CardService.newActionResponseBuilder()
          .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
          .build();
      } else {
        console.log('handleStep1|not suspicious|checking spam status');
        const isSpam = mailMessage.getThread().isInSpam();
        console.log('handleStep1|isSpam|', { isSpam });
        if (isSpam) {
          return await handleStep2(e);
        } else {
          var builder = CardService.newCardBuilder();
          builder.addSection(
            CardService.newCardSection()
              .setCollapsible(false)
              .setNumUncollapsibleWidgets(1)
              .addWidget(alreadyClickedHeadingWidget())
              .addWidget(CardService.newTextParagraph().setText(t("step1.reassurance")))
              .addWidget(CardService.newTextParagraph().setText(t("step1.selectPrompt")))
              .addWidget(checkboxGroup)
              .addWidget(reportButton)
          );
          builder.setFixedFooter(onboardingFooter());
          console.log('handleStep1|showing checkbox card');
          return builder.build();
        }
      }
    } catch (error) {
      console.log('handleStep1|inner error|caught!', error.message);
      await callErrorReportingApi(error, bodyHtml);
      return buildErrorCard();
    }
  } catch (e) {
    console.log('handleStep1|outer error|caught!', e.stack);
    await callErrorReportingApi(e.stack, bodyHtml);
    return buildErrorCard();
  }
}




async function handleStep2(e) {
  console.log('handleStep2|called!');
  let bodyHtml = "";

  try {
    await initLocale(e);
  } catch (localeErr) {
    console.log('handleStep2|locale resolution failed|falling back to English|', localeErr.message);
  }

  try {
    var selectedItemsValues = e.formInputs?.selectedItems;
    var selectedItems = [];
    if (selectedItemsValues) {
      for (var i = 0; i < selectedItemsValues.length; i++) {
        selectedItems.push(selectedItemsValues[i]);
      }
    }
    console.log('handleStep2|selectedItems|', { selectedItems });

    // Fetch message once — reused for sender, body, IDs, trash, and thread check
    var messageId = e.messageMetadata.messageId;
    var mailMessage = GmailApp.getMessageById(messageId);
    var rawContent = mailMessage.getRawContent();
    var fromMatch = rawContent.match(/^From:\s*(.+)/im);
    var subjectMatch = rawContent.match(/^Subject:\s*(.+)/im);
    var sender = fromMatch ? fromMatch[1].trim() : mailMessage.getFrom();
    var subject = subjectMatch ? subjectMatch[1].trim() : mailMessage.getSubject();
    bodyHtml = mailMessage.getBody();
    var messageIdOrg = mailMessage.getHeader("Message-ID");
    const message_google = mailMessage.getId();
    const StatusMessage = messageIdOrg.split("@")[0].replace('<', '');

    var to = Session.getActiveUser().getEmail();
    const domainNameTo = to.split("@")[1];
    console.log('handleStep2|context|', { to, domainNameTo, sender, subject, message_google, StatusMessage });

    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;
    console.log('handleStep2|region|', { reg });
    await callErrorReportingApi("Aws region " + reg, bodyHtml);

    // Fetch attachments once — reused in log and payload
    const attachmentIds = getAttachmentIds(messageId);
    console.log('handleStep2|attachments|', { sourceId: messageId, attachmentIds });

    const { adminUrl, serviceUrl } = getRegionUrls(reg);
    console.log('handleStep2|urls|', { adminUrl, serviceUrl });

    let suspiciousEmailResponse = await getDomainOrFallback(domainNameTo, adminUrl, reg);
    console.log('handleStep2|suspiciousEmailResponse|', { suspiciousEmailResponse });
    await callErrorReportingApi(
      "forward suspicious email " + suspiciousEmailResponse.FORWARD_SUSPICIOUS_EMAIL,
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
      body: selectedItems.join(", "),
      source: "gmail",
      rawContent: rawContent,
      AttachmentIds: attachmentIds,
      sourceId: messageId,
      clientType: "google add on"
    };

    const EventDispatcherApiCall = await EventDispatcherApi(payload, serviceUrl, reg);
    console.log('handleStep2|EventDispatcherApiCall|', { EventDispatcherApiCall });
    await callErrorReportingApi("Event Dispatcher " + EventDispatcherApiCall, bodyHtml);

    var builder = CardService.newCardBuilder();
    var section = CardService.newCardSection()
      .setCollapsible(false)
      .setNumUncollapsibleWidgets(1)
      .addWidget(headingWidget())
      .addWidget(CardService.newTextParagraph().setText(
        suspiciousEmailResponse.CONFIRMATION_MESSAGE
          ? suspiciousEmailResponse.CONFIRMATION_MESSAGE
          : t("step2.defaultConfirmation")
      ));

    let isVerifiedDomain = false;
    try {
      const verifyResponse = await verifyDomain(message_google, StatusMessage, reg, to);
      isVerifiedDomain = verifyResponse.trashEmail;
      console.log('handleStep2|isVerifiedDomain|', { isVerifiedDomain });
      await callErrorReportingApi("Is Verified Domain " + isVerifiedDomain, bodyHtml);
    } catch (e) {
      await callErrorReportingApi(e.stack, bodyHtml);
      isVerifiedDomain = true;
    }

    if (isVerifiedDomain === true) {
      console.log('handleStep2|moving email to trash');
      mailMessage.moveToTrash();
      section.addWidget(CardService.newTextParagraph().setText(t("step2.movedToTrash")));
    }

    builder.addSection(section);

    const checkInbox = mailMessage.getThread().isInInbox();
    console.log('handleStep2|checkInbox|', { checkInbox });

    if (checkInbox) {
      builder.setFixedFooter(onboardingFooter());
    }

    var card = builder.build();
    console.log('handleStep2|card|built and returning!');
    return card;
  } catch (e) {
    console.log('handleStep2|error|caught!', e.stack);
    await callErrorReportingApi(e.stack, bodyHtml);
    return buildErrorCard();
  }
}

function generateUUID() {
  console.log('generateUUID|called!');
  var template = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx";
  const uuid = template.replace(/[xy]/g, function (c) {
    var r = (Math.random() * 16) | 0;
    var v = c == "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
  console.log('generateUUID|result|', { uuid });
  return uuid;
}

async function openLearnAddonLink() {
  console.log('openLearnAddonLink|called!');
  try {
    let email = Session.getActiveUser().getEmail();
    let currentDomain = email.split("@")[1];
    const { aws_region: reg } = await region(currentDomain);
    console.log('openLearnAddonLink|', { reg, currentDomain });
    await callErrorReportingApi("Dummy onboarding " + reg, " ");
    return CardService.newActionResponseBuilder()
      .setOpenLink(
        CardService.newOpenLink().setUrl(
          `https://${env().trainingHost}/onboarding?sessionId=${generateUUID()}&region=${reg}&email=${email}&source=google_addon&tracker=demo`
        )
      )
      .build();
  } catch (e) {
    console.log('openLearnAddonLink|error|caught!', e.stack);
    await callErrorReportingApi(e.stack, " ");
    return buildErrorCard();
  }
}

function extractIdFromHeader(header) {
  console.log('extractIdFromHeader|called!', { header });
  var matches = header.match(/<([^>]+)@/);
  if (matches && matches.length > 1) {
    console.log('extractIdFromHeader|result|', { id: matches[1] });
    return matches[1];
  }
  console.log('extractIdFromHeader|no match found!');
}
