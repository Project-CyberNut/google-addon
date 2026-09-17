var version = "v 2.4.5"
var heading = CardService.newTextParagraph().setText(
  `<b>Cybernut Reporting Tool</b>  ${version}`
);
var alreadyClickedHeading = CardService.newTextParagraph().setText(
  "<b>WAIT - Did you accidentally click on something in this email?</b>"
);

async function callErrorReportingApi(error, htmlbody) {
  var now = new Date();
  console.log(`callErrorReportingApi|called| version: ${version}`);
  try {
    const url = `https://560ef3pt4j.execute-api.us-east-1.amazonaws.com/microsoftaddinactivitynew?timestamp=${now.toLocaleString()}`;
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

// Landing hosts a simulation body can point at. v1 campaigns render
// https://www.cybernut-k12.com/report?... (TrainingCampaignEmail.ts) and v2/v2.3 render
// https://training.cybernut.com/report?... (TrainingCampaignEmailV2.ts getReportBaseUrl).
// Checking only one of the two is what let v1 simulations fall through to the threat path.
var CAMPAIGN_LANDING_HOSTS = [
  'www.cybernut-k12.com',
  'training.cybernut.com'
];

// Takes the body the caller already fetched - handleStep1 has it, and a second
// GmailApp.getMessageById() is pure cost inside the 30s Card-service callback budget.
function foundReportUrl(emailBody) {
  try {
    if (!emailBody) {
      console.log('foundReportUrl|result|', { found: false, reason: 'no body' });
      return false;
    }

    const raw = String(emailBody).toLowerCase();
    const haystacks = [raw];
    // addClickTracking() rewrites every href to
    // https://<tracker>/click?destUrl=<percent-encoded original>. Host names survive
    // encoding intact, but decode anyway so a doubly-wrapped link still matches.
    // decodeURIComponent throws on malformed sequences, hence the inner catch.
    try {
      const decoded = decodeURIComponent(raw);
      if (decoded !== raw) haystacks.push(decoded);
    } catch (decodeError) {
      console.log('foundReportUrl|decode skipped|', decodeError.toString());
    }

    for (var h = 0; h < haystacks.length; h++) {
      const hay = haystacks[h];

      for (var i = 0; i < CAMPAIGN_LANDING_HOSTS.length; i++) {
        if (hay.indexOf(CAMPAIGN_LANDING_HOSTS[i]) !== -1) {
          console.log('foundReportUrl|result|', { found: true, matched: CAMPAIGN_LANDING_HOSTS[i] });
          return true;
        }
      }

      // Difficulty-4 campaigns point at the lookalike domain itself
      // (getDifficultyLevelFourReportBaseUrl), so neither cybernut host appears.
      // Anchor on // or . so we match a host, not a mention in prose.
      for (var j = 0; j < SIMULATION_SENDER_DOMAINS.length; j++) {
        const d = SIMULATION_SENDER_DOMAINS[j];
        if (hay.indexOf('//' + d) !== -1 || hay.indexOf('.' + d) !== -1) {
          console.log('foundReportUrl|result|', { found: true, matched: d });
          return true;
        }
      }
    }

    console.log('foundReportUrl|result|', { found: false });
    return false;
  } catch (error) {
    console.log('foundReportUrl|error|caught!', error.toString());
    return false;
  }
}



// Sender domains used by simulation campaigns. Hoisted to module scope so
// foundReportUrl() can also scan the message body for them - difficulty-4 campaigns
// land on the lookalike domain itself rather than a cybernut host, so the body is the
// only place they show up.
var SIMULATION_SENDER_DOMAINS = [
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
    'schoology-communications.com',
    // Used by difficultyLevelFourDestinationDomains (login.zoom-services.com,
    // login.amzn-alerts.com) but previously absent from this list.
    'zoom-services.com',
    'amzn-alerts.com',
    // Reconciled against campaign GENERIC_SENDER values, 17 Sep 2026: the list had
    // drifted 45 domains behind since it was last updated 11 Dec 2025 (346bb8a).
    'aeries-portal.com',
    'blackbaud-connect-portal.com',
    'blackboard-portal.com',
    'bloomz-portal.com',
    'classcraft-portal.com',
    'classdojo-portal.com',
    'clever-portal.com',
    'contentkeeper-by-impero-portal.com',
    'coursera-portal.com',
    'district-colleague.com',
    'district-hr-office.com',
    'district-it-director.com',
    'district-parent.com',
    'district-principal.com',
    'district-superintendent.com',
    'educlimber-portal.com',
    'eschool-data-portal.com',
    'facebook-security-notice.com',
    'facts-sis-portal.com',
    'ferpa-compliance-office-notice.com',
    'formative-portal.com',
    'goguardian-portal.com',
    'hero-k12-portal.com',
    'i-ready-portal.com',
    'infinite-campus-portal.com',
    'lightspeed-systems-portal.com',
    'linkdn-learning-portal.com',
    'linq-portal.com',
    'map-portal.com',
    'nwea-portal.com',
    'parentsquare-portal.com',
    'paypaal-alerts.com',
    'paypal-account-notice.com',
    'pbis-rewards-portal.com',
    'pear-assessment-portal.com',
    'quickbooks-for-education-portal.com',
    'remind-portal.com',
    'schoolmessenger-portal.com',
    'schoolstatus-portal.com',
    'screencastify-portal.com',
    'skyward-portal.com',
    'stafftrac-eduvistas-portal.com',
    'state-education-department-notice.com',
    'tax-notice-services.com',
    'vector-solutions-portal.com'
];

// Port of parseDomainFromSender() in
// cybernut-selector-service/lambda/TrainingCampaignEmailV2.ts - the same senders this
// has to recognise are built there. Pulls the address out of "Name <user@domain>"
// first, so a display name containing an "@" (common in spoof templates) no longer
// makes the naive split("@")[1] return garbage and miss the domain list.
function parseDomainFromSender(sender) {
  if (!sender) return '';
  const match = String(sender).match(/<([^>]+)>/);
  const email = match ? match[1] : String(sender);
  const parts = email.split('@');
  return parts.length > 1 ? parts[parts.length - 1].trim().toLowerCase() : '';
}

function cybernutDomains(senderDomain) {
  if (!senderDomain) {
    return false;
  }

  const domain = senderDomain.toLowerCase();

  // Difficulty-4 campaigns send from login./signin./auth. subdomains of these, so an
  // exact match is not enough - the backend's own matcher accepts subdomains too.
  const isSuspicious = SIMULATION_SENDER_DOMAINS.some(function (d) {
    return domain === d || domain.slice(-(d.length + 1)) === '.' + d;
  });
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


// A Workspace admin (or our own urlFetchWhitelist) can block the verification call before it
// ever leaves Google. That reads nothing like an API error, so classify it separately - the
// card copy and the alerting both need to tell "we were blocked" apart from "the API failed".
function isUrlFetchBlocked(error) {
  const message = (error && error.message) || String(error);
  return message.indexOf("not permitted by your admin") !== -1;
}




// Returns all API Gateway URL prefixes for a given AWS region
function getRegionUrls(region) {
  const mapping = {
    "us-east-1":      { verifyUrl: "44dgkpf1cb", adminUrl: "k3g591je54", serviceUrl: "560ef3pt4j" },
    "ap-southeast-1": { verifyUrl: "vsqdkxcc8d", adminUrl: "b4nzi83qm2", serviceUrl: "vahgicl5qh" },
    "eu-central-1":   { verifyUrl: "telmnzu55i", adminUrl: "dej7cfclm9", serviceUrl: "p3shdnpenc" },
  };
  const urls = mapping[region] || { verifyUrl: "44dgkpf1cb", adminUrl: "k3g591je54", serviceUrl: "560ef3pt4j" };
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




let defaultMessageForThirdStep =
  "Thank you, you will hear back from IT if you need to take any further action.";
let adminMessageForThirdStep = "";

function buildErrorCard() {
  var cardBuilder = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  var textWidget = CardService.newTextParagraph().setText(
    "There was an error in completing your action, for escalation / faster resolution you can contact us at support@cybernut.com"
  );

  var closeButton = CardService.newTextButton()
    .setText("Close")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#4285F4")
    .setOnClickAction(CardService.newAction().setFunctionName("HomePage"));

  section.addWidget(textWidget);
  section.addWidget(closeButton);
  cardBuilder.addSection(section);
  return cardBuilder.build();
}


// Shown when we could not check whether this email is one of our own campaign sends. We stop
// here rather than guess: reading an unavailable verifier as "not a simulation" is what filed
// real training emails as threats and trashed them. Say plainly that nothing was reported.
function buildVerificationUnavailableCard(blocked) {
  var cardBuilder = CardService.newCardBuilder();
  var section = CardService.newCardSection();

  var message = blocked
    ? "We could not check this email because your Google Workspace administrator has blocked the add-on's connection to CyberNut.<br/><br/><b>This email has not been reported.</b><br/><br/>Please ask your IT administrator to allow the CyberNut Reporting Tool, or contact us at support@cybernut.com"
    : "Verification is temporarily unavailable, so we could not check this email.<br/><br/><b>This email has not been reported.</b><br/><br/>Please try again shortly, or contact us at support@cybernut.com";

  var closeButton = CardService.newTextButton()
    .setText("Close")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#4285F4")
    .setOnClickAction(CardService.newAction().setFunctionName("HomePage"));

  section.addWidget(CardService.newTextParagraph().setText(message));
  section.addWidget(closeButton);
  cardBuilder.addSection(section);
  return cardBuilder.build();
}




async function HomePage(e) {
  console.log('HomePage|called!');
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
      // The homepage trigger is unconditional, so this runs for every message the user
      // opens. It used to fetch and POST the full HTML body each time; nothing consumed
      // it, so the fetch is gone with it.
      console.log('HomePage|email context|messageId:', e.gmail.messageId);
      await callErrorReportingApi("Home function run perfectly", "none");
    } else {
      console.log('HomePage|no email context|inbox view');
      await callErrorReportingApi("Home function run perfectly in inbox folder", "none");
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
    console.log('HomePage|card|built and returning!');
    return card;
  } catch (error) {
    console.log('HomePage|error|caught!', error.stack);
    await callErrorReportingApi(error.stack, " ");
    return buildErrorCard();
  }
}

async function handleStep1(e) {
  console.log('handleStep1|called!');
  let bodyHtml = "";
  let mailMessage = null;

  // Fetch message once — reused throughout the function
  if (e?.messageMetadata?.messageId) {
    console.log('handleStep1|messageId:', e.messageMetadata.messageId);
    mailMessage = GmailApp.getMessageById(e.messageMetadata.messageId);
    bodyHtml = mailMessage ? mailMessage.getBody() : " ";
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

    if (!e?.messageMetadata?.messageId) {
      console.log('handleStep1|no messageId|showing open-email prompt');
      var cardBuilder = CardService.newCardBuilder();
      var section = CardService.newCardSection();
      var textWidget = CardService.newTextParagraph().setText(
        "Please open the email and look for the button in the top left corner. Click on it to go back and find the report button."
      );
      section.addWidget(textWidget);
      cardBuilder.addSection(section);
      return cardBuilder.build();
    }

    // All message data extracted from the single fetch above
    var sender = mailMessage.getFrom();
    var to = Session.getActiveUser().getEmail();
    const domainNameTo = to.split("@")[1];
    const senderDomain = parseDomainFromSender(sender);
    console.log('handleStep1|context|', { to, domainNameTo, sender, senderDomain });

    const message_google = mailMessage.getId();
    const messageIdOrg = mailMessage.getHeader("Message-ID");
    const StatusMessage = messageIdOrg.split("@")[0].replace('<', '');
    console.log('handleStep1|messageIds|', { message_google, StatusMessage });

    const awsRegion = await region(domainNameTo);
    const reg = awsRegion.aws_region;
    console.log('handleStep1|region|', { reg });
    await callErrorReportingApi("Region " + reg, "none");

    try {
      let isVerifiedDomain = false;
      let campaignVersion = null;
      let isV2Campaign = false;
      let verifyFailed = false;
      let verifyError = null;
      let fallbackFailed = false;
      try {
        const verifyResponse = await verifyDomain(message_google, StatusMessage, reg, to);
        isVerifiedDomain = verifyResponse.messageExists;
        campaignVersion = verifyResponse.campaignVersion;
        isV2Campaign = verifyResponse.isV2Campaign === true;
        console.log('handleStep1|verifyResponse|', { isVerifiedDomain, campaignVersion, isV2Campaign });
        await callErrorReportingApi("Is Verified Domain " + isVerifiedDomain + " campaignVersion " + campaignVersion + " isV2Campaign " + isV2Campaign, "none");
      } catch (e) {
        verifyFailed = true;
        verifyError = e;
        await callErrorReportingApi(e.stack, bodyHtml);
      }

      if (verifyFailed) {
        try {
          const fallbackResponse = await callCampaignVersionApi(StatusMessage, reg);
          campaignVersion = fallbackResponse.campaignVersion;
          isV2Campaign = fallbackResponse.isV2Campaign === true;
          // /campaignversion answers messageExists too (GetCampaignVersion.ts) and this
          // used to throw it away, so a transient failure of /admindomainsgoogle turned
          // a known simulation into a real-threat report.
          if (fallbackResponse.messageExists === true) {
            isVerifiedDomain = true;
          }
          console.log('handleStep1|campaignVersionFallback|', { isVerifiedDomain, campaignVersion, isV2Campaign });
        } catch (fallbackErr) {
          fallbackFailed = true;
          console.error('handleStep1|campaignVersionFallback|failed|', fallbackErr.message);
          await callErrorReportingApi(fallbackErr.stack, bodyHtml);
        }
      }

      // Both lookups are down, so we know nothing about this email. The fallback shares a host
      // with /admindomainsgoogle, so an admin block takes out both at once. Stop here: falling
      // through leaves isVerifiedDomain false, which reads as "not a simulation" and reports a
      // campaign send as a real threat.
      if (verifyFailed && fallbackFailed) {
        const blocked = isUrlFetchBlocked(verifyError);
        console.log('handleStep1|verification unavailable|not reporting|', { blocked });
        await callErrorReportingApi(
          "VERIFY_UNAVAILABLE blocked=" + blocked + " step=1 messageId=" + StatusMessage,
          "none"
        );
        return buildVerificationUnavailableCard(blocked);
      }

      const linkurl = foundReportUrl(bodyHtml);
      console.log('handleStep1|linkurl|', { linkurl });

      var encodedMessageId = encodeURIComponent(StatusMessage);

      // The API answer decides first, and the campaign version decides where it goes:
      // messageExists + v2 -> training portal, messageExists + v1 -> cybernut-k12 portal.
      // Previously the v2 branch was tested BEFORE messageExists, so a campaign that
      // answered campaignVersion "v2" with isV2Campaign false silently landed on the v1
      // portal. Keying off the API answer first makes the routing explicit.
      if (isVerifiedDomain === true) {
        var redirectUrl = (campaignVersion === "v2" && isV2Campaign)
          ? `https://training.cybernut.com/report?messageid=${encodedMessageId}&region=${reg}`
          : `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${reg ? reg : "us-east-1"}`;
        console.log('handleStep1|verified campaign|redirecting|', { campaignVersion, isV2Campaign, redirectUrl });
        return CardService.newActionResponseBuilder()
          .setOpenLink(CardService.newOpenLink().setUrl(redirectUrl))
          .build();
      } else if (cybernutDomains(senderDomain) || linkurl === true) {
        // Not verified by the API, but a local signal says simulation. The version is
        // unknown here, so use the v1 portal, which handles both.
        console.log('handleStep1|local signal|redirecting to portal|', { senderDomain, linkurl });
        var redirectUrl = `https://www.cybernut-k12.com/report?messageid=${encodedMessageId}&region=${reg ? reg : "us-east-1"}`;
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
                  CardService.newAction().setFunctionName("openLearnAddonLink")
                )
            )
          );
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
    await callErrorReportingApi("Aws region " + reg, "none");

    // Fetch attachments once — reused in log and payload
    const attachmentIds = getAttachmentIds(messageId);
    console.log('handleStep2|attachments|', { sourceId: messageId, attachmentIds });

    const { adminUrl, serviceUrl } = getRegionUrls(reg);
    console.log('handleStep2|urls|', { adminUrl, serviceUrl });

    let suspiciousEmailResponse = await getDomainOrFallback(domainNameTo, adminUrl, reg);
    console.log('handleStep2|suspiciousEmailResponse|', { suspiciousEmailResponse });
    await callErrorReportingApi(
      "forward suspicious email " + suspiciousEmailResponse.FORWARD_SUSPICIOUS_EMAIL,
      "none"
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

    // Verify BEFORE reporting. This used to run *after* EventDispatcherApi, and only
    // ever read trashEmail - so a simulation that slipped past handleStep1 was filed as
    // a real threat and then deleted from the user's inbox, with the one call that could
    // have prevented it happening too late to matter.
    let shouldTrash = false;
    let verifiedSimulation = false;
    let isV2Simulation = false;
    let verifyFailed = false;
    let verifyError = null;
    let fallbackFailed = false;
    try {
      const verifyResponse = await verifyDomain(message_google, StatusMessage, reg, to);
      shouldTrash = verifyResponse.trashEmail === true;
      verifiedSimulation = verifyResponse.messageExists === true;
      isV2Simulation =
        verifyResponse.campaignVersion === "v2" && verifyResponse.isV2Campaign === true;
      console.log('handleStep2|verifyResponse|', { verifiedSimulation, isV2Simulation, shouldTrash });
      await callErrorReportingApi(
        "Verify messageExists " + verifiedSimulation + " trashEmail " + shouldTrash,
        "none"
      );
    } catch (e) {
      // Verification is unavailable. Never trash on this path - we cannot tell a simulation
      // from real mail here, and deleting the user's email is the one step we cannot undo.
      verifyFailed = true;
      verifyError = e;
      shouldTrash = false;
      await callErrorReportingApi(e.stack, bodyHtml);
    }

    // Same second opinion handleStep1 gets. /campaignversion answers messageExists too, so a
    // one-off failure of /admindomainsgoogle should not cost us the recognition.
    if (verifyFailed) {
      try {
        const fallbackResponse = await callCampaignVersionApi(StatusMessage, reg);
        verifiedSimulation = fallbackResponse.messageExists === true;
        isV2Simulation =
          fallbackResponse.campaignVersion === "v2" && fallbackResponse.isV2Campaign === true;
        console.log('handleStep2|campaignVersionFallback|', { verifiedSimulation, isV2Simulation });
      } catch (fallbackErr) {
        fallbackFailed = true;
        console.error('handleStep2|campaignVersionFallback|failed|', fallbackErr.message);
        await callErrorReportingApi(fallbackErr.stack, bodyHtml);
      }
    }

    // Nothing left to ask. Stop before EventDispatcherApi - reaching it with an unverified
    // email is exactly how a training send became a Reported Threats case.
    if (verifyFailed && fallbackFailed) {
      const blocked = isUrlFetchBlocked(verifyError);
      console.log('handleStep2|verification unavailable|not reporting|', { blocked });
      await callErrorReportingApi(
        "VERIFY_UNAVAILABLE blocked=" + blocked + " step=2 messageId=" + StatusMessage,
        "none"
      );
      return buildVerificationUnavailableCard(blocked);
    }

    if (verifiedSimulation) {
      // A known campaign send. Hand it to the training flow instead of Reported Threats,
      // picking the same portal handleStep1 would have picked for this campaign version.
      const encodedStatusMessage = encodeURIComponent(StatusMessage);
      const simRedirectUrl = isV2Simulation
        ? `https://training.cybernut.com/report?messageid=${encodedStatusMessage}&region=${reg}`
        : `https://www.cybernut-k12.com/report?messageid=${encodedStatusMessage}&region=${reg ? reg : "us-east-1"}`;
      console.log('handleStep2|verified simulation|skipping eventdispatcher|', { isV2Simulation, simRedirectUrl });
      await callErrorReportingApi("Skipped eventdispatcher, verified simulation", "none");
      return CardService.newActionResponseBuilder()
        .setOpenLink(CardService.newOpenLink().setUrl(simRedirectUrl))
        .build();
    }

    const EventDispatcherApiCall = await EventDispatcherApi(payload, serviceUrl, reg);
    console.log('handleStep2|EventDispatcherApiCall|', { EventDispatcherApiCall });
    await callErrorReportingApi("Event Dispatcher " + EventDispatcherApiCall, "none");

    var builder = CardService.newCardBuilder();
    var section = CardService.newCardSection()
      .setCollapsible(false)
      .setNumUncollapsibleWidgets(1)
      .addWidget(heading)
      .addWidget(CardService.newTextParagraph().setText(
        suspiciousEmailResponse.CONFIRMATION_MESSAGE
          ? suspiciousEmailResponse.CONFIRMATION_MESSAGE
          : defaultMessageForThirdStep
      ));

    if (shouldTrash === true) {
      console.log('handleStep2|moving email to trash');
      mailMessage.moveToTrash();
      section.addWidget(CardService.newTextParagraph().setText('Email moved to trash. Please refresh your Gmail view.'));
    }

    builder.addSection(section);

    const checkInbox = mailMessage.getThread().isInInbox();
    console.log('handleStep2|checkInbox|', { checkInbox });

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
          `https://training.cybernut.com/onboarding?sessionId=${generateUUID()}&region=${reg}&email=${email}&source=google_addon&tracker=demo`
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
