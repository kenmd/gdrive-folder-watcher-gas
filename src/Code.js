// Official documents
// standalone overview: https://developers.google.com/apps-script/guides/standalone
// add-on sample code: https://developers.google.com/workspace/add-ons/translate-addon-sample
// card reference: https://developers.google.com/apps-script/reference/card-service
// test and publish add-on https://developers.google.com/workspace/add-ons/how-tos/publish-add-on-overview
// https://developers.google.com/workspace/add-ons/drive/building-drive-interfaces
// Make sure addOns.drive exists in appsscript.json to see this add-on in Gdrive page

// Most of unofficial info on the internet is outdated but these might be helpful
// https://www.pre-practice.net/2019/01/blog-post_13.html
// https://news.mynavi.jp/article/gas-13/

// assuming traditional style (legacy-ish) slack webhook url
const defaultWebhookUrl = "https://hooks.slack.com/services/AAA/BBB/CCC";

function onHomepage(e) {
  const config = getConfig();
  const isTimerOn = getIsTimerOn();

  return createHomeCard(config.folderId, config.slackUrl, isTimerOn);
}

function createHomeCard(folderId, slackUrl, isTimerOn) {
  Logger.log(`createHomeCard(${folderId}, ${slackUrl}, ${isTimerOn})`);

  let folderName = "";

  // validation
  try {
    const folder = DriveApp.getFolderById(folderId);
    folderName = folder.getName();
  }
  catch (e) {
    Logger.log(e);
  }

  const builder = CardService.newCardBuilder();

  const configSection = CardService.newCardSection()
    .addWidget(CardService.newTextInput()
      .setTitle('Enter Gdrive Folder ID')
      .setFieldName('folderId')
      .setValue(folderId || "")
    )
    .addWidget(CardService.newDecoratedText()
      .setTopLabel("Folder Name")
      .setText(folderName)
      .setButton(CardService.newImageButton()
        .setIcon(CardService.Icon.VIDEO_PLAY)
        .setOpenLink(
          CardService.newOpenLink()
          .setUrl(`https://drive.google.com/drive/folders/${folderId}`)
        )
      )
    )
    .addWidget(CardService.newTextInput()
      .setTitle('Enter Slack Webhook URL')
      .setFieldName('slackUrl')
      .setValue(slackUrl || defaultWebhookUrl)
    )
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Save Config')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction().setFunctionName('saveConfig')
        )
      )
    );

  builder.addSection(configSection);

  const triggerSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setTopLabel('For Testing')
      .setText('Check the folder and notify if new files are found')
      .setWrapText(true)
      .setButton(CardService.newTextButton()
        .setText('Try Now')
        .setTextButtonStyle(CardService.TextButtonStyle.TEXT)
        .setOnClickAction(
          CardService.newAction().setFunctionName('checkFolder')
        )
      )
    )
    .addWidget(CardService.newDecoratedText()
      .setTopLabel("Check the folder periodically")
      .setText("Timer Trigger ON/OFF")
      .setSwitchControl(CardService.newSwitch()
        .setFieldName("isTimerOn")
        .setSelected(isTimerOn === "true")
        .setValue("true")
        .setOnChangeAction(CardService.newAction()
            .setFunctionName("saveIsTimerOn")
        )
      )
    )
    .addWidget(CardService.newDecoratedText()
      .setTopLabel('Detail configuration')
      .setText('Open App Script console to update the timer (default: every 1 hour)')
      .setWrapText(true)
      .setButton(CardService.newImageButton()
        .setIcon(CardService.Icon.FLIGHT_DEPARTURE)
        .setOpenLink(
          CardService.newOpenLink()
          .setUrl("https://script.google.com/")
        )
      )
    );

  builder.addSection(triggerSection);

  return builder.build();
}

function checkFolder(e) {
  const config = getConfig();
  const folderId = config.folderId;
  const slackUrl = config.slackUrl;

  // validation
  const folder = DriveApp.getFolderById(folderId);
  Logger.log(`checkFolder: ${folder.getName()} ${folderId} ${slackUrl}`);

  // for testing
  // let cutOffTime = (new Date(now.getTime()-24*60*60*1000*280)).toISOString();

  const newCutOffTime = findNewCutOffTime();
  const prevCutOffTime = findPrevCutOffTime(newCutOffTime);
  const files = findNewFiles(folder, prevCutOffTime, newCutOffTime);

  if (files.hasNext()) {
    sendSlack(folder, files, slackUrl);
  } else {
    Logger.log("New file not found");
  }

  setCutOffTime(newCutOffTime);

  if ("triggerUid" in e) {
    Logger.log(`Started from trigger ${e.triggerUid}`);
    return
  }

  const isTimerOn = getIsTimerOn();
  return createHomeCard(folderId, slackUrl, isTimerOn);
}

function findNewCutOffTime() {
  const now = new Date();
  // check up to 10 sec before to make sure all files are copied on Gdrive
  now.setSeconds(now.getSeconds() - 10);
  const newCutOffTime = now.toISOString();

  return newCutOffTime;
}

function findPrevCutOffTime(defaultCutOffTime) {
  let prevCutOffTime = getCutOffTime() || defaultCutOffTime;

  // check maximum 1 day to prevent to find too many files
  const oneDayMilSec = 24*60*60*1000;
  const intervalDays = ((new Date()).getTime() - (new Date(prevCutOffTime)).getTime())/oneDayMilSec;

  if (intervalDays > 1) {
    Logger.log(`Override prev CutOffTime more than 1 day ago: ${intervalDays}`);
    prevCutOffTime = defaultCutOffTime;
  }

  return prevCutOffTime;
}

function findNewFiles(folder, prevCutOffTime, newCutOffTime) {
  Logger.log(`findNewFiles in ${folder.getName()}`);

  const query = `modifiedDate > "${prevCutOffTime}" and modifiedDate <= "${newCutOffTime}"`;
  Logger.log("query: " + query);

  const files = folder.searchFiles(query);

  return files;
}

function sendSlack(folder, files, slackUrl) {
  let message = `:ledger: ${folder.getName()}`;

  const fileIds = [];

  while (files.hasNext()) {
    const file = files.next();
    Logger.log(`id: ${file.getId()}, name: ${file.getName()}, created: ${file.getDateCreated().toISOString()}`);

    fileIds.push(file.getId());
    message += ` :page_facing_up: ${file.getName()}`;
  }

  Logger.log(`Found number of files: ${fileIds.length}`);

  const buttonUrl = `https://drive.google.com/drive/folders/${folder.getId()}`;
  Logger.log(`Slack ${message} ${buttonUrl}`);

  const res = postMessage(slackUrl, message, buttonUrl);
  Logger.log(res);  // ok
}

// --------------------------------------------------
// Slack Utilities
// https://dev.classmethod.jp/articles/google-apps-script-slack-api-launch-01/
// --------------------------------------------------

function postMessage(slackUrl, text, buttonUrl) {
  var payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": text
        },
        "accessory": {
          "type": "button",
          "text": {
            "type": "plain_text",
            "text": "Open :open_file_folder:",
            "emoji": true
          },
          "url": buttonUrl
        }
      }
    ]
  };
  var options = {
    "method": "POST",
    "payload": JSON.stringify(payload),
  };

  return UrlFetchApp.fetch(slackUrl, options);
}

// --------------------------------------------------
// OnClick Handlers
// --------------------------------------------------

function saveConfig(e) {
  const folderId = e.formInput.folderId;
  const slackUrl = e.formInput.slackUrl;

  // validation
  const folder = DriveApp.getFolderById(folderId);
  Logger.log(`Config Folder Name: ${folder.getName()}`);

  setConfig(folderId, slackUrl);

  const isTimerOn = getIsTimerOn();
  return createHomeCard(folderId, slackUrl, isTimerOn);
}

function saveIsTimerOn(e) {
  // e.formInput.isTimerOn is "true" when switch on, undefined when switch off
  // isTimerOn is always string "true" or "false"
  const isTimerOn = e.formInput.isTimerOn === "true" ? "true" : "false";

  Logger.log(`handleTrigger isTimerOn: ${isTimerOn}`);
  handleTrigger(isTimerOn);

  setIsTimerOn(isTimerOn);

  const config = getConfig();
  return createHomeCard(config.folderId, config.slackUrl, isTimerOn);
}

function handleTrigger(isTimerOn) {
  if (isTimerOn === "true") {
    const triggerId = createTimeDrivenTriggers();
    Logger.log(`Created triggerId: ${triggerId}`);
    setTriggerId(triggerId);
  } else {
    const triggerId = getTriggerId();
    deleteTrigger(triggerId);
    Logger.log(`Deleted triggerId: ${triggerId}`);
    deleteTriggerId();
  }
}

// https://developers.google.com/apps-script/guides/triggers/installable
function createTimeDrivenTriggers() {
  const trigger = ScriptApp.newTrigger('checkFolder')
      .timeBased()
      .everyHours(1)
      .create();

  return trigger.getUniqueId();
}

function deleteTrigger(triggerId) {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}

// --------------------------------------------------
// PropertiesService Utilities
// --------------------------------------------------

function getConfig() {
  const props = PropertiesService.getUserProperties();
  return {
    "folderId": props.getProperty("folderId"),
    "slackUrl": props.getProperty("slackUrl"),
  };
}

function setConfig(folderId, slackUrl) {
  const props = PropertiesService.getUserProperties();
  props.setProperty("folderId", folderId);
  props.setProperty("slackUrl", slackUrl);
  return {
    "folderId": folderId,
    "slackUrl": slackUrl,
  };
}

function getIsTimerOn() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty("isTimerOn");
}

function setIsTimerOn(isTimerOn) {
  const props = PropertiesService.getUserProperties();
  props.setProperty("isTimerOn", isTimerOn);
}

function getTriggerId() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty("triggerId");
}

function setTriggerId(triggerId) {
  const props = PropertiesService.getUserProperties();
  props.setProperty("triggerId", triggerId);
}

function deleteTriggerId() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty("triggerId");
}

function getCutOffTime() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty("cutOffTime");
}

function setCutOffTime(cutOffTime) {
  const props = PropertiesService.getUserProperties();
  props.setProperty("cutOffTime", cutOffTime);
}
