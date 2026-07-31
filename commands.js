/* Draft Reminder — background runtime
 *
 * Runs inside the long-lifetime runtime declared in manifest.xml.
 * Started the first time OnNewMessageCompose fires in an Outlook session
 * (see LaunchEvent in the manifest). From then on it stays resident and
 * polls the Drafts folder once an hour for as long as Outlook keeps this
 * runtime alive.
 *
 * IMPORTANT CAVEATS (read before relying on this):
 * 1. This only runs while Outlook is open. No Outlook process = no checks.
 * 2. Outlook/WebView2 can suspend or reload this runtime to save resources.
 *    There's no platform guarantee it survives indefinitely — treat this as
 *    "best effort while Outlook is running," not a guaranteed cron job.
 * 3. Notification permission must be granted once via the task pane
 *    (taskpane.js) before this background page can show OS notifications.
 *    Browsers do not allow permission prompts from hidden/background pages.
 */

const CHECK_INTERVAL_MS = 60 * 60 * 1000; // 1 hour
const STALE_THRESHOLD_MS = 60 * 60 * 1000; // consider a draft "stale" after 1 hour

let intervalHandle = null;

Office.onReady(() => {
  // no-op here; actual start is triggered by the associated launch event below
});

// Associated with the OnNewMessageCompose LaunchEvent in manifest.xml.
function onComposeStart(event) {
  startHourlyCheck();
  event.completed({ allowEvent: true });
}
Office.actions.associate("onComposeStart", onComposeStart);

function startHourlyCheck() {
  if (intervalHandle) return; // already running in this session, don't double-schedule

  checkDraftsAndNotify(); // run once immediately, then hourly
  intervalHandle = setInterval(checkDraftsAndNotify, CHECK_INTERVAL_MS);
}

function checkDraftsAndNotify() {
  const cutoffIso = new Date(Date.now() - STALE_THRESHOLD_MS).toISOString();

  const ewsRequest = `
    <?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"
                  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                  Traversal="Shallow">
          <ItemShape>
            <t:BaseShape>IdOnly</t:BaseShape>
            <t:AdditionalProperties>
              <t:FieldURI FieldURI="item:Subject" />
              <t:FieldURI FieldURI="item:DateTimeCreated" />
            </t:AdditionalProperties>
          </ItemShape>
          <Restriction>
            <t:IsLessThan>
              <t:FieldURI FieldURI="item:DateTimeCreated" />
              <t:FieldURIOrConstant>
                <t:Constant Value="${cutoffIso}" />
              </t:FieldURIOrConstant>
            </t:IsLessThan>
          </Restriction>
          <ParentFolderIds>
            <t:DistinguishedFolderId Id="drafts" />
          </ParentFolderIds>
        </FindItem>
      </soap:Body>
    </soap:Envelope>
  `.trim();

  Office.context.mailbox.makeEwsRequestAsync(ewsRequest, (asyncResult) => {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("EWS request failed:", asyncResult.error);
      return;
    }

    const subjects = parseSubjectsFromEwsResponse(asyncResult.value);
    if (subjects.length > 0) {
      notify(subjects);
    }
  });
}

function parseSubjectsFromEwsResponse(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const ns = "http://schemas.microsoft.com/exchange/services/2006/types";
  const subjectNodes = doc.getElementsByTagNameNS(ns, "Subject");
  const subjects = [];
  for (let i = 0; i < subjectNodes.length; i++) {
    subjects.push(subjectNodes[i].textContent || "(no subject)");
  }
  return subjects;
}

function notify(subjects) {
  if (typeof Notification === "undefined") {
    console.warn("Notification API unavailable in this runtime.");
    return;
  }
  if (Notification.permission !== "granted") {
    console.warn("Notification permission not granted yet — open the task pane once to enable it.");
    return;
  }

  const body =
    subjects.length === 1
      ? `1 draft has been unsent for over an hour: "${subjects[0]}"`
      : `${subjects.length} drafts have been unsent for over an hour.`;

  new Notification("Unsent draft reminder", {
    body,
    icon: "assets/icon-80.png",
  });
}
