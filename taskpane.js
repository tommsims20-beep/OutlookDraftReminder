Office.onReady(() => {
  const statusEl = document.getElementById("status");
  const enableBtn = document.getElementById("enableBtn");
  const checkNowBtn = document.getElementById("checkNowBtn");
  const draftListEl = document.getElementById("draftList");

  refreshStatus();

  enableBtn.addEventListener("click", async () => {
    if (typeof Notification === "undefined") {
      statusEl.textContent = "Notifications aren't supported in this Outlook client.";
      return;
    }
    const result = await Notification.requestPermission();
    refreshStatus();
    if (result === "granted") {
      new Notification("Draft Reminder enabled", {
        body: "You'll be notified here of drafts unsent for over an hour.",
      });
    }
  });

  checkNowBtn.addEventListener("click", () => {
    draftListEl.innerHTML = "<li>Checking…</li>";
    checkDraftsNow((subjects, errorMessage) => {
      if (errorMessage) {
        draftListEl.innerHTML = `<li>Error: ${errorMessage}</li>`;
        return;
      }
      if (subjects.length === 0) {
        draftListEl.innerHTML = "<li>No stale drafts. Nice.</li>";
        return;
      }
      draftListEl.innerHTML = subjects.map((s) => `<li>${escapeHtml(s)}</li>`).join("");
    });
  });

  function refreshStatus() {
    if (typeof Notification === "undefined") {
      statusEl.textContent = "Notifications: not supported here";
      enableBtn.disabled = true;
      return;
    }
    statusEl.textContent = `Notifications: ${Notification.permission}`;
    enableBtn.disabled = Notification.permission === "granted";
  }
});

function checkDraftsNow(callback) {
  const cutoffIso = new Date(Date.now() - 60 * 60 * 1000).toISOString();

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
      callback([], asyncResult.error.message);
      return;
    }
    const doc = new DOMParser().parseFromString(asyncResult.value, "text/xml");
    const ns = "http://schemas.microsoft.com/exchange/services/2006/types";
    const subjectNodes = doc.getElementsByTagNameNS(ns, "Subject");
    const subjects = [];
    for (let i = 0; i < subjectNodes.length; i++) {
      subjects.push(subjectNodes[i].textContent || "(no subject)");
    }
    callback(subjects, null);
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
