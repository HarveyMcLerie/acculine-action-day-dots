// Action Day Dots — command functions
// Completely standalone from the Sticky Notes add-in: own webhook, own logic, no shared code.

var WEBHOOK_URL = "https://acculine135.app.n8n.cloud/webhook/action-day-dots";

var DAYS = {
  Monday: "🔴",
  Tuesday: "🟠",
  Wednesday: "🟡",
  Thursday: "🟢",
  Friday: "🔵"
};

Office.onReady();

function applyDayDot(day, event) {
  var item = Office.context.mailbox.item;

  var restId = Office.context.mailbox.convertToRestId
    ? Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0)
    : item.itemId;

  var payload = {
    itemId: restId,
    currentSubject: item.subject || "",
    day: day
  };

  fetch(WEBHOOK_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  })
    .then(function (response) {
      if (!response.ok) {
        console.error("Action Day Dots: webhook returned " + response.status);
      }
      event.completed();
    })
    .catch(function (error) {
      console.error("Action Day Dots: webhook call failed", error);
      event.completed();
    });
}

function setMonday(event) { applyDayDot("Monday", event); }
function setTuesday(event) { applyDayDot("Tuesday", event); }
function setWednesday(event) { applyDayDot("Wednesday", event); }
function setThursday(event) { applyDayDot("Thursday", event); }
function setFriday(event) { applyDayDot("Friday", event); }

// Register functions for Office to call from the ribbon buttons
Office.actions = Office.actions || {};
Office.actions.associate("setMonday", setMonday);
Office.actions.associate("setTuesday", setTuesday);
Office.actions.associate("setWednesday", setWednesday);
Office.actions.associate("setThursday", setThursday);
Office.actions.associate("setFriday", setFriday);
