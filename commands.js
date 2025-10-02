/* global Office */
Office.onReady(() => { /* ready */ });

// Referenced by manifest
function createQuote(event) {
  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Html, async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get email body:", result.error);
      notify("Could not read email body.");
      event.completed();
      return;
    }

    const htmlBody = result.value;

    try {
      const resp = await fetch(
        "https://amxmobileportalapi-cpbdh9gnanc7bces.centralus-01.azurewebsites.net/api/Quote/OutlookCreateQuote",
        {
          method: "POST",
          headers: {
            "Content-Type": "text/plain"
          },
          body: htmlBody
        }
      );

      if (!resp.ok) {
        const text = await resp.text().catch(() => "");
        console.error("API error:", resp.status, resp.statusText, text);
        notify(`API failed: ${resp.status}`);
      } else {
        notify("Quote created.");
      }
    } catch (e) {
      console.error("Network error:", e);
      notify("Network error.");
    }

    event.completed();
  });
}

function notify(message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("quoteStatus", {
      type: "informationalMessage",
      message,
      icon: "icon16",
      persistent: false
    });
  } catch { /* noop */ }
}

window.createQuote = createQuote;
