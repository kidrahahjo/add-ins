Office.onReady(function (info) {
    // Office is ready
    if (info.hostType === Office.HostType.Outlook) {
      document.getElementById("send-text").onclick = () => tryCatch(sendSelectedText);
    }
  });
  
  async function sendSelectedText() {
    // Get a reference to the current message
    const item = Office.context.mailbox.item;
  
    // Get the selected text from the message
    const selectedText = item.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log(result.value.data);
      } else {
        console.error(result.error);
      }
    });
  }
  
  /** Default helper for invoking an action and handling errors. */
  async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    }
  }
  