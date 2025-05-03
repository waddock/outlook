Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("applyTags").onclick = applyMetadata;
      Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, validateBeforeSend);
    }
  });
  
  async function applyMetadata() {
    const projectCode = document.getElementById('projectCode').value.trim();
    const emailType = document.getElementById('emailType').value.trim();
  
    if (!projectCode || !emailType) {
      alert('Both Project Code and Email Type are required!');
      return;
    }
  
    const subjectPrefix = `[${projectCode}][${emailType}] `;
    
    Office.context.mailbox.item.subject.getAsync({ asyncContext: subjectPrefix }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const currentSubject = asyncResult.value || '';
        Office.context.mailbox.item.subject.setAsync(subjectPrefix + currentSubject);
      }
    });
  }
  
  // Validate before sending the email
  function validateBeforeSend(event) {
    Office.context.mailbox.item.subject.getAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const subject = asyncResult.value || '';
        
        // Simple validation: subject must include both tags
        const hasProjectCode = subject.includes('[PRJ-');
        const hasEmailType = subject.match(/\[.*?\]/g)?.length >= 2;
  
        if (!hasProjectCode || !hasEmailType) {
          event.completed({ allowEvent: false, errorMessage: 'Project Code and Email Type are required before sending.' });
        } else {
          event.completed({ allowEvent: true });
        }
      } else {
        event.completed({ allowEvent: false, errorMessage: 'Validation failed.' });
      }
    });
  }  