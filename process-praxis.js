const { ensureAuthenticated } = require('./auth');

async function getEmailDetails(emailId) {
  const token = await ensureAuthenticated();
  const url = `https://graph.microsoft.com/v1.0/me/messages/${emailId}`;
  const response = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  return response.json();
}

async function getPraxisEmails() {
  const token = await ensureAuthenticated();
  const praxisFolderId = 'AAMkAGE5MjJhNjdiLWFiMTctNDIxYy04ZGE0LWNlOGExZjUxNGE0YgAuAAAAAABo2vjQMM0yTrgOpJbxV1KeAQDJZ1BLtR7DQJOH_bEghE0MAADl5-U7AAA=';

  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${praxisFolderId}/messages?$filter=isRead eq false&$orderby=receivedDateTime asc&$top=20&$select=id,subject,from,receivedDateTime,body`;

  const response = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  return response.json();
}

function extractOriginalSender(bodyContent) {
  // Extract original sender from forwarded email
  const fromMatch = bodyContent.match(/From:\s*([^<\n]+)<([^>]+)>/i) ||
                    bodyContent.match(/From:\s*([^\n<]+@[^\n\s]+)/i);

  if (fromMatch) {
    if (fromMatch[2]) {
      return { name: fromMatch[1].trim(), email: fromMatch[2].trim().toLowerCase() };
    } else {
      return { name: '', email: fromMatch[1].trim().toLowerCase() };
    }
  }
  return null;
}

function extractResponseIntent(bodyContent) {
  const lower = bodyContent.toLowerCase();

  // Check for decline signals
  if (lower.includes('not investing') || lower.includes('not currently') ||
      lower.includes('not active') || lower.includes('pass on') ||
      lower.includes('not a fit') || lower.includes('decline') ||
      lower.includes("don't invest") || lower.includes('not allocating') ||
      lower.includes('no interest') || lower.includes('not interested')) {
    return 'DECLINED';
  }

  // Check for interest signals
  if (lower.includes('interested') || lower.includes('happy to') ||
      lower.includes('schedule') || lower.includes('set up a call') ||
      lower.includes('learn more') || lower.includes('send deck') ||
      lower.includes('send materials') || lower.includes('would love')) {
    return 'INTERESTED';
  }

  // Check for more info request
  if (lower.includes('more information') || lower.includes('tell me more') ||
      lower.includes('can you share') || lower.includes('send over')) {
    return 'INFO_REQUEST';
  }

  return 'UNCLEAR';
}

async function main() {
  const emails = await getPraxisEmails();

  console.log(JSON.stringify(emails.value.map(e => {
    const sender = extractOriginalSender(e.body.content);
    const intent = extractResponseIntent(e.body.content);

    // Extract org name from subject
    const subjectMatch = e.subject.match(/Fwd:\s*(?:\[EXTERNAL\]\s*)?(?:RE:\s*)?(?:Re:\s*)?(.+?)(?:\s*<-->|\s*<->|\s*\/|\s*x\s+|\s*-\s+|\s*&\s+)?\s*CrossLayer/i);
    const orgName = subjectMatch ? subjectMatch[1].trim() : e.subject;

    return {
      id: e.id,
      date: e.receivedDateTime,
      subject: e.subject,
      orgName: orgName,
      originalSender: sender,
      intent: intent,
      bodyPreview: e.body.content.substring(0, 500).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()
    };
  }), null, 2));
}

main().catch(console.error);
