const { ensureAuthenticated } = require('./auth');

async function listPraxisEmails() {
  const token = await ensureAuthenticated();

  // First get the praxis folder ID
  const foldersResponse = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/childFolders', {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const folders = await foldersResponse.json();
  const praxisFolder = folders.value.find(f => f.displayName.toLowerCase() === 'praxis');

  if (!praxisFolder) {
    console.log('Praxis folder not found');
    return;
  }

  console.log('Found praxis folder:', praxisFolder.id);

  // Get unread emails from praxis folder
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${praxisFolder.id}/messages?$filter=isRead eq false&$orderby=receivedDateTime asc&$top=20&$select=id,subject,from,receivedDateTime,bodyPreview`;

  const emailsResponse = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const emails = await emailsResponse.json();

  console.log('\nUnread emails in Praxis folder:', emails.value?.length || 0);
  if (emails.value) {
    emails.value.forEach((e, i) => {
      console.log('---');
      console.log((i+1) + '. ' + e.subject);
      console.log('   From: ' + e.from.emailAddress.address);
      console.log('   Date: ' + e.receivedDateTime);
      console.log('   Preview: ' + e.bodyPreview.substring(0, 150) + '...');
    });
  }
}

listPraxisEmails().catch(console.error);
