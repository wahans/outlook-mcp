/**
 * Create email draft functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create email draft handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateDraft(args) {
  const { to, cc, bcc, subject, body, importance = 'normal' } = args;

  // Validate required parameters
  if (!subject) {
    return {
      content: [{
        type: "text",
        text: "Subject is required."
      }]
    };
  }

  if (!body) {
    return {
      content: [{
        type: "text",
        text: "Body content is required."
      }]
    };
  }

  try {
    // Get access token
    const accessToken = await ensureAuthenticated();

    // Format recipients (optional for drafts)
    const toRecipients = to ? to.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    const ccRecipients = cc ? cc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    const bccRecipients = bcc ? bcc.split(',').map(email => ({
      emailAddress: { address: email.trim() }
    })) : [];

    // Prepare draft message object (different structure than sendMail)
    const draftMessage = {
      subject,
      body: {
        contentType: /<[a-z][\s\S]*>/i.test(body) ? 'html' : 'text',
        content: body
      },
      importance
    };

    // Only add recipients if provided
    if (toRecipients.length > 0) {
      draftMessage.toRecipients = toRecipients;
    }
    if (ccRecipients.length > 0) {
      draftMessage.ccRecipients = ccRecipients;
    }
    if (bccRecipients.length > 0) {
      draftMessage.bccRecipients = bccRecipients;
    }

    // Create draft via POST /me/messages
    const response = await callGraphAPI(accessToken, 'POST', 'me/messages', draftMessage);

    return {
      content: [{
        type: "text",
        text: `Draft created successfully!\n\nDraft ID: ${response.id}\nSubject: ${subject}\nRecipients: ${toRecipients.length > 0 ? toRecipients.map(r => r.emailAddress.address).join(', ') : '(none yet)'}\n\nThe draft is now in your Drafts folder and can be edited or sent from Outlook.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error creating draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateDraft;
