/**
 * Create reply draft functionality
 * Uses Graph API createReply/createReplyAll to draft replies within existing threads
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create reply draft handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateReplyDraft(args) {
  const { messageId, body, replyAll = false, cc, bcc } = args;

  if (!messageId) {
    return {
      content: [{
        type: "text",
        text: "Message ID is required. Use read-email or search-emails to find the message ID."
      }]
    };
  }

  if (!body) {
    return {
      content: [{
        type: "text",
        text: "Reply body content is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Step 1: Create the reply draft via Graph API
    // POST /me/messages/{id}/createReply or createReplyAll
    const action = replyAll ? 'createReplyAll' : 'createReply';
    const replyDraft = await callGraphAPI(
      accessToken,
      'POST',
      `me/messages/${messageId}/${action}`,
      null
    );

    // Step 2: Update the draft with the reply body and optional recipients
    const updatePayload = {
      body: {
        contentType: /<[a-z][\s\S]*>/i.test(body) ? 'html' : 'text',
        content: body
      }
    };

    // Add CC if provided
    if (cc) {
      updatePayload.ccRecipients = cc.split(',').map(email => ({
        emailAddress: { address: email.trim() }
      }));
    }

    // Add BCC if provided
    if (bcc) {
      updatePayload.bccRecipients = bcc.split(',').map(email => ({
        emailAddress: { address: email.trim() }
      }));
    }

    const updatedDraft = await callGraphAPI(
      accessToken,
      'PATCH',
      `me/messages/${replyDraft.id}`,
      updatePayload
    );

    const toList = updatedDraft.toRecipients?.map(r => r.emailAddress.address).join(', ') || '(unknown)';

    return {
      content: [{
        type: "text",
        text: `Reply draft created successfully!\n\nDraft ID: ${updatedDraft.id}\nSubject: ${updatedDraft.subject}\nTo: ${toList}\nType: ${replyAll ? 'Reply All' : 'Reply'}\n\nThe reply draft is in your Drafts folder. Open it in Outlook to review and send.`
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
        text: `Error creating reply draft: ${error.message}`
      }]
    };
  }
}

module.exports = handleCreateReplyDraft;
