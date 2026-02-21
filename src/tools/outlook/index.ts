import type { Tool, ToolResultObject } from '@github/copilot-sdk';

function fail(e: unknown): ToolResultObject {
  const msg = e instanceof Error ? e.message : String(e);
  return { textResultForLlm: msg, resultType: 'failure', error: msg, toolTelemetry: {} };
}

function getMailbox(): Office.Mailbox {
  if (typeof Office === 'undefined' || !Office.context?.mailbox) {
    throw new Error('Outlook mailbox API is not available.');
  }
  return Office.context.mailbox;
}

const getMailItem: Tool = {
  name: 'get_mail_item',
  description:
    'Get an overview of the current mail item: subject, sender, recipients (to/cc/bcc), date, and item type (read or compose).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: async (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const itemType = item.itemType;
      const isRead = itemType === (Office.MailboxEnums.ItemType.Message as string);

      const lines: string[] = [`Mail Item Overview`, `${'='.repeat(40)}`];
      lines.push(`Type: ${String(itemType)}`);

      if (item.subject) {
        if (typeof item.subject === 'string') {
          lines.push(`Subject: ${item.subject}`);
        } else {
          const subject = await new Promise<string>((resolve, reject) => {
            (item.subject as Office.Subject).getAsync(result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
              else reject(new Error(result.error?.message ?? 'Failed to get subject'));
            });
          });
          lines.push(`Subject: ${subject}`);
        }
      }

      if (isRead && item.from) {
        const from = item.from as Office.EmailAddressDetails;
        lines.push(`From: ${from.displayName} <${from.emailAddress}>`);
      }

      if (item.to) {
        if (Array.isArray(item.to)) {
          const recipients = (item.to as Office.EmailAddressDetails[])
            .map(r => `${r.displayName} <${r.emailAddress}>`)
            .join(', ');
          lines.push(`To: ${recipients}`);
        } else {
          const recipients = await new Promise<Office.EmailAddressDetails[]>((resolve, reject) => {
            (item.to as Office.Recipients).getAsync(result => {
              if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
              else reject(new Error(result.error?.message ?? 'Failed to get recipients'));
            });
          });
          lines.push(
            `To: ${recipients.map(r => `${r.displayName} <${r.emailAddress}>`).join(', ')}`
          );
        }
      }

      if (isRead && item.dateTimeCreated) {
        lines.push(`Date: ${String(item.dateTimeCreated)}`);
      }

      lines.push(`Mode: ${isRead ? 'read' : 'compose'}`);

      return lines.join('\n');
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailBody: Tool = {
  name: 'get_mail_body',
  description: 'Get the full body content of the current mail item as HTML or plain text.',
  parameters: {
    type: 'object',
    properties: {
      format: {
        type: 'string',
        enum: ['html', 'text'],
        description: 'Body format to retrieve. Defaults to "text".',
      },
    },
    required: [],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const format = a.format === 'html' ? Office.CoercionType.Html : Office.CoercionType.Text;

      const body = await new Promise<string>((resolve, reject) => {
        item.body.getAsync(format, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error(result.error?.message ?? 'Failed to get body'));
        });
      });

      return body;
    } catch (e) {
      return fail(e);
    }
  },
};

const getMailAttachments: Tool = {
  name: 'get_mail_attachments',
  description:
    'List all attachments on the current mail item with name, size, and content type. Does not download attachment content.',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      if (!item.attachments || item.attachments.length === 0) {
        return Promise.resolve('No attachments on this mail item.');
      }

      const attachments = item.attachments;
      const lines = [`Attachments (${String(attachments.length)})`, `${'='.repeat(40)}`];
      for (const att of attachments) {
        lines.push(`- ${att.name} (${att.contentType}, ${String(att.size)} bytes)`);
      }
      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const setMailBody: Tool = {
  name: 'set_mail_body',
  description:
    'Set or replace the body content of the current mail item. Only works in compose mode. Supports HTML or plain text.',
  parameters: {
    type: 'object',
    properties: {
      content: {
        type: 'string',
        description: 'The body content to set.',
      },
      format: {
        type: 'string',
        enum: ['html', 'text'],
        description: 'Content format. Defaults to "html".',
      },
    },
    required: ['content'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const content = a.content ?? '';
      const format = a.format === 'text' ? Office.CoercionType.Text : Office.CoercionType.Html;

      await new Promise<void>((resolve, reject) => {
        item.body.setAsync(content, { coercionType: format }, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to set body'));
        });
      });

      return 'Mail body updated successfully.';
    } catch (e) {
      return fail(e);
    }
  },
};

const setMailSubject: Tool = {
  name: 'set_mail_subject',
  description: 'Set the subject of the current mail item. Only works in compose mode.',
  parameters: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'The subject text to set.',
      },
    },
    required: ['subject'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as Record<string, string | undefined>;
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');
      if (!item.subject || typeof item.subject === 'string') {
        throw new Error('Cannot set subject — item is in read mode.');
      }

      const subject = a.subject ?? '';
      await new Promise<void>((resolve, reject) => {
        (item.subject as Office.Subject).setAsync(subject, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to set subject'));
        });
      });

      return `Subject set to: "${subject}"`;
    } catch (e) {
      return fail(e);
    }
  },
};

const addMailRecipient: Tool = {
  name: 'add_mail_recipient',
  description:
    'Add one or more recipients to the current mail item. Only works in compose mode. Specify the field (to, cc, or bcc).',
  parameters: {
    type: 'object',
    properties: {
      field: {
        type: 'string',
        enum: ['to', 'cc', 'bcc'],
        description: 'Recipient field to add to.',
      },
      recipients: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            displayName: { type: 'string', description: 'Display name of the recipient.' },
            emailAddress: { type: 'string', description: 'Email address of the recipient.' },
          },
          required: ['emailAddress'],
        },
        description: 'Array of recipients to add.',
      },
    },
    required: ['field', 'recipients'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as {
        field?: string;
        recipients?: { displayName?: string; emailAddress: string }[];
      };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const field = a.field ?? 'to';
      const recipients = a.recipients;
      if (!recipients || recipients.length === 0) throw new Error('No recipients provided.');

      const recipientField = item[field as 'to' | 'cc' | 'bcc'];
      if (!recipientField || Array.isArray(recipientField)) {
        throw new Error(
          `Cannot add recipients — item is in read mode or field "${field}" is not available.`
        );
      }

      const formatted = recipients.map(r => ({
        displayName: r.displayName ?? r.emailAddress,
        emailAddress: r.emailAddress,
      }));

      await new Promise<void>((resolve, reject) => {
        recipientField.addAsync(formatted, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(result.error?.message ?? 'Failed to add recipients'));
        });
      });

      return `Added ${String(formatted.length)} recipient(s) to ${field}.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const replyToMail: Tool = {
  name: 'reply_to_mail',
  description:
    'Create a reply to the current mail item with the given HTML content. Only works in read mode. Use replyAll to reply to all recipients.',
  parameters: {
    type: 'object',
    properties: {
      htmlBody: {
        type: 'string',
        description: 'HTML content for the reply body.',
      },
      replyAll: {
        type: 'boolean',
        description: 'If true, reply to all recipients. Defaults to false.',
      },
    },
    required: ['htmlBody'],
  },
  handler: async (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { htmlBody?: string; replyAll?: boolean };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const htmlBody = a.htmlBody ?? '';
      const replyAll = Boolean(a.replyAll);

      await new Promise<void>((resolve, reject) => {
        const displayCall = replyAll ? item.displayReplyAllForm : item.displayReplyForm;
        if (!displayCall) {
          reject(new Error('Reply is not supported on this item.'));
          return;
        }
        displayCall.call(item, {
          htmlBody,
          callback: (result: Office.AsyncResult<void>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error(result.error?.message ?? 'Failed to create reply'));
          },
        });
      });

      return `Reply${replyAll ? ' all' : ''} form opened with provided content.`;
    } catch (e) {
      return fail(e);
    }
  },
};

const forwardMail: Tool = {
  name: 'forward_mail',
  description:
    'Open a forward form for the current mail item with optional HTML content pre-filled. Only works in read mode.',
  parameters: {
    type: 'object',
    properties: {
      htmlBody: {
        type: 'string',
        description: 'Optional HTML content to prepend to the forwarded message.',
      },
    },
    required: [],
  },
  handler: (args: unknown): Promise<ToolResultObject | string> => {
    try {
      const a = args as { htmlBody?: string };
      const mailbox = getMailbox();
      const item = mailbox.item;
      if (!item) throw new Error('No mail item is currently open.');

      const htmlBody = a.htmlBody;

      // displayForwardForm is available on MessageRead but not typed in all @types/office-js versions
      const itemRead = item as unknown as {
        displayForwardForm?: (formData: { htmlBody?: string }) => void;
      };
      if (!itemRead.displayForwardForm) {
        throw new Error('Forward is not supported on this item.');
      }

      itemRead.displayForwardForm(htmlBody ? { htmlBody } : {});
      return Promise.resolve('Forward form opened.');
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

const getUserProfile: Tool = {
  name: 'get_user_profile',
  description: 'Get the current mailbox user profile information (name, email, time zone).',
  parameters: { type: 'object', properties: {}, required: [] },
  handler: (): Promise<ToolResultObject | string> => {
    try {
      const mailbox = getMailbox();
      const userProfile = mailbox.userProfile;

      const lines = [
        `User Profile`,
        `${'='.repeat(40)}`,
        `Name: ${userProfile.displayName}`,
        `Email: ${userProfile.emailAddress}`,
        `Time Zone: ${userProfile.timeZone}`,
      ];
      return Promise.resolve(lines.join('\n'));
    } catch (e) {
      return Promise.resolve(fail(e));
    }
  },
};

export const outlookTools: Tool[] = [
  getMailItem,
  getMailBody,
  getMailAttachments,
  setMailBody,
  setMailSubject,
  addMailRecipient,
  replyToMail,
  forwardMail,
  getUserProfile,
];
