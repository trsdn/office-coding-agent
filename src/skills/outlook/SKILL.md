---
name: outlook
description: General-purpose Outlook skill for email analysis, drafting, replying, and mailbox management tasks.
license: MIT
hosts: [outlook]
---

# Outlook Default Skill

Use this as the default orchestration skill for Outlook tasks.

## Operating Loop

1. **Discover** — Read the current mail item context (subject, sender, recipients, date) first.
2. **Read** — Fetch the full email body and attachments before composing a response.
3. **Execute** — Draft, reply, or forward with focused content.
4. **Verify** — Confirm the action taken (body set, reply opened, recipients added).
5. **Summarize** — Finish with a concise plain-language summary of what was done.

## Read Mode vs Compose Mode

- **Read mode** (viewing a received email): read all properties, use `reply_to_mail` or `forward_mail` to respond.
- **Compose mode** (drafting): use `set_mail_body`, `set_mail_subject`, `add_mail_recipient` to build the email.
- Always check item context before attempting write operations.

## High-Level Tool Guidance

| Task                        | Primary Tool                                    |
| --------------------------- | ----------------------------------------------- |
| Understand current email    | `get_mail_item`                                 |
| Read email content          | `get_mail_body`                                 |
| Check attachments           | `get_mail_attachments`                          |
| Set email body (compose)    | `set_mail_body`                                 |
| Set subject (compose)       | `set_mail_subject`                              |
| Add recipients (compose)    | `add_mail_recipient`                            |
| Reply to email (read)       | `reply_to_mail`                                 |
| Forward email (read)        | `forward_mail`                                  |
| Get user info               | `get_user_profile`                              |

## Common Workflows

### Summarize an email
1. `get_mail_item` → understand context
2. `get_mail_body` → read full content
3. Provide a concise summary to the user

### Draft a reply
1. `get_mail_item` → understand who sent it and the subject
2. `get_mail_body` → read what was said
3. `reply_to_mail` → compose a contextual reply

### Analyze attachments
1. `get_mail_item` → check context
2. `get_mail_attachments` → list what's attached
3. Report findings to the user

## Always-On Defaults

- Always read the email context before any action.
- Use HTML format for rich replies, plain text for simple responses.
- Always finish with a clear summary of actions taken.

## Multi-Step Requests

Execute all requested steps in sequence where possible. If one step fails, report the failure clearly and continue independent remaining steps.
