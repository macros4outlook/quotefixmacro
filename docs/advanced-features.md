---
nav_order: 3
---
# Advanced usage

Configuration is done via constants in the QuoteFix code (see below for a storage in the registry)

1. Start the VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd>)
2. Open the module "QuoteFixMacro"
3. Scroll down to the block "Configuration constants"

## Configure the template inside the code

The variable `QUOTING_TEMPLATE` can be used to store the quoting template.
Thus, the Outlook configuration can be left untouched.

If this is not enabled, one has to configure Outlook differently:

Tools > Options > Mail Format > Signatures...

* Create a signature that is only used for reply and forward. You have to insert at least `%Q` to get the quoted original mail.
* Assign this signature to every mail account you want to use.

## English replies

The variable `QUOTING_TEMPLATE_EN` can be used to store en English quoting template.
In case `USE_QUOTING_TEMPLATE` is `True` (Default since 2.0) and `FixedReplyAllEnglish()` is called, that template is used.

## Auto conversion to plain format

By setting `CONVERT_TO_PLAIN` to `True` (Default since 2.0), HTML mails are automatically converted to text mails.

If this is not configured, one has to configure Outlook differently:

Tools > Options > Mail Format

* Message format: Plain Text
* Note: this is not necessary, if `CONVERT_TO_PLAIN` is set to `True`.

## Read all email as text

QuoteFixMacro requires plain text to work. One can either read all emails as plain text from the beginning or set `CONVERT_TO_PLAIN` is set to `True`.
In case all texts should be read as plain text, see Microsoft [KB 831607](https://support.microsoft.com/en-us/office/change-the-message-format-to-html-rich-text-format-or-plain-text-338a389d-11da-47fe-b693-cf41f792fefa?ui=en-us&rs=en-us&ad=us) for an explanation how to turn on this feature. For Outlook 2010 and later (also described at ["Read email messages in plain text"](https://support.microsoft.com/en-us/office/read-email-messages-in-plain-text-16dfe54a-fadc-4261-b2ce-19ad072ed7e3?ui=en-US&rs=en-US&ad=US)): File > Options > Security Center > Options for the Security Center > E-Mail Security > "Read as Plain Text" > `[X]` Read all standard mail in plain text, `[X]` Read all digitally signed mail in plain text".

## Condense Headers

With `CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS`, one condenses reply/forwarding headers added by outlook so that the email gets even shorter
The format of the condensed header is configured at `CONDENSED_HEADER_FORMAT`

One can also condense the first header only `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER`.

### Date format

The date format used is [ISO-8601](https://xkcd.com/1179/), which is `YYYY-MM-DD`.
One can change the format in the variable `DEFAULT_DATE_FORMAT`.

## Strip sender's signature

By default, the sender's signature is removed from the reply. If you don't want this, set `STRIP_SIGNATURE` to `False`.

## SoftWrap

When enabled, this feature resizes the window in a way that the text editor wraps the text automatically after N characters.
Outlook wraps text automatically after sending it, but doesn't display the wrap when editing.
Thus, this is useful to double-check that no new line breaks are introduced by Outlook when sending an email.

One can set `USE_SOFTWRAP` to `False` to disable it.

<!-- markdownlint-disable-file MD033 -->
