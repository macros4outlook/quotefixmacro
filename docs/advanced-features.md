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
In case `USE_QUOTING_TEMPLATE` is `True` and `FixedReplyAllEnglish()` is called, that template is used.

## Auto conversion to plain format

By setting `CONVERT_TO_PLAIN` to `True`, HTML mails are automatically converted to text mails.

Note that if this makes following Outlook obsolete:

Tools > Options > Mail Format

* Message format: Plain Text

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

## Use templates from the code

Instead of confuring a template in the signature setting, one can set `DEFAULT_USE_QUOTING_TEMPLATE` to `True`.
Then, QuoteFixMacro reads the signature from `DEFAULT_QUOTING_TEMPLATE_EN` for English emails and from `DEFAULT_QUOTING_TEMPLATE` for all other languages.

## Random Signature Generation

In case you want to try out the current "random signature generation", import `RandomSignature.bas`.

<!-- markdownlint-disable-file MD033 -->
