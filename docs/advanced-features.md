---
nav_order: 3
---
# Advanced usage

Configuration is done via constants in the QuoteFix code (see below for a storage in the registry)

1. Start the VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd>)
2. Open the module "QuoteFixMacro"
3. Scroll down to the block "Configuration constants"

## Configure the template inside the code

With `USE_QUOTING_TEMPLATE` set to `True`, QuoteFixMacro uses the template `QUOTING_TEMPLATE` instead of the signature from Outlook (for mails in English: `QUOTING_TEMPLATE_EN`, see below).
Thus, the Outlook configuration can be left untouched.
The defaults are `DEFAULT_QUOTING_TEMPLATE` and `DEFAULT_QUOTING_TEMPLATE_EN` in the code.
To change a template, store it in the registry, e.g., with [`configs/exampleTemplateWithSignature.reg`](https://github.com/macros4outlook/quotefixmacro/blob/main/configs/exampleTemplateWithSignature.reg); a new line is written as `\n`.

Without the setting, one has to configure Outlook:

Tools > Options > Mail Format > Signatures...

* Create a signature that is only used for reply and forward. You have to insert at least `%Q` to get the quoted original mail.
* Assign this signature to every mail account you want to use.

## Keep the Outlook signature

With `USE_QUOTING_TEMPLATE`, the template replaces the signature Outlook puts into the reply.
With `KEEP_SIGNATURE` set to `True`, that signature stays below the template: the template has the greeting and the quote, the signature is maintained in Outlook as before.
Outlook chooses the signature as usual, that is, the one set for replies of the account.

* A colored reply (`USE_COLORIZER`) to an HTML mail keeps the signature as HTML, with its pictures (e.g., a logo).
  If the mail is converted to plain text before it is sent (`COLORIZER_SEND_AS_PLAIN`), the pictures are removed, and the text of the signature remains (bold text as `*text*`).
* Otherwise, the reply is a plain text mail and gets the text of the signature.

See [`configs/exampleTemplateWithSignature.reg`](https://github.com/macros4outlook/quotefixmacro/blob/main/configs/exampleTemplateWithSignature.reg).

## English replies

`QUOTING_TEMPLATE_EN` is the template for replies to mails written in English.
If `USE_QUOTING_TEMPLATE` is `True`, it is used automatically when the original mail is detected as English (the same detection as for the header of the original mail), and always by `FixedReplyAllEnglish()`.
For a mail in any other language, or if the language is not clear, `QUOTING_TEMPLATE` is used.

## Condense Headers

In a reply, the header of each older mail within the quoted text (`From:`, `Sent:`, `To:`, `Subject:`, in any language, with or without the line `-----Original Message-----`; also the one-line attribution of ticket systems, `01.10.2026 16:15 - Firstname Lastname schrieb:`) is condensed to one line, and the text below it gets one quote level more:

```text
Art Ross wrote on 2011-04-07 09:36:
> Hi Adam,
>
> Adam Swift wrote on 2011-04-06 15:12:
>> is it ok?
```

* `CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS` switches this off (`False`).
* `CONDENSED_HEADER_FORMAT` is the format of the condensed line, by default `%SN wrote on %D:`.
  Placeholders: `%SN` sender, `%SE` address of the sender, `%D` date (in `DATE_FORMAT`), `%TO` recipients.
* `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER` (default `True`) also condenses the header of the mail you reply to.
  With `False`, the quote starts without any line about the original mail, so your template has to say it: either a line of your own (e.g., "You wrote on %D:") or `%OH`, which stands for the whole header of the original mail.

When forwarding (`FixedForward`), only headers marked with `-----Original Message-----` and quoted deeper than the text around them are condensed, as in earlier versions.

### Date format

The date format used is [ISO-8601](https://xkcd.com/1179/), which is `YYYY-MM-DD`.
One can change the format in the variable `DEFAULT_DATE_FORMAT`.

## Colored quotes

With `USE_COLORIZER` set to `True`, the reply is written as HTML mail in which each author has their own color.
This helps to see who wrote what while answering.
The author of a quote level is taken from the condensed header above it ("X wrote on ...:"), which is shown as heading in the author's color.
Your own quoted text is dark gray.
If the headers are not condensed, each quote level gets a color instead.
Text marked as `*bold*` or `_underlined_` (e.g., from the bold or underlined text of an HTML mail) is shown bold or underlined; the markers stay, so that the plain text mail has them, too.
To choose per reply, put the macros `FixedReplyColored` and `FixedReplyPlain` (or `FixedReplyAllColored` and `FixedReplyAllPlain`) on the toolbar: they reply with and without colors, whatever `USE_COLORIZER` says; `FixedReply` and `FixedReplyAll` follow the setting.

To answer inline, press <kbd>Enter</kbd> at the end of a quoted line: the new paragraph is black.
If you split a quoted line in its middle, the text you type keeps the color of the quote; <kbd>Ctrl</kbd>+<kbd>Space</kbd> resets it.
`NUM_QUOTE_COLORS` (at most six) is the number of colors in use: blue, green, purple, amber, teal, brown.

Before the mail is sent, it is converted to a plain text mail, so that the recipients get the same mail as without the colors.
This conversion needs the following procedure in 'ThisOutlookSession' (Project1 > Microsoft Outlook Objects in the Visual Basic Editor); it is part of `ThisOutlookSession.doccls`, the other procedures of that file are optional:

```vb
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
   Call QuoteFixMacro.BeforeSend(Item)
End Sub
```

Without it, the mail is sent as HTML mail with the colors.
With `COLORIZER_SEND_AS_PLAIN` set to `False`, every recipient gets the colors on purpose.
`COLORIZER_HTML_RECIPIENTS` lists the recipients who get the colors nevertheless, as addresses or domains separated by `;` (e.g., `@example.org;jennifer.muster@example.com`); a mail is sent as HTML mail if all of its recipients are listed.
Keep in mind that the plain text alternative of such an HTML mail is made by Outlook, which wraps lines longer than 71 characters anew; readers of plain text see broken quotes.
A mail has one format for all of its recipients: if one of them is not listed, everybody gets plain text.
To send a single mail with its colors whatever the list says, put the macro `SendWithColors` on the toolbar of the message window and use it instead of the Send button.
See [`configs/exampleColoredQuotes.reg`](https://github.com/macros4outlook/quotefixmacro/blob/main/configs/exampleColoredQuotes.reg); [`configs/exampleWrapForPlainTextReaders.reg`](https://github.com/macros4outlook/quotefixmacro/blob/main/configs/exampleWrapForPlainTextReaders.reg) sets `LINE_WRAP_AFTER` to 71 for the readers of plain text mentioned above.

## Strip sender's signature

By default, the sender's signature is removed from the reply. If you don't want this, set `STRIP_SIGNATURE` to `False`.

## SoftWrap

When enabled, this feature resizes the window in a way that the text editor wraps the text automatically after N characters.
Outlook wraps text automatically after sending it, but doesn't display the wrap when editing.
Thus, this is useful to double-check that no new line breaks are introduced by Outlook when sending an email.

One can set `USE_SOFTWRAP` to `False` to disable it.

## Random Signature Generation

In case you want to try out the current "random signature generation", import `RandomSignature.bas`.
