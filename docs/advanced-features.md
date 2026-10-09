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

`QUOTING_TEMPLATE_EN` is the template for replies to mails written in English.
If `USE_QUOTING_TEMPLATE` is `True`, it is used automatically when the original mail is detected as English (the same detection as for the header of the original mail), and always by `FixedReplyAllEnglish()`.
For a mail in any other language, or if the language is not clear, `QUOTING_TEMPLATE` is used.

## Replying to HTML mails

QuoteFixMacro also handles mails which are not plain text mails (HTML and Rich Text).
The original mail is left untouched.
The reply is a plain text mail, whose quoted text is taken from the HTML of the original mail:

* Quotes within the original mail (`<blockquote>`, as written by Gmail, Thunderbird, and Apple Mail) are kept as quote levels.
* Each paragraph is wrapped on its own at `LINE_WRAP_AFTER`.
* A link is shown as `text <target>`.
* Pictures and formatting are lost.

Forwarding such a mail is left to Outlook.

The setting `CONVERT_TO_PLAIN` of former versions does not exist anymore.

## Header of the original mail

In a reply, the header of the original mail (`-----Original Message-----`, which `%OH` stands for) is written by QuoteFixMacro.
Its language follows the language of the original mail, which is detected by counting frequent German and English words in the newest part of the mail.
If there is no clear result (e.g., for a very short mail), the header is German if Outlook is German, and English otherwise.
There is no setting for it.

## Wrapping of the quoted text

In a reply, QuoteFixMacro prefixes and wraps the original text itself.
The rules, and why [par](http://www.nicemice.net/par/) is not used for that, are described in the [decision on wrapping](https://macros4outlook.github.io/quotefixmacro/decisions/0001-wrap-quoted-text-with-built-in-code-instead-of-par.html).

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
  Set it to `False` if your template says that already (e.g., "You wrote on %D:").

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

## Use templates from the code

Instead of confuring a template in the signature setting, one can set `DEFAULT_USE_QUOTING_TEMPLATE` to `True`.
Then, QuoteFixMacro reads the signature from `DEFAULT_QUOTING_TEMPLATE_EN` for English emails and from `DEFAULT_QUOTING_TEMPLATE` for all other languages.

## Random Signature Generation

In case you want to try out the current "random signature generation", import `RandomSignature.bas`.

<!-- markdownlint-disable-file MD033 -->
