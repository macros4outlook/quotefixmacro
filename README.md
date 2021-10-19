# QuoteFixMacro

QuoteFixMacro can modify MS Outlook's message composition windows on-the-fly to allow for **correct quoting** and to change the appearance of your plain-text replies and forwards in general: move your signature, use compressed indentation, customize your quote header, etc.

This style of quoting is described in [Chapter 6. Email Quotes and Inclusion Conventions of The Jargon File](http://www.catb.org/jargon/html/email-style.html).

## Quoting / formatting

If you use Outlook as your email client and make use of plain-text messages, you will have noticed that Outlook doesn't exactly feature the most intelligent quoting algorithm; in fact, it's the silliest one imaginable.

The following will probably look familiar...

```text
See what I mean?

> -----Original Message-----
> From: Dominik Jain [mailto:djain@web.de]
> Sent: Sunday, August 11, 2002 10:15 PM
> To: Dominik Jain
> Subject: RE: test
>
>
> This is a sample text, and as you will see, Microsoft Outlook
> will wrap this
> line. And since I'd already written a fix for Outlook Express, I thought,
> "What the hell... I might as well help you guys, too!"
>
> > -----Original Message-----
> > From: Dominik Jain [mailto:djain@web.de]
> > Sent: Sunday, August 11, 2002 10:15 PM
> > To: Dominik Jain
> > Subject: RE: test
> >
> >
> > This is a sample text, and as you will see, Microsoft Outlook
> > will wrap this
> > line. And since I'd already written a fix for Outlook Express,
> I thought,
> > "What the hell... I might as well help you guys, too!"
> >
> > > -----Original Message-----
> > > From: Dominik Jain [mailto:djain@web.de]
> > > Sent: Sunday, August 11, 2002 10:14 PM
> > > To: djain@web.de
> > > Subject: test
> > >
> > >
> > > This is a sample text, and as you will see, Microsoft Outlook
> > > will wrap this
> > > line. And since I'd already written a fix for Outlook Express,
> > I thought,
> > > "What the hell... I might as well help you guys, too!"
> > >
> >
>
```

Now, what is wrong with this email?

* Horribly broken quotes – line breaks in several wrong places!
* Outlook basically forced me to do a 'top-post', because empty lines were inserted at the top and the cursor was positioned there. And if I had used a signature, it would have been inserted at the top, too.
* Empty lines at the end of the message were quoted.
* The quote header is far too long (5 lines + 2 empty lines) and cannot be modified for a 'personal touch'.

BUT... It doesn't have to be that way. QuoteFixMacro fixes all of the above – and more! The following is an example of how the above message dialog could have looked with QuoteFixMacro:

```text
Dominik Jain wrote:
> Dominik Jain wrote:
>> Dominik Jain wrote:
>>> This is a sample text, and as you will see, Microsoft Outlook will
>>> wrap this line. And since I'd already written a fix for Outlook
>>> Express, I thought, "What the hell... I might as well help you guys,
>>> too!"
>>
>> This is a sample text, and as you will see, Microsoft Outlook will
>> wrap this line. And since I'd already written a fix for Outlook
>> Express, I thought, "What the hell... I might as well help you guys,
>> too!"
>
> This is a sample text, and as you will see, Microsoft Outlook will
> wrap this line. And since I'd already written a fix for Outlook
> Express, I thought, "What the hell... I might as well help you guys,
> too!"

See what I mean?
```

And the best thing about QuoteFixMacro is, **there is absolutely nothing you have to do**. It's all done automatically. You click reply and QuoteFixMacro will immediately reformat the message for proper quoting! No message you send is ever going to look like the one in the example at the top again... And if you get all your friends to use this program, too, such messy quotes will soon become the exception to the rule. And even if you don't, QuoteFixMacro will attempt to fix all their messy quotes when you reply to their messages!

## Setup

### Download

Download the latest version from the [GitHub release page](https://github.com/macros4outlook/quotefixmacro/releases).
You can also download the latest development version using <https://github.com/macros4outlook/quotefixmacro/archive/refs/heads/main.zip>.

### Import macros

1. Extract the downloaded zip-file
2. Open Outlook's VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd> or "Tools > Macro > Visual Basic-Editor")
3. File > Import File ... > Select `QuoteFixMacro.bas` > Open
4. If you don't want to get a security warning when you use the macros, go to "Tools > Macro > Security" and disable the security check. A better solution is to sign the macro. See "Signing a Macro" below.

### Advanced features

In case you want to try out the current "random signature generation", import `RandomSignature.bas`.

The `QuoteFixWithPAR.bas` is not ready for testing yet.

You can easily import all files at once by dragging them from the Explorer into the VBA editor and dropping them onto the project tree.

### Assign macros to buttons

After importing the module, you need to replace the original "Reply" and "ReplyAll" buttons with buttons that trigger the macros defined in the file you just imported. Remember, these buttons are in Outlook's main window, and also in the message window that pops up when you double click on an email.

1. Right-click on the toolbar and select "Customize..."
2. Go to the "Commands" tab and navigate to the "Macro" category
3. Drag the "FixedReply" and "FixedReplyAll" entries and drop it onto the toolbar

You can also change the name and image of the newly created buttons using the customization dialog. If you use "Fixed&Reply" as the name, <kbd>Alt</kbd>+<kbd>R</kbd> is kept as a shortcut for reply. Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro. Nevertheless, the mapping can be done by using AutoHotkey (see below).

### Set up email

1. Tools > Options > Preferences > E-mail Options... > On replies and forwards

   * Change the value "When replying to a message" to "Prefix each line of the original message"
   <!-- markdownlint-disable-next-line MD038 -->
   * Set "Prefix each line with" to "`> `"
   * Change the value "When replying to a message" back to to "Include original message text"

2. Tools > Options > Mail Format

   * Message format: Plain Text
   * Note: this is not necessary, if `CONVERT_TO_PLAIN` is set to `True`.

3. Tools > Options > Mail Format > Internet Format...

   * Automatic wordwrap after: 76 characters (which is the default when you did not touch that setting)

4. Tools > Options > Mail Format > Signatures...

   * Create a signature that is only used for reply and forward. You have to insert at least `%Q` to get the quoted original mail.
   * Assign this signature to every mail account you want to use.
   * Alternatively, you can configure `QUOTING_TEMPLATE` (see below).

5. Display all E-Mail as Text

   * QuoteFixMacro requires plain text to work. One can either read all emails as plain text from the beginning or set `CONVERT_TO_PLAIN` is set to `True`.
     In case all texts should be read as plain text, see Microsoft [KB 831607](https://support.microsoft.com/en-us/office/change-the-message-format-to-html-rich-text-format-or-plain-text-338a389d-11da-47fe-b693-cf41f792fefa?ui=en-us&rs=en-us&ad=us) for an explanation how to turn on this feature. For Outlook 2010 and later (also described at ["Read email messages in plain text"](https://support.microsoft.com/en-us/office/read-email-messages-in-plain-text-16dfe54a-fadc-4261-b2ce-19ad072ed7e3?ui=en-US&rs=en-US&ad=US)): File > Options > Security Center > Options for the Security Center > E-Mail Security > "Read as Plain Text" > `[X]` Read all standard mail in plain text, `[X]` Read all digitally signed mail in plain text".

## Templates

Templates are **the** place to take full advantage of QuoteFixMacro.
The macro replaces certain tokens in the signature. Therefore, the signature can also be used as a template for a message.

Please double check that the template is used as "Forward/Reply" signature under Extra... > Options > E-Mail-Format > Signatures... > E-Mail-Signature

| Pattern | Description |
| -- | -- |
| `%C` | Where to put the cursor. If no `%C` is given, the cursor is put at the first line of the quote |
| `%Q` | Where to put the quote |
| `%OH` | Original Outlook header |
| `%FN` | Sender's first name |
| `%LN` | Sender's last name |
| `%SN` | Sender's name |
| `%SE` | Sender's email address |
| `%D` | Date of the quoted mail in `yyyy-mm-dd HH:MM` |

### Examples

#### Simple with some QuoteFixMacro advertisement

```text
Hello %FN,

(inline reply powered by QuoteFixMacro - see https://macros4outlook.github.io/quotefixmacro/)

You wrote on %D:

%Q

Cheers,

Oliver
```

#### Cursor above the quote

```text
%FN,

%C

%SN wrote on %D:

%Q

Greetings,
Hans
```

#### Minimal template

```text
%FN,

%Q

Best,

Amie
```

### Custom first names

Sometimes, one wants to call someone by a nick name. E.g., one wants to call "Jennifer Muster" just "Jenny". Sometimes, someone does not put his full first name in the email. E.g., "Adelinde Muster" has "<a.muster@example.org>".

QuoteFixMacro supports that replacement. You have to use the registry keys at `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames`.
The key `Count` states how many entries you made.
`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames\1` contains the first entry, `...\2` the second, and so on.
At each entry, there are two keys: `email` stating the email to match and `firstName` the first name to use.

#### Step-by-step instruction

1. Open regedit
1. Navigate to `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro`
1. Create key `firstnames`
1. Create string (!) "Count" with value `X`, where `X` is the number of replacements you want to configure
1. Create key `firstnames.1`
1. Create string value `email` with the email you want to specify a first name for
1. Create string value `firstName` with the first name to be used
1. Repeat steps 5 to 7 until `X` is reached. Replace `1` at `firstnames.1` by the appropriate number

#### Direct import using `.reg` files

Alternatively, create an `example.reg` file with following content and adapt it to your needs. Then double click on "example.reg" and import it into your registry.
The distribution of QuoteFixMacro already contains an [`exampleFirstNameConfiguration.reg`](exampleFirstNameConfiguration.reg) with the content below.

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames]
"Count"="2"

[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames\1]
"email"="jennifer.muster@example.org"
"firstName"="Jenny"

[HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames\2]
"email"="A.Muster@example.org"
"firstName"="Adelinde"
```

## AutoHotkey

Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro. Nevertheless, the mapping can be done by using [AutoHotkey](https://www.autohotkey.com/). It has to listen for <kbd>Ctrl</kbd>+<kbd>R</kbd> in Outlook sessions and send <kbd>Alt</kbd>+<kbd>R</kbd> to Outlook instead of <kbd>Ctrl</kbd>+<kbd>R</kbd>.

### AutoHotkey macro for a German Outlook 2007

```autohotkey
;A class matching is not possible in the outlook 2007 beta 2,
;therefore title matching is used
SetTitleMatchMode, 2

;For the message window
#IfWinActive Nachricht
^r::
Send !6

;For the outlook window
#IfWinActive Outlook
^r::
Send !r
```

Remark: In the message window, the reply button cannot be inserted as in Outlook 2003. The shortcut bar has to be used instead. The button was assigned <kbd>Alt</kbd>+<kbd>6</kbd> after insertion.

### AutoHotkey macro for Outlook 2010 and later

Similar to the above. Outlook 2010, however, does not enable the use of `&` any more. You have to find out the number of the button in the shortcut bar. Just press <kbd>Alt</kbd> and you'll see the number. Use that in the Autohotkey Macro.

## Signing the macro

With Outlook 2016, this doesn't seem to be necessary any more since self-written macros seem to be usable without error.
Otherwise, following article should explain everything: <https://www.groovypost.com/howto/create-self-signed-digital-certificate-microsoft-office-2016/>.
There is also an [MSN Article](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa163622(v=office.10)?redirectedfrom=MSDN) on macro code signing for Office XP.

## Advanced usage

Configuration is done via constants in the QuoteFix code (see below for a storage in the registry)

1. Start the VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd>)
2. Open the module "QuoteFixMacro"
3. Scroll down to the block "Configuration constants"

### Configure the template inside the code

The variable `QUOTING_TEMPLATE` can be used to store the quoting template.
Thus, the Outlook configuration can be left untouched.

### English replies

The variable `QUOTING_TEMPLATE_EN` can be used to store en English quoting template.
In case `USE_QUOTING_TEMPLATE` is `True` and `FixedReplyAllEnglish()` is called, that template is used.

### Auto conversion to plain format

By setting `CONVERT_TO_PLAIN` to `True`, HTML mails are automatically converted to text mails.

### Condense Headers

With `CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS`, one condenses reply/forwarding headers added by outlook so that the email gets even shorter
The format of the condensed header is configured at `CONDENSED_HEADER_FORMAT`

One can also condense the first header only `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER`.

### QuoteColorizer

QuoteColorizer colorizes the indented parts in different colors to make it easier to distinguish who wrote which text. Set `USE_COLORIZER` to `True` to use this. The mail format is automatically set to Rich-Text.

The recipient will receive the colors, too.
In case, you don't want this, enable convert RTF-to-Text at sending.

### Strip sender's signature

By default, the sender's signature is removed from the reply. If you don't want this, set `STRIP_SIGNATURE` to `False`.

### SoftWrap

When enabled, this feature resizes the window in a way that the text editor wraps the text automatically after N characters.
Outlook wraps text automatically after sending it, but doesn't display the wrap when editing.
Thus, this is useful to double-check that no new line breaks are introduced by Outlook when sending an email.

One can set `USE_SOFTWRAP` to `True` to enable it.

### Date format

The date format used is [ISO-8601](https://xkcd.com/1179/), which is `YYYY-MM-DD`.
One can change the format in the variable `DEFAULT_DATE_FORMAT`.

### Persist settings across updates

An update of QuoteFixMacro happens by replacing the content of the `.bas` file.
Thus, any settings are overwritten during an update.
QuoteFixMacro can read settings from the registry.
The macro **NEVER** stores entries in the registry by itself.

You can store the default configuration in the registry:

1. by executing `StoreDefaultConfiguration()`
2. by writing a routing executing command similar to the following: `Call SaveSetting(APPNAME, REG_GROUP_CONFIG, "CONVERT_TO_PLAIN", "true")`
3. by manually creating entries in this registry hive: `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro`

### Intercept the normal buttons

Add the content of `ThisOutlookSession.cls` to 'ThisOutlookSession' in the Outlook Visual Basic Macro Editor after having installed QuoteFixMacro and restart Outlook.
You can then use the normal Reply/ReplyAll/Forward buttons.
No need to add custom Buttons to the menubar.
Currently a separate button for "ReplyAllEnglish" is required using the previous method as described in `README.md`.

This also works with the standard reply buttons that appear in the reading pane.
It should also work if the reply event is triggered otherwise (e.g. by another macro) but I have not tested this.

## FAQ

Q: What if the whole mail text disappears?
<!-- markdownlint-disable-next-line MD038 -->
A: The reply setting in Outlook is not configured as required. Double check that the original text should be prefixed with `> `.

## Developing

We recommend using [Rubberduck](https://rubberduckvba.com/) as plugin to the Visual Basic Editor.

## Acknowledgements

QuoteFix Macro is a VB-Macro for Outlook created by Oliver Kopp, Lars Monsees, and Daniel Martin.
QuoteFix Macro is inspired by [Outlook-QuoteFix](http://web.archive.org/web/20120316151928/http://home.in.tum.de/%7Ejain/software/outlook-quotefix/) written by [Dominik Jain](https://github.com/opcode81/) and reimplemented as a Visual Basic macro.
The ideas for integrating it in Outlook came from [Daniele Bochicchio](https://github.com/dbochicchio), especially from his [quoting macro](http://lab.aspitalia.com/35/Outlook-2007-2003-Reply-With-Quoting-Macro.aspx).

<!-- markdownlint-disable-file MD033 MD038 -->
