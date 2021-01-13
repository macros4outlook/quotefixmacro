# QuoteFixMacro

QuoteFix Macro is a VB-Macro for Outlook written by Oliver Kopp, Lars Monsees, and Daniel Martin. It works in Outlook 2003, 2007, and 2010. QuoteFix Macro is inspired by [Outlook-QuoteFix](http://web.archive.org/web/20120316151928/http://home.in.tum.de/%7Ejain/software/outlook-quotefix/) written by Dominik Jain and implemented as a Visual Basic macro. The ideas for integrating it in Outlook came from [Daniele Bochicchio](https://github.com/dbochicchio), especially from his [quoting macro](http://lab.aspitalia.com/35/Outlook-2007-2003-Reply-With-Quoting-Macro.aspx).
Contents

## Setup

### Download

Download the latest version from the GitHub release page. The Basic Edition solely contains the QuoteFix Macro. The SoftWrap Edition includes QuoteFix Macro with SoftWrap Macro. SoftWrap Macro is useful if you don't use Outlook maximized.

### Import Macros

1. Extract the downloaded zip-file
2. In Outlook's VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd> or "Tools > Macro > Visual Basic-Editor"), import the downloaded file by right-clicking on "Modules" and selecting "Import...". You can easily import all files at once by dragging them from the Explorer into the VBA editor and dropping them onto the project tree.
3. If you don't want to get a security warning when you use the macros, go to "Tools > Macro > Security" and disable the security check. A better solution is to sign the macro. See "Signing a Macro" below.

### Assign macros to buttons

After that, you need to replace the original "Reply" and "ReplyAll" buttons with buttons that trigger the macros defined in the file you just imported. Remember, these buttons are in Outlook´s main window, and also in the message window that pops up when you double click on an email.

1. Right-click on the toolbar and select "Customize..."
2. Go to the "Commands" tab and navigate to the "Macro" category
3. Drag the "FixedReply" and "FixedReplyAll" entries and drop it onto the toolbar

You can also change the name and image of the newly created buttons using the customization dialog. If you use "Fixed&Reply" as the name, <kbd>Alt</kbd>+<kbd>R</kbd> is kept as a shortcut for reply. Since Outlook does not support custom keybindings, you cannot map the shortcut CTRL+r to the new FixedReply macro. Neverthelesss, the mapping can be done by using Autohotkey (see below).

### Set up eMail

1. Tools > Options > Preferences > E-mail Options... > On replies and forwards

   * When replying to a message: "Prefix each line of the original message"
   * When forwarding a message: "Include original message text" or "Prefix each line of the original message"
   * Prefix each line with: "`> `"

2. Tools > Options > Mail Format

   * Message format: Plain Text

3. Tools > Options > Mail Format > Internet Format...

   * Automatic wordwrap after: 76 characters

4. Tools > Options > Mail Format > Signatures...

   * Create a signature that is only used for reply and forward. You have to insert at least %Q to get the quoted original mail.
   * Assign this signature to every mail account you want to use.

5. Display all E-Mail as Text

   * Otherwise, QuoteFix does not work. -- See Microsoft KB 831607 for an explanation how to turn this feature on.
   * For Outlook 2010: File / Options / Security Center / Options for the Security Center / E-Mail Security / "Read as Plain Text" / `[X]` Read all standard mail in plain text, `[X]` Read all digitally signed mail in plain text"

## Templates

Templates are **the** place to take full advantage of QuoteFixMacro.
The macro replaces certain tokens in the signature. Therefore, the signature can also be used as a template for a message.

Please double check that the template is used as "Forward/Reply" signature under Extra.../Options/E-Mail-Format/Signatures.../E-Mail-Signature

| Pattern | Description |
| -- | -- |
| `%C` | Where to put the cursor. If no `%C` is given, the cursor is put at the first line of the quote |
| `%D` | Date of the quoted mail in `yyyy-mm-dd` |
| `%FN` | Sender's first name |
| `%OH` | Original Outlook header |
| `%SN` | Sender´s name |
| `%Q` | Where to put the quote |

This allows following templates for a reply:

```text
%FN,

%C

%SN wrote on %D:

%Q

Greetings,
Hans
```

or

```text
%FN,

%Q

Best,

Amie
```

## Configuration

Configuration is done via constants in the QuoteFix code:

1. Start the VBA editor (<kbd>Alt</kbd>+<kbd>F11</kbd>)
2. Open the module "QuoteFixMacro"
3. Scroll down to the block "Configuration constants"

### Strip sender's signature

By default, the sender's signature is removed from the reply. If you don´t want this, set `STRIP_SIGNATURE` to `false`.

### QuoteColorizer

QuoteColorizer colorizes the indented parts in different colors to make it easier to distinguish who wrote which text. Set `USE_COLORIZER` to `true` to use this. The mail format is automatically set to Rich-Text.

### Custom Firstnames

Sometimes, one wants to call someone by a nick name. E.g., one wants to call "Jennifer Muster" just "Jenny". Sometimes, someone does not put his full first name in the Email. E.g., "Adelinde Muster" has "<a.muster@example.org>".

QuoteFixMacro supports that replacement. You have to use the registry keys at `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames`.
The key `Count` states how many entries you made.
`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro\firstnames\1` contains the first entry, `...\2` the second, and so on.
At each entry, there are two keys: email stating the email to match and firstName the first name to use.

#### Step-by-step instruction

1. Open regedit
1. Navigate to `HKEY_CURRENT_USER\Software\VB and VBA Program Settings\QuoteFixMacro`
1. Create key `firstnames`
1. Create string (!) "Count" with value `X`, where `X` is the number of replacements you want to configure
1. Create key `firstnames.1`
1. Create string value `email` with the email you want to specify a firstName for
1. Create string value `firstName` with the firstname to be used
1. Repeat steps 5 to 7 until `X` is reached. Replace `1` at `firstnames.1` by the appropriate number

#### Direct Import Using .reg Files

Alternatively, create a `example.reg` file with following content and adapt it to your needs. Then double click on "example.reg" and import it into your registry. The distribution of QuoteFixMacro already contains an "exampleFirstNameConfiguration.reg" with the content below.

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

Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro. Neverthelesss, the mapping can be done by using [AutoHotkey](https://www.autohotkey.com/). It has to listen for <kbd>Ctrl</kbd>+<kbd>R</kbd> in Outlook sessions and send <kbd>Alt</kbd>+<kbd>R</kbd> to Outlook instead of <kbd>Ctrl</kbd>+<kbd>R</kbd>.

### AutoHotkey Macro for a German Outlook 2007

```autohotkey
;A class matching is not possible in the outlook 2007 beta 2,
;thefore title matching is used
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

### AutoHotkey Macro for a Outlook 2010

Similar to the above. Outlook 2010, however, does not enable the use of `&` any more. You have to find out the number of the button in the shortcut bar. Just press <kbd>Alt</kbd> and you'll see the number. Use that in the Autohotkey Macro.

## Signing the Macro

That article should explain everything: <https://www.groovypost.com/howto/create-self-signed-digital-certificate-microsoft-office-2016/>.

There is a [MSN Article](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa163622(v=office.10)?redirectedfrom=MSDN) on Macro code signing for Office XP.