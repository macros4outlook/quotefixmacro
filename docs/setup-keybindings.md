---
nav_order: 3
parent: Setup
---
# Setup buttons and key bindings

## Intercept the normal buttons

Add the content of `ThisOutlookSession.cls` to 'ThisOutlookSession' in the Outlook Visual Basic Macro Editor after having installed QuoteFixMacro and restart Outlook.
You can then use the normal Reply/ReplyAll/Forward buttons.
No need to add custom Buttons to the menubar.
Currently a separate button for "ReplyAllEnglish" is required using the previous method as described in `README.md`.

This also works with the standard reply buttons that appear in the reading pane.
It should also work if the reply event is triggered otherwise (e.g. by another macro) but I have not tested this.

## Remap key bindings

Since Outlook does not support custom keybindings, you cannot map the shortcut <kbd>Ctrl</kbd>+<kbd>R</kbd> to the new FixedReply macro.
Nevertheless, the mapping can be done by using [AutoHotkey](https://www.autohotkey.com/).
It has to listen for <kbd>Ctrl</kbd>+<kbd>R</kbd> in Outlook sessions and send <kbd>Alt</kbd>+<kbd>R</kbd> to Outlook instead of <kbd>Ctrl</kbd>+<kbd>R</kbd>.

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

Similar to the above. Outlook 2010, however, does not enable the use of `&` any more. You have to find out the number of the button in the shortcut bar.
Just press <kbd>Alt</kbd> and you'll see the number. Use that in the Autohotkey Macro.

<!-- markdownlint-disable-file MD033 -->
