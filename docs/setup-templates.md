---
nav_order: 1
parent: Setup
---
# Templates

Templates are **the** place to take full advantage of QuoteFixMacro.
The macro replaces certain tokens in the signature. Therefore, the signature can also be used as a template for a message.
The default templates in `QuoteFixMacro.bas` work without any change and show both styles: `DEFAULT_QUOTING_TEMPLATE` (German, for all mails not detected as English) closes with "Mit freundlichen Grüßen" and your full name (`%MN`), `DEFAULT_QUOTING_TEMPLATE_EN` with "Cheers," and your first name (`%MFN`).

Please double check that the template is used as "Forward/Reply" signature under Extra... > Options > E-Mail-Format > Signatures... > E-Mail-Signature

| Pattern | Description                                                                                    |
| ------- | ---------------------------------------------------------------------------------------------- |
| `%C`    | Where to put the cursor. If no `%C` is given, the cursor is put at the first line of the quote |
| `%Q`    | Where to put the quote                                                                         |
| `%OH`   | Original Outlook header                                                                        |
| `%FN`   | Sender's first name                                                                            |
| `%LN`   | Sender's last name                                                                             |
| `%SN`   | Sender's name                                                                                  |
| `%SE`   | Sender's email address                                                                         |
| `%D`    | Date of the quoted mail in `yyyy-mm-dd HH:MM`                                                  |
| `%MN`   | Your own name, as Outlook knows it (`Firstname Lastname`)                                      |
| `%MFN`  | Your own first name                                                                            |

By default, the quote starts with the line "Sender wrote on date:", so a template does not need to say that.
If you prefer to write that line in the template, see `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER` at [Condense Headers](https://macros4outlook.github.io/quotefixmacro/advanced-features.html#condense-headers).

## Examples

### Simple with some QuoteFixMacro advertisement

```text
Hello %FN,

(inline reply powered by QuoteFixMacro - see https://macros4outlook.github.io/quotefixmacro/)

%Q

Cheers,

%MFN
```

### Cursor above the quote

```text
%FN,

%C

%Q

Greetings,
%MFN
```

### Minimal template

```text
%FN,

%Q

Best,

%MFN
```
