---
nav_order: 1
parent: Setup
---
# Templates

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

## Examples

### Simple with some QuoteFixMacro advertisement

```text
Hello %FN,

(inline reply powered by QuoteFixMacro - see https://macros4outlook.github.io/quotefixmacro/)

You wrote on %D:

%Q

Cheers,

Oliver
```

### Cursor above the quote

```text
%FN,

%C

%SN wrote on %D:

%Q

Greetings,
Hans
```

### Minimal template

```text
%FN,

%Q

Best,

Amie
```
