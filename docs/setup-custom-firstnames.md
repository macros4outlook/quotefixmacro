---
nav_order: 2
parent: Setup
---
# Custom first names

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
