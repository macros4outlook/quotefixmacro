---
parent: Decisions
nav_order: 1
status: accepted
date: 2026-10-01
decision-makers: "@koppor"
---
# Wrap quoted text with built-in code instead of par

## Context and Problem Statement

Since QuoteFixMacro prefixes the original text of a reply itself (`QuoteText`), Outlook no longer wraps the quoted text.
A line which the sender wrapped at 75 characters is 77 characters long after it got the prefix `>` and a space.
Outlook wraps at 76 characters (its default) when the reply is sent and thereby breaks the quote.
Therefore, QuoteFixMacro has to wrap the quoted text before it is sent.

[par](http://www.nicemice.net/par/) is a paragraph reformatter which knows about quote prefixes.
An integration was started in 2011 ([#1](https://github.com/macros4outlook/quotefixmacro/pull/1), `QuoteFixWithPar.bas`) and stopped, because par did not provide the expected output.

Which code should wrap the quoted text?

## Decision Drivers

* A quote has to keep its quote levels.
* Lines which are short on purpose (greeting, closing, list items, signature) should stay as they are.
* Quotes which Outlook broke in older replies of a thread should be repaired.
* QuoteFixMacro should work after importing the `.bas` files, without installing other programs.
* The reply window should open without a noticeable delay.

## Considered Options

* Built-in code, only for paragraphs with a line which does not fit
* par for the complete quoted text
* par, only for paragraphs with a line which does not fit
* Wrap each line on its own

## Decision Outcome

Chosen option: "Built-in code, only for paragraphs with a line which does not fit", because it is the only option which repairs quotes broken by Outlook, it leaves text alone which already fits, and it needs no other program.

The rules for the quoted text of a plain text mail (`ReFormatText` with `SourcePlainText`):

* A paragraph is a sequence of lines of the same quote level without an empty line in between.
* A paragraph is wrapped anew if one of its lines does not fit and is at most 80 characters long.
  Such a line is regarded as wrapped by the sender.
* A longer line is regarded as a paragraph of its own and is wrapped without joining it with other lines.
* A line with a lower quote level between two lines of the same quote level is joined with the line above, if its first word did not fit into that line.
  This is what Outlook produces when it breaks a quoted line.
  Otherwise, the line is an answer between two quotes and is kept.

Text converted from HTML (`SourceHtml`) is not repaired, because its quote levels come from the markup.
Each of its lines is a paragraph and is wrapped on its own.

### Consequences

* Good, because nothing has to be installed and the reply opens without calling another program.
* Good, because a mail whose lines fit is quoted unchanged.
* Good, because the behavior is covered by the tests in `tests/QuoteFixMacroTest.bas`.
* Bad, because a paragraph which is wrapped anew swallows a greeting or a closing which is not separated by an empty line (par does the same).
* Bad, because the repair is a heuristic: an answer between two quotes is still taken for a broken line if the quoted line above it is nearly full.
* Bad, because the lines are filled greedily, whereas par balances the line lengths of a paragraph.
* [#1](https://github.com/macros4outlook/quotefixmacro/pull/1) is closed.
  `QuoteFixWithPar.bas` (reformatting the selected text with par through the clipboard) and `TestPar` in `Tools.bas` are removed, too.

### Confirmation

The tests of the category `reformat` with `SourcePlainText` and `SourceHtml` in `tests/QuoteFixMacroTest.bas` cover the rules above, among them `reformatPlainTextWrapsParagraphOfSenderAnew`, `reformatPlainTextWrapsLongLineOnItsOwn`, `reformatPlainTextKeepsAnswersBetweenQuotes`, and `reformatPlainTextRepairsWrapsOfOlderReplies`.

## Pros and Cons of the Options

### Built-in code, only for paragraphs with a line which does not fit

* Good, because quotes broken by Outlook are repaired.
* Good, because lists, closings and signatures stay as they are, unless they belong to a paragraph which has to be wrapped.
* Good, because no other program is needed.
* Bad, because the repair can take an answer between two quotes for a broken line.
* Bad, because the wrapping is simple (greedy).

### par for the complete quoted text

* Good, because par keeps the quote levels of a well-formed mail apart, also for answers between quotes.
* Good, because par wraps paragraphs well.
* Bad, because par wraps every paragraph, so that lines which are short on purpose are joined: list items, the closing, the signature (see the first example).
  This holds for mails of Thunderbird and of QuoteFixMacro, too: it does not depend on Outlook having broken something.
* Bad, because a quote broken by Outlook stays broken (see the third example).
  QuoteFixMacro cannot know in advance whether a mail contains such a quote, as the broken quote can stem from any older reply of the thread.
* Bad, because par puts an empty line between two quote levels, which makes answers between quotes longer (see the second example).
* Bad, because par has to be installed (Cygwin, WSL, or Docker), and each reply starts a process.

### par, only for paragraphs with a line which does not fit

* Good, because lines which are short on purpose stay as they are, as with the built-in code.
* Good, because par wraps a paragraph better than the built-in code.
* Bad, because the built-in code is still needed to find the paragraphs and to repair the quotes broken by Outlook.
* Bad, because par has to be installed, and each reply starts a process.
* Neutral, because it can be added later as an optional setting without changing the rules above.

### Wrap each line on its own

* Good, because no lines are joined at all.
* Bad, because each full line of a paragraph wrapped by the sender gets a line with one or two words behind it.

## More Information

### How par was checked

par 1.53 (Debian package `par` 1.53.0-2) in a Docker container (`debian:stable-slim`), called with the settings of `QuoteFixWithPar.bas`:

```text
PARINIT='rTbgqR B=.,?_A_a Q=_s>|' par 75q
```

One call through `docker run --rm -i` took about 0.3 seconds.
The sample of `TestPar` (formerly in `Tools.bas`), for which par "combines all the lines together" according to the comment of 2011, is wrapped correctly by this version, both with LF and with CRLF line ends.

In the outputs below, blanks at the end of the lines are removed.

### Example 1: a well-formed mail

Input (the mail of the sender, prefixed for the reply):

```text
> Hi Adam,
>
> this is a paragraph which the sender wrapped at seventy-five characters, as
> QuoteFixMacro does it. With the prefix of the reply, the lines are too long
> by two characters.
>
> - first item of a list
> - second item of a list
>
> Thanks
> Art
>
> --
> Art Ross
> Example Inc.
> Phone 0123 456
```

par:

```text
> Hi Adam,
>
> this is a paragraph which the sender wrapped at seventy-five characters,
> as QuoteFixMacro does it. With the prefix of the reply, the lines are too
> long by two characters.
>
> - first item of a list second item of a list
>
> Thanks Art
>
> -- Art Ross Example Inc.  Phone 0123 456
```

Built-in code:

```text
> Hi Adam,
>
> this is a paragraph which the sender wrapped at seventy-five characters,
> as QuoteFixMacro does it. With the prefix of the reply, the lines are
> too long by two characters.
>
> - first item of a list
> - second item of a list
>
> Thanks
> Art
>
> --
> Art Ross
> Example Inc.
> Phone 0123 456
```

### Example 2: answers between quotes

Input:

```text
> > > Should we meet on Monday?
> > Tuesday fits better.
> > > At ten?
> > Yes.
> Fine with me.
```

par:

```text
> > > Should we meet on Monday?
> >
> > Tuesday fits better.
> >
> > > At ten?
> >
> > Yes.
>
> Fine with me.
```

Built-in code:

```text
>>> Should we meet on Monday?
>> Tuesday fits better.
>>> At ten?
>> Yes.
> Fine with me.
```

### Example 3: a quote broken by Outlook

`OpenLDAP` belongs to the line above it.
Outlook moved it into a line of its own when an older reply was written.

Input:

```text
> > I have a Win 2k3 SBS and I want to replicate the users into my
> OpenLDAP
> > 2.4.11.
>
> This is not possible.
```

par:

```text
> > I have a Win 2k3 SBS and I want to replicate the users into my
>
> OpenLDAP
>
> > 2.4.11.
>
> This is not possible.
```

Built-in code:

```text
>> I have a Win 2k3 SBS and I want to replicate the users into my OpenLDAP
>> 2.4.11.
>
> This is not possible.
```
