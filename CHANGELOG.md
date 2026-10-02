# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).
Since 2026-10-02, versions follow [Calendar Versioning](https://calver.org/) in the form `YYYY-MM-DD`.

## [Unreleased]

### Added

* With `USE_QUOTING_TEMPLATE`, the English template (`QUOTING_TEMPLATE_EN`) is chosen automatically for a mail written in English; the language is detected from the text, as for the header of the original mail. `FixedReplyAllEnglish` still forces it. [#3](https://github.com/macros4outlook/quotefixmacro/issues/3)
* `COLORIZER_HTML_RECIPIENTS`: recipients (addresses or domains) who get a colored reply as HTML mail, while everybody else gets plain text.
* Example `.reg` files for the configuration are in the new folder `configs/`.
* `%MN` and `%MFN` in templates stand for your own name and first name; the default template uses `%MN` instead of `{Name}` and works without any change.
* `SendWithColors`, a macro for the message window: sends the current colored reply as HTML mail whatever the configuration says.

### Removed

* Removed `QuoteFixWithPar.bas` and `TestPar`, the experiments with par (see the decision record on wrapping).

### Fixed

* The cursor is moved to the quote (or to `%C`) through the editor instead of one `SendKeys` call per line, which could switch off NumLock. [#33](https://github.com/macros4outlook/quotefixmacro/issues/33)

## [2026-10-02] - 2026-10-02

### Added

* Replies to mails which are not plain text mails (HTML, Rich Text) are handled: The quoted text is taken from the HTML of the original mail, quotes within it (`<blockquote>`) are kept as quote levels, and the reply is a plain text mail. This requires the new file `QuoteFixHtml.bas`.
* In a reply, the headers of older mails are condensed even if they lack the line `-----Original Message-----` (as in HTML mails), and the text below such a header gets one quote level more. The new placeholder `%TO` of `CONDENSED_HEADER_FORMAT` stands for the recipients.
* The colored mode (`USE_COLORIZER`) works again, without `mapirtf.dll`: the reply is an HTML mail with one color per author (known from the condensed headers; your own text is gray), or per quote level otherwise. `ThisOutlookSession` converts it to plain text before it is sent (`COLORIZER_SEND_AS_PLAIN`, default `True`).

### Changed

* User documentation is now put inside the "docs/" folder.
* `DEFAULT_QUOTING_TEMPLATE` changed to have a salutation at the beginning.
* A reply to a mail which is not a plain text mail is no longer left to Outlook.
* Replies no longer depend on the Outlook setting "Prefix each line of the original message": QuoteFixMacro prefixes and wraps the original text itself. Forwarding is unchanged.
* In a reply, the header of the original mail is written by QuoteFixMacro. Its language (German or English) follows the language of the original mail. If that cannot be detected, the language of Outlook is used.
* In a reply to a plain text mail, a line between two quotes is kept as an answer if it would have fit into the line above. Before, it was always joined with the quote.
* `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER` is `True` by default: the header of the mail you reply to is condensed, too. Set it to `False` if your template contains a line such as "You wrote on %D:".
* `NUM_RTF_COLORS` is now called `NUM_QUOTE_COLORS`. The old name in the registry is still read.

### Removed

* Removed `CONVERT_TO_PLAIN`: The original mail is not converted anymore.
* Removed the colorizer based on `mapirtf.dll` (`ReadRTF`, `WriteRTF`, `DisplayMailItemByID`). [#18](https://github.com/macros4outlook/quotefixmacro/issues/18)

### Fixed

* The last line of the quoted text was dropped if it followed a line with a deeper quote level.
* Name suffixes (e.g., `Jr.`) are stripped: `Firstname Lastname Jr.` and `Lastname Jr., Firstname` now yield the correct first name and last name.
* Condensed headers: A date without a weekday (e.g., `Sent: April 7, 2011 9:52 AM`) is parsed correctly.
* Condensed headers: If the date or the sender name of a header cannot be determined, the one of the previously condensed header is no longer used.

## [1.9] - 2022-03-08

### Added

* Added `ThisOutlookSession.cls` for directly intercepting the standard reply buttons. [#11](https://github.com/macros4outlook/quotefixmacro/pull/11)
* In case the email target ends with a number, the reversal of names also works. I.e., in the case of `Lastname Firstname <firstname.lastname3@example.org>` "Firstname" is used as firstname.
* Include "Dr." in the name if that is present

### Fixed

* Empty line is kept after the header for forwarded mails [#12](https://github.com/macros4outlook/quotefixmacro/pull/12)
* Single-line emails keep the quoting character in the created reply email

### Changed

* Created separate file `QuoteFixNames.bas` (to ease development) [#22](https://github.com/macros4outlook/quotefixmacro/pull/22)

## [1.8] - 2021-02-06

### Added

* In case departments are added at the end of a name, it is removed (e.g., `Firstname Lastname DEP DEP2` becomes `Firstname Lastname`)
* In case the sender format is `Lastname Firstname <firstname.lastname@example.org>`, it is assumed that the typing of the email (firstname before lastname) is correct.

### Changed

* Default pattern for `%D` (date) now includes time in the format `HH:MM`

### Fixed

* Names with dashes are correctly cased (before, they were converted to First-first)

## [1.7] - 2021-01-24

### Added

* Now merges consequitve `> ` lines into a single line
* Support for extraction of sender's last name (stored in `%LN`)
* `%LN` also supports more complex names (e.g., Dr. John Smith III)
* Add support for replying to calender emails

### Fixed

* If sender writes FIRSTNAME LASTNAME, first name is correctly detected

## [1.6] - 2021-01-15

### Changed

* Homepage and code moved from sourceforge to GitHub.
* Linebreaks in `DEFAULT_QUOTING_TEMPLATE` changed from `vbCr` to `"\n"`

### Added

* Now recognizes `Lastname Firstname` as sender name format, too.
* Internationalization: Add `FixedReplyAllEnglish()` with a separate template for replies in English.
* In case a sender name takes something in braces at the end, that text is removed (e.g., "Test Name (42)" is converted to "Test Name")

### Fixed

* If sender name is encloded in quotes, these quotes are stripped
* Applied fix by "helper-01" to enable macro usage at 64bit Outlook
* Always use "Firstname Lastname" as sender name, even if some names are formatted "Lastname, Firstname"

## [1.5] - 2012-01-11

### Added

* support for fixed firstNames for configured email adresses

### Fixed

* When a mail was signed or encrypted with PGP, the reformatting would yield incorrect results
* When a sender's name could not be determined correctly, it would have thrown an error `5`
* Letters of first name are also lower cased
* Only the first word of a potential first name is used as first name

## [1.4] - 2011-07-04

### Added

* Added `CONDENSE_EMBEDDED_QUOTED_OUTLOOK_HEADERS`, which condenses quoted outlook headers.
  The format of the condensed header is configured at `CONDENSED_HEADER_FORMAT`
* Added `CONDENSE_FIRST_EMBEDDED_QUOTED_OUTLOOK_HEADER`
* Added support for custom template configured in the macro (`QUOTING_TEMPLATE`) - this can be used instead of the signature configuration.
* Added `LoadConfiguration()` so you can store personal settings in the registry. These won't get lost when updating the macro.

### Changed

* Merged SoftWrap and QuoteColorizerMacro into `QuoteFixMacro.bas`

### Fixed

* Fixed compile time constants to work with Outlook 2007 and 2010
* Applied patch 3296731 by Matej Mihelic - Replaced hardcoded call to "MAPI"

## [1.3] - 2011-04-22

### Added

* added support to strip quotes of level N and greater
* more support of alternative name formatting
  * added support of reversed name format (`Lastname, Firstname` instead of `Firstname Lastname`)
  * added support of `LASTNAME firstname` format
  * if no firstname is found, then the destination is used
    * `firstname.lastname@domain` is supported
  * firstName always starts with an uppercase letter
  * Added support for `Dr.`
* added `USE_COLORIZER` and `USE_SOFTWRAP` conditional compiling flags.
  They enable QuoteColorizerMacro and SoftWrapMacro.
* added support of removing the sender's signature
* added `CONVERT_TO_PLAIN` flag to enable viewing mails as HTML first.

### Changed

* check for beginning of quote is now language independent
* splitted code for parsing mailtext from `FixMailText()` into smaller functions
* renamed `fromName` to `senderName` to reflect real content of the variable

### Fixed

* included `%C` patch 2778722 by Karsten Heimrich
* included `%SE` patch 2807638 by Peter Lindgren
* `FinishBlock()` would in some cases throw error `5`
* Prevent error 91 when mail is marked as possible phishing mail
* Original mail is marked as read
* fixed cursor position in the case of absence of `%C`, but presence of `%Q`

## [1.2b] - 2007-01-24

### Added

* included on-behalf-of handling written by Per Soderlind

## [1.2a] - 2006-09-26

### Fixed

* quick fix of bug introduced by reformating first-level-quotes (it was reformated too often)

## [1.2] - 2006-09-25

### Added

* QuoteFix now also fixes newly introduced first-level-quotes (`> text`)
* Header matching now matches the English header

## [1.1] - 2006-09-15

### Added

* Macro `%OH` introduced

### Changed

* Outlook header contains `> ` at the end
* If no macros are in the signature, the default behavior of outlook (insert header and quoted text) text is used. (1.0a removed the header)

## [1.0a] - 2006-09-14

### Added

* First public release

[Unreleased]: https://github.com/macros4outlook/quotefixmacro/compare/v2026-10-02...HEAD
[2026-10-02]: https://github.com/macros4outlook/quotefixmacro/compare/v1.9...v2026-10-02
[1.9]: https://github.com/macros4outlook/quotefixmacro/compare/v1.8...v1.9
[1.8]: https://github.com/macros4outlook/quotefixmacro/compare/v1.7...v1.8
[1.7]: https://github.com/macros4outlook/quotefixmacro/compare/v1.6...v1.7
[1.6]: https://github.com/macros4outlook/quotefixmacro/compare/v1.5...v1.6
[1.5]: https://github.com/macros4outlook/quotefixmacro/compare/v1.4...v1.5
[1.4]: https://github.com/macros4outlook/quotefixmacro/compare/v1.3...v1.4
[1.3]: https://github.com/macros4outlook/quotefixmacro/compare/v1.2b...v1.3
[1.2b]: https://github.com/macros4outlook/quotefixmacro/releases/tag/v1.2b
[1.2a]: https://github.com/macros4outlook/quotefixmacro/commits/v1.2b
[1.2]: https://github.com/macros4outlook/quotefixmacro/commits/v1.2b
[1.1]: https://github.com/macros4outlook/quotefixmacro/commits/v1.2b
[1.0a]: https://github.com/macros4outlook/quotefixmacro/commits/v1.2b

<!-- markdownlint-disable-file MD024 -->
