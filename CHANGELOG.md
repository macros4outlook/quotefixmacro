# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [Unreleased]

### Changed

* Homepage and code moved from sourceforge to GitHub.
* Linebreaks in `DEFAULT_QUOTING_TEMPLATE` changed from `vbCr` to `"\n"`

### Added

* Now recognizes `LastnameFirstname` as sender name format, too.
* Internationalization: Add `FixedReplyAllEnglish()` with a separate template for replies in English.

### Fixed

* If sender name is encloded in quotes, these quotes are stripped
* Applied fix by "helper-01" to enable macro usage at 64bit Outlook

## Version [1.5] - 2012-01-11

### Added

* support for fixed firstNames for configured email adresses

### Fixed

* When a mail was signed or encrypted with PGP, the reformatting would yield incorrect results
* When a sender's name could not be determined correctly, it would have thrown an error `5`
* Letters of first name are also lower cased
* Only the first word of a potential first name is used as first name

## Version [1.4] - 2011-07-04

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

## Version [1.3] - 2011-04-22

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

## Version 1.2b - 2007-01-24

### Added

* included on-behalf-of handling written by Per Soderlind

## Version 1.2a - 2006-09-26

### Fixed

* quick fix of bug introduced by reformating first-level-quotes (it was reformated too often)

## Version 1.2 - 2006-09-25

### Added

* QuoteFix now also fixes newly introduced first-level-quotes (`> text`)
* Header matching now matches the English header

## Version 1.1 - 2006-09-15

### Added

* Macro `%OH` introduced

### Changed

* Outlook header contains `> ` at the end
* If no macros are in the signature, the default behavior of outlook (insert header and quoted text) text is used. (1.0a removed the header)

## Version 1.0a - 2006-09-14

* first public release

[Unreleased]: https://github.com/macros4outlook/quotefixmacro/compare/v1.5...HEAD
[1.5]: https://github.com/macros4outlook/quotefixmacro/compare/v1.4...v1.5
[1.4]: https://github.com/macros4outlook/quotefixmacro/compare/v1.3...v1.4
[1.3]: https://github.com/macros4outlook/quotefixmacro/compare/v1.2b...v1.3
