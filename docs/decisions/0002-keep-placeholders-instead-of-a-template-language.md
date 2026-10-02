---
parent: Decisions
nav_order: 2
status: accepted
date: 2026-10-02
decision-makers: "@koppor"
---
# Keep placeholders instead of a template language

## Context and Problem Statement

A reply is built from a template: the signature configured in Outlook, or `QUOTING_TEMPLATE` in the registry.
The template is plain text with placeholders such as `%FN` (first name of the sender), `%D` (date), `%Q` (the quote) and `%C` (the cursor).

Mail and news readers such as [GoldED+](https://github.com/golded-plus/golded-plus) have a template language of their own: conditional sections (`@reply`, `@forward`, `@moved`), dozens of macros (`@ofrom`, `@odate`, `@oecho`), and language files.
Should QuoteFixMacro adopt such a language, or keep its placeholders?
See [#20](https://github.com/macros4outlook/quotefixmacro/issues/20).

## Decision Drivers

* A new user should get a sensible reply without configuring anything.
* The template is edited in Outlook's signature dialog or in the registry, by users who are not programmers.
* The macro is VBA without a parser library; every construct of a template language is code to write, test and document.
* The things people want to vary are few: salutation, attribution line, closing, where cursor and quote go.

## Considered Options

* Keep the placeholders, add one when a need shows up
* Adopt the GoldED+ template language
* Design a template language of our own with conditions and includes

## Decision Outcome

Chosen option: "Keep the placeholders, add one when a need shows up", because the placeholders cover what users have asked for in twenty years of the macro, they are understood at a glance, and they cost nothing to parse.

The reply mode (reply, reply all, forward) and the language of the original mail are decided in code, not in the template: the English template is chosen for an English mail, the forward keeps Outlook's header.
Information which is the same for every mail of a user, such as the user's own name, becomes a placeholder (`%MN`) rather than a configuration value, so that the default template works unchanged.

### Consequences

* Good, because the default template works after the import of the `.bas` files, without any setting.
* Good, because a template is a few lines of text that anybody can change.
* Bad, because a template cannot express conditions: a different closing for forwards, or a different salutation for a known person, needs a second template or code.
* Neutral, because placeholders are added when needed; each one is one `Replace$` call and one line in the table of `docs/setup-templates.md`.

### Confirmation

`docs/setup-templates.md` lists every placeholder; the default template in `QuoteFixMacro.bas` uses nothing but placeholders.

## Pros and Cons of the Options

### Keep the placeholders, add one when a need shows up

* Good, because nothing has to be learned or parsed.
* Good, because the signature dialog of Outlook is enough to write a template.
* Bad, because every variation beyond substitution needs code.

### Adopt the GoldED+ template language

* Good, because it exists and is documented, with conditional sections and many macros.
* Bad, because most of its macros are about echomail areas, origin lines and message ids, which have no meaning for Outlook mail.
* Bad, because its parser would have to be written anew in VBA.
* Bad, because the users of QuoteFixMacro do not know it.

### Design a template language of our own with conditions and includes

* Good, because it could express everything the macro decides in code today.
* Bad, because it is a parser plus documentation plus tests, for variations nobody has asked for.
* Bad, because templates stop being readable at a glance.

## More Information

The decision was taken in [#20](https://github.com/macros4outlook/quotefixmacro/issues/20) ("the current configuration possibility is IMHO enough").
It can be revisited when a user asks for something the placeholders cannot express; the first candidate is a closing depending on the reply mode.
