---
nav_order: 4
---
# Experimental Features

## QuoteColorizer

QuoteColorizer colorizes the indented parts in different colors to make it easier to distinguish who wrote which text. Set `USE_COLORIZER` to `True` to use this. The mail format is automatically set to Rich-Text.

The recipient will receive the colors, too.
In case, you don't want this, enable convert RTF-to-Text at sending.

This feature is currently not working.
Discussions is made available at [#18](https://github.com/macros4outlook/quotefixmacro/issues/18).

## Advanced rewrapping

[par](http://www.nicemice.net/par/) is a tool implementing rewrapping of quotes.
It could be used to replace the internal VBA codde to rewrap messages.
However, the current implementation at `QuoteFixWithPar.bas` is not ready for testing yet.
See <https://github.com/macros4outlook/quotefixmacro/pull/1> for more implementation details.
