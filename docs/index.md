---
nav_order: 1
title: "Introduction"
---
# QuoteFixMacro

QuoteFixMacro can modify MS Outlook's message composition windows on-the-fly to allow for **correct quoting** and to change the appearance of your plain-text replies and forwards in general: move your signature, use compressed indentation, customize your quote header, etc.

This style of quoting is described in [Chapter 6. Email Quotes and Inclusion Conventions of The Jargon File](http://www.catb.org/jargon/html/email-style.html).

## Quoting / formatting

If you use Outlook as your email client and make use of plain-text messages, you will have noticed that Outlook doesn't exactly feature the most intelligent quoting algorithm; in fact, it's the silliest one imaginable.

The following will probably look familiar...

```text
See what I mean?

> -----Original Message-----
> From: Dominik Jain [mailto:djain@web.de]
> Sent: Sunday, August 11, 2002 10:15 PM
> To: Dominik Jain
> Subject: RE: test
>
>
> This is a sample text, and as you will see, Microsoft Outlook
> will wrap this
> line. And since I'd already written a fix for Outlook Express, I thought,
> "What the hell... I might as well help you guys, too!"
>
> > -----Original Message-----
> > From: Dominik Jain [mailto:djain@web.de]
> > Sent: Sunday, August 11, 2002 10:15 PM
> > To: Dominik Jain
> > Subject: RE: test
> >
> >
> > This is a sample text, and as you will see, Microsoft Outlook
> > will wrap this
> > line. And since I'd already written a fix for Outlook Express,
> I thought,
> > "What the hell... I might as well help you guys, too!"
> >
> > > -----Original Message-----
> > > From: Dominik Jain [mailto:djain@web.de]
> > > Sent: Sunday, August 11, 2002 10:14 PM
> > > To: djain@web.de
> > > Subject: test
> > >
> > >
> > > This is a sample text, and as you will see, Microsoft Outlook
> > > will wrap this
> > > line. And since I'd already written a fix for Outlook Express,
> > I thought,
> > > "What the hell... I might as well help you guys, too!"
> > >
> >
>
```

Now, what is wrong with this email?

* Horribly broken quotes – line breaks in several wrong places!
* Outlook basically forced me to do a 'top-post', because empty lines were inserted at the top and the cursor was positioned there. And if I had used a signature, it would have been inserted at the top, too.
* Empty lines at the end of the message were quoted.
* The quote header is far too long (5 lines + 2 empty lines) and cannot be modified for a 'personal touch'.

BUT... It doesn't have to be that way. QuoteFixMacro fixes all of the above – and more! The following is an example of how the above message dialog could have looked with QuoteFixMacro:

```text
Dominik Jain wrote:
> Dominik Jain wrote:
>> Dominik Jain wrote:
>>> This is a sample text, and as you will see, Microsoft Outlook will
>>> wrap this line. And since I'd already written a fix for Outlook
>>> Express, I thought, "What the hell... I might as well help you guys,
>>> too!"
>>
>> This is a sample text, and as you will see, Microsoft Outlook will
>> wrap this line. And since I'd already written a fix for Outlook
>> Express, I thought, "What the hell... I might as well help you guys,
>> too!"
>
> This is a sample text, and as you will see, Microsoft Outlook will
> wrap this line. And since I'd already written a fix for Outlook
> Express, I thought, "What the hell... I might as well help you guys,
> too!"

See what I mean?
```

And the best thing about QuoteFixMacro is, **there is absolutely nothing you have to do**. It's all done automatically. You click reply and QuoteFixMacro will immediately reformat the message for proper quoting! No message you send is ever going to look like the one in the example at the top again... And if you get all your friends to use this program, too, such messy quotes will soon become the exception to the rule. And even if you don't, QuoteFixMacro will attempt to fix all their messy quotes when you reply to their messages!

## Acknowledgements

QuoteFix Macro is a VB-Macro for Outlook created by Oliver Kopp, Lars Monsees, and Daniel Martin.
QuoteFix Macro is inspired by [Outlook-QuoteFix](http://web.archive.org/web/20120316151928/http://home.in.tum.de/%7Ejain/software/outlook-quotefix/) written by [Dominik Jain](https://github.com/opcode81/) and reimplemented as a Visual Basic macro.
The ideas for integrating it in Outlook came from [Daniele Bochicchio](https://github.com/dbochicchio), especially from his [quoting macro](http://lab.aspitalia.com/35/Outlook-2007-2003-Reply-With-Quoting-Macro.aspx).

<!-- markdownlint-disable-file MD033 MD038 -->
