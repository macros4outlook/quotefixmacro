Attribute VB_Name = "QuoteFixHtmlTest"

Option Explicit
Option Private Module

'@TestModule
'@Folder("QuoteFixMacro.Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private html As String
Private expectedResult As String

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'The HTML snippets follow the structure of the mails written by the named programs.
'They are written by hand, not taken from real mails.

'@TestMethod("html")
Private Sub htmlParagraphsAreSeparatedByEmptyLine()
    On Error GoTo TestFail

    html = "<html><body><p>Hello Adam,</p>" & vbNewLine & "<p>this is a test.</p></body></html>"
    expectedResult = vbNullString & _
        "Hello Adam," & vbNewLine & _
        vbNewLine & _
        "this is a test."

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlOutlookParagraphsAreLines()
    On Error GoTo TestFail

    html = "<html><body lang=DE><div class=WordSection1>" & _
        "<p class=MsoNormal>Hello Adam,<o:p></o:p></p>" & _
        "<p class=MsoNormal>this is a test.<o:p></o:p></p>" & _
        "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" & _
        "<p class=MsoNormal><span style='font-size:10.0pt'>Art<o:p></o:p></span></p>" & _
        "</div></body></html>"
    expectedResult = vbNullString & _
        "Hello Adam," & vbNewLine & _
        "this is a test." & vbNewLine & _
        vbNewLine & _
        "Art"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlGmailLinesAndNestedQuotes()
    On Error GoTo TestFail

    html = "<div dir=""ltr""><div>Yes.</div><div><br></div><div>Adam</div></div><br>" & _
        "<div class=""gmail_quote""><div dir=""ltr"" class=""gmail_attr"">On Thu, Apr 7, 2011 at 9:36 AM Art Ross &lt;<a href=""mailto:art@example.com"">art@example.com</a>&gt; wrote:<br></div>" & _
        "<blockquote class=""gmail_quote"" style=""margin:0px 0px 0px 0.8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex"">" & _
        "<div>Is it ok?</div>" & _
        "<blockquote class=""gmail_quote""><div>First mail</div></blockquote>" & _
        "</blockquote></div>"
    expectedResult = vbNullString & _
        "Yes." & vbNewLine & _
        vbNewLine & _
        "Adam" & vbNewLine & _
        vbNewLine & _
        "On Thu, Apr 7, 2011 at 9:36 AM Art Ross <art@example.com> wrote:" & vbNewLine & _
        "> Is it ok?" & vbNewLine & _
        ">> First mail"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlThunderbirdParagraphsInQuote()
    On Error GoTo TestFail

    html = "<html><head><meta http-equiv=""content-type"" content=""text/html; charset=UTF-8""></head><body>" & _
        "<p>Reply</p>" & vbNewLine & _
        "<div class=""moz-cite-prefix"">On 07.04.2011 09:36, Art Ross wrote:<br></div>" & vbNewLine & _
        "<blockquote type=""cite"" cite=""mid:1234@example.com"">" & vbNewLine & _
        "<p>First paragraph</p>" & vbNewLine & _
        "<p>Second paragraph</p>" & vbNewLine & _
        "</blockquote></body></html>"
    expectedResult = vbNullString & _
        "Reply" & vbNewLine & _
        vbNewLine & _
        "On 07.04.2011 09:36, Art Ross wrote:" & vbNewLine & _
        "> First paragraph" & vbNewLine & _
        ">" & vbNewLine & _
        "> Second paragraph"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlOutlookHeaderOfQuotedMail()
    On Error GoTo TestFail

    html = "<p class=MsoNormal>Answer<o:p></o:p></p>" & _
        "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" & _
        "<div><div style='border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0cm 0cm 0cm'>" & _
        "<p class=MsoNormal><b>From:</b> Art Ross &lt;art@example.com&gt; <br>" & _
        "<b>Sent:</b> Thursday, April 7, 2011 9:36 AM<br>" & _
        "<b>To:</b> Adam Swift &lt;adam@example.com&gt;<br>" & _
        "<b>Subject:</b> Testing<o:p></o:p></p></div></div>" & _
        "<p class=MsoNormal><o:p>&nbsp;</o:p></p>" & _
        "<p class=MsoNormal>Question<o:p></o:p></p>"
    expectedResult = vbNullString & _
        "Answer" & vbNewLine & _
        vbNewLine & _
        "From: Art Ross <art@example.com>" & vbNewLine & _
        "Sent: Thursday, April 7, 2011 9:36 AM" & vbNewLine & _
        "To: Adam Swift <adam@example.com>" & vbNewLine & _
        "Subject: Testing" & vbNewLine & _
        vbNewLine & _
        "Question"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlHorizontalRule()
    On Error GoTo TestFail

    html = "<div>Answer</div><hr style=""display:inline-block;width:98%"" tabindex=""-1""><div id=""divRplyFwdMsg""><b>From:</b> Art Ross</div>"
    expectedResult = vbNullString & _
        "Answer" & vbNewLine & _
        "________________________________" & vbNewLine & _
        "From: Art Ross"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlWhiteSpaceIsCollapsed()
    On Error GoTo TestFail

    html = "<div>  Hello" & vbNewLine & "   <b>Adam</b> ,  how" & vbTab & "are <i> you</i>?  </div>"

    Assert.AreEqual "Hello Adam , how are you?", HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlCharacterReferences()
    On Error GoTo TestFail

    html = "<p>Gr&uuml;&szlig;e &amp; more &lt;tag&gt; &#8364; &#x20AC; &unknown; A&nbsp;B &#128512;</p>"
    expectedResult = "Gr" & ChrW$(252) & ChrW$(223) & "e & more <tag> " & ChrW$(8364) & " " & ChrW$(8364) & " &unknown; A B " & ChrW$(55357) & ChrW$(56832)

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

    'characters without width are dropped
    Assert.AreEqual "ABC", HtmlToPlainText(ChrW$(65279) & "<p>A&#8203;B" & ChrW$(8205) & "C</p>")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlScriptStyleAndCommentsAreDropped()
    On Error GoTo TestFail

    html = "<html><head><title>Title</title><style><!-- p {color:red} --></style></head><body>" & _
        "<!-- a comment with <p>a tag</p> -->" & _
        "<!--[if gte mso 9]><xml><o:shapedefaults v:ext=""edit""/></xml><![endif]-->" & _
        "<p>Text</p>" & _
        "<script type=""text/javascript"">if (1 < 2) { document.write(""<p>written</p>""); }</script>" & _
        "<STYLE>p {margin:0}</STYLE>" & _
        "</body></html>"

    Assert.AreEqual "Text", HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlLinks()
    On Error GoTo TestFail

    html = "<p>See <a href=""https://example.com/page"">our page</a> or <a href=""https://example.com/"">https://example.com</a>, " & _
        "mail <a href=""mailto:art@example.com?subject=Hi"">Art</a> or jump <a href=""#top"">up</a>.</p>"

    Assert.AreEqual "See our page <https://example.com/page> or https://example.com, mail Art <art@example.com> or jump up.", HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlMentionIsKeptAsText()
    On Error GoTo TestFail

    html = "<p class=MsoNormal><a id=""OWAAM123"" href=""mailto:art@example.com""><span style='font-family:""Calibri"",sans-serif;text-decoration:none'>@Ross, Art</span></a>, could you have a look?</p>"

    Assert.AreEqual "@Ross, Art, could you have a look?", HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlLists()
    On Error GoTo TestFail

    html = "<div>Todo:</div><ul><li>one</li><li>two<ul><li>two a</li></ul></li></ul><ol><li><p>first</p></li><li>second</li></ol>"
    expectedResult = vbNullString & _
        "Todo:" & vbNewLine & _
        "* one" & vbNewLine & _
        "* two" & vbNewLine & _
        "  * two a" & vbNewLine & _
        "1. first" & vbNewLine & _
        "2. second"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlTableRowsAreLines()
    On Error GoTo TestFail

    html = "<table><tbody><tr><th>Name</th><th>Value</th></tr><tr><td>a</td><td>1</td></tr></tbody></table>"
    expectedResult = vbNullString & _
        "Name Value" & vbNewLine & _
        "a 1"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlPreformattedText()
    On Error GoTo TestFail

    html = "<div>Code:</div><pre>" & vbNewLine & "line 1" & vbNewLine & "  indented   line" & vbNewLine & "</pre><div>Done</div>"
    expectedResult = vbNullString & _
        "Code:" & vbNewLine & _
        "line 1" & vbNewLine & _
        "  indented   line" & vbNewLine & _
        "Done"

    Assert.AreEqual expectedResult, HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlTagEndWithinAttributeValue()
    On Error GoTo TestFail

    html = "<p title=""a > b"" class='x'>Text <img src=""cid:image001.png"" alt=""a picture""> 1 < 2</p>"

    Assert.AreEqual "Text [a picture] 1 < 2", HtmlToPlainText(html)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("html")
Private Sub htmlPlainText()
    On Error GoTo TestFail

    Assert.AreEqual "no tags at all", HtmlToPlainText("no tags at all")
    Assert.AreEqual vbNullString, HtmlToPlainText(vbNullString)
    Assert.AreEqual vbNullString, HtmlToPlainText("<html><body><p>&nbsp;</p></body></html>")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
