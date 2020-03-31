# ClipboardRtfToHtml

This project was my sandbox for playing with solutions to paste rich text from Visual
Studio into OneNote, preserving colorized syntax highlighting.

OneNote doesn't allow pasting RTF directly; it never seemed to implement that feature
and instead just pastes the raw text. So this solution uses the clipboard as a working
buffer to convert RTF to HTML.

The conversion is done using the .NET WPF RichTextBox component control. It can take
RTF and convert its content to XAML. We can then parse the XAML and convert it HTML.

Then the easiest way to get the HTML into OneNote, without hacking away at the page
HTML directly, is to paste it from the clipboard. This is done simply by invoking
the Ctrl-V key sequence.

