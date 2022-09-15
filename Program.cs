//************************************************************************************************
// Copyright © 2020 Steven M Cohn.  All rights reserved.
//************************************************************************************************

namespace ClipboardRtfToHtml
{
	using System;
	using System.IO;
	using System.Linq;
	using System.Text;
	using System.Threading;
    using System.Windows;
	using System.Windows.Controls;
	using System.Windows.Documents;
	using System.Xml;
	using System.Xml.Linq;

	// Added references:
	// - PresentationCore
	// - PresentationFramework
	// - System.Windows.Forms
	// - System.ValueType - NuGet

	class Program
	{
		// https://www.freeclipboardviewer.com/

		private static readonly char Space = '\u00a0'; // Unicode no-break space


		static void Main(string[] args)
		{
			new Program().Run(args);
		}

		private void Run(string[] args)
		{
			Console.WriteLine(new string('-', 90));
			DumpClipboard();

			if (args.Contains("--dump"))
			{
				return;
			}

			PrepareClipboard();

			Console.WriteLine();
			ConsoleWriteLine("RESULT " + new string('-', 80), ConsoleColor.Green);

			DumpClipboard(true);
		}

		private void DumpClipboard(bool fancy = false)
		{
			// create STA context
			var thread = new Thread(() =>
			{
				if (Clipboard.ContainsText(TextDataFormat.Html))
					DumpContent(Clipboard.GetText(TextDataFormat.Html),
						TextDataFormat.Html, "CLIPBOARD", fancy);

				if (Clipboard.ContainsText(TextDataFormat.Rtf))
					DumpContent(Clipboard.GetText(TextDataFormat.Rtf),
						TextDataFormat.Rtf, "CLIPBOARD");

				if (Clipboard.ContainsText(TextDataFormat.Xaml))
					DumpContent(Clipboard.GetText(TextDataFormat.Xaml),
						TextDataFormat.Xaml, "CLIPBOARD", fancy);

				if (Clipboard.ContainsText(TextDataFormat.CommaSeparatedValue))
					DumpContent(Clipboard.GetText(TextDataFormat.CommaSeparatedValue),
						TextDataFormat.CommaSeparatedValue, "CLIPBOARD");

				if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
					DumpContent(Clipboard.GetText(TextDataFormat.UnicodeText),
						TextDataFormat.UnicodeText, "CLIPBOARD");

				if (Clipboard.ContainsText(TextDataFormat.Text))
					DumpContent(Clipboard.GetText(TextDataFormat.Text),
						TextDataFormat.Text, "CLIPBOARD");

				DumpDataObject();
			});

			thread.SetApartmentState(ApartmentState.STA);
			thread.Start();
			thread.Join();

			ConsoleWriteLine(new string('-', 90), ConsoleColor.DarkGray);
			Console.WriteLine();
		}

		private void DumpContent(
			string content, TextDataFormat format, string title = "CONTENT", bool fancy = false)
		{
			string preamble = null;

			if (content.StartsWith("Version:"))
			{
				var start = content.IndexOf('<');
				preamble = content.Substring(0, start);

				// trim html preamble if it's present
				content = content.Substring(start);
			}

			if (fancy)
			{
				content = XElement.Parse(content).ToString(SaveOptions.None);
			}

			ConsoleWrite($"{title}: ({format.ToString()}) [", ConsoleColor.Yellow);

			if (preamble != null)
			{
				ConsoleWrite(preamble, ConsoleColor.DarkGray);
			}

			ConsoleWrite(content, ConsoleColor.DarkGray);
			ConsoleWriteLine("]", ConsoleColor.DarkYellow);
			Console.WriteLine();
		}


		private void DumpDataObject()
		{
			var data = Clipboard.GetDataObject();

			ConsoleWrite("DataObject.Formats=[", ConsoleColor.Yellow);
			var formats = data.GetFormats();
			ConsoleWrite(string.Join(", ", formats), ConsoleColor.DarkGray);
			ConsoleWriteLine("]", ConsoleColor.Yellow);
		}


		private void PrepareClipboard()
		{
			// OneNote runs in MTA context but Clipboard required STA context
			var thread = new Thread(() =>
			{
				if (Clipboard.ContainsText(TextDataFormat.Rtf))
				{
					var text = AddHtmlPreamble(
						ConvertXamlToHtml(
							ConvertRtfToXaml(Clipboard.GetText(TextDataFormat.Rtf))));

					RebuildClipboard(text);
					ConsoleWriteLine("... Rtf -> Html", ConsoleColor.Red);
				}
				else if (Clipboard.ContainsText(TextDataFormat.Xaml))
				{
					var text = AddHtmlPreamble(
						ConvertXamlToHtml(
							Clipboard.GetText(TextDataFormat.Xaml)));

					RebuildClipboard(text);
					ConsoleWriteLine("... Xaml -> Html", ConsoleColor.Red);
				}
				else
				{
					var formats = string.Join(",", Clipboard.GetDataObject().GetFormats(false));
					ConsoleWriteLine($"... saving {formats} content", ConsoleColor.Green);
				}
			});

			thread.SetApartmentState(ApartmentState.STA);
			thread.Start();
			thread.Join();
		}


		private void RebuildClipboard(string text)
		{
			var dob = new DataObject();

			// replace any Html with our own
			dob.SetText(text, TextDataFormat.Html);

			// keep Unicode
			if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
			{
				dob.SetText(
					Clipboard.GetText(TextDataFormat.UnicodeText), TextDataFormat.UnicodeText);
			}

			// keep Text
			if (Clipboard.ContainsText(TextDataFormat.Text))
			{
				dob.SetText(
					Clipboard.GetText(TextDataFormat.Text), TextDataFormat.Text);
			}

			// replace clipboard contents, maybe locked so retry if fail
					Clipboard.SetDataObject(dob, true);
		}


		// called from STA context
		private string ConvertRtfToXaml(string rtf)
		{
			if (string.IsNullOrEmpty(rtf))
			{
				return string.Empty;
			}

			// use the old RichTextBox trick to convert RTF to Xaml
			// we'll convert the Xaml to HTML later...

			var box = new RichTextBox();
			var range = new TextRange(box.Document.ContentStart, box.Document.ContentEnd);

			// store RTF in memory stream and load into rich text box
			var istream = new MemoryStream(); // will be disposed by the following StreamWriter
			using (var writer = new StreamWriter(istream))
			{
				writer.Write(rtf);
				writer.Flush();
				istream.Seek(0, SeekOrigin.Begin);
				range.Load(istream, DataFormats.Rtf);
			}

			// read Xaml from rich text box
			using (var stream = new MemoryStream())
			{
				range = new TextRange(box.Document.ContentStart, box.Document.ContentEnd);
				range.Save(stream, DataFormats.Xaml);
				stream.Seek(0, SeekOrigin.Begin);
				using (var reader = new StreamReader(stream))
				{
					return reader.ReadToEnd();
				}
			}
		}


		private string ConvertXamlToHtml(string xaml)
		{
			if (string.IsNullOrEmpty(xaml))
			{
				ConsoleWriteLine("xaml is empty", ConsoleColor.Cyan);
				return string.Empty;
			}

			ConsoleWrite("xaml=[", ConsoleColor.Cyan);
			ConsoleWrite(xaml, ConsoleColor.DarkCyan);
			ConsoleWriteLine("]", ConsoleColor.Cyan);
			Console.WriteLine();

			var root = XElement.Parse(xaml);
			ConsoleWrite("root=[", ConsoleColor.Blue);
			ConsoleWrite(root.ToString(SaveOptions.None), ConsoleColor.DarkBlue);
			ConsoleWriteLine("]", ConsoleColor.Blue);
			Console.WriteLine();


			var builder = new StringBuilder(); 
			builder.AppendLine("<html>");
			builder.AppendLine("<body>");

			using (var outer = new XmlTextReader(new StringReader(xaml)))
			{
				// skip outer <Section> to get to subtree
				while (outer.Read() && outer.NodeType != XmlNodeType.Element) { /**/ }
				if (!outer.EOF)
				{
					using (var writer = new XmlTextWriter(new StringWriter(builder)))
					{
						// prepare proper HTML clipboard skeleton

						writer.WriteComment("StartFragment");

						using (var reader = outer.ReadSubtree())
						{
							ConvertXaml(reader, writer);
						}

						writer.WriteComment("EndFragment");
					}
				}
			}
			builder.AppendLine();
			builder.AppendLine("</body>");
			builder.AppendLine("</html>");

			return builder.ToString();
		}


		private void ConvertXaml(XmlReader reader, XmlTextWriter writer)
		{
			while (reader.Read())
			{
				switch (reader.NodeType)
				{
					case XmlNodeType.Element:
						if (!reader.IsEmptyElement)
						{
							var n = TranslateElementName(reader.Name, reader);
							writer.WriteStartElement(n);

							Report("element", reader.Name, n, reader.HasAttributes);

							if (reader.HasAttributes)
							{
								TranslateAttributes(reader, writer);
							}
						}
						break;

					case XmlNodeType.EndElement:
						{
							var n = TranslateElementName(reader.Name);
							writer.WriteEndElement();
							Report("endelement", reader.Name, n);
						}
						break;

					case XmlNodeType.CDATA:
						{
							var t = Untabify(reader.Value);
							writer.WriteCData(t);
							Report("cdata", t);
						}
						break;

					case XmlNodeType.Text:
						{
							var t = Untabify(reader.Value);
							writer.WriteValue(t);
							Report("text", t);
						}
						break;

					case XmlNodeType.SignificantWhitespace:
						{
							var t = Untabify(reader.Value);
							writer.WriteValue(t);
							Report("whitespace!", t);
						}
						break;

					case XmlNodeType.Whitespace:
						if (reader.XmlSpace == XmlSpace.Preserve)
						{
							var t = Untabify(reader.Value);
							writer.WriteValue(t);
							Report("whitespace", t);
						}
						break;

					default:
						// ignore
						Report("**", $"'{reader.NodeType}'");
						break;
				}
			}

			Console.WriteLine();
		}

		private void Report(string title, string tag1, string tag2 = null, bool hasAttributes = false)
		{
			if (hasAttributes) title += '+';

			ConsoleWrite($"{title,-12}", ConsoleColor.Blue);

			if (tag2 == null) // tag1 is a value
			{
				tag1 = tag1.Replace(" ", "·").Replace("\t", "→");
				ConsoleWrite("[", ConsoleColor.Yellow);
				ConsoleWrite(tag1, ConsoleColor.DarkCyan);
				ConsoleWrite("]", ConsoleColor.Yellow);
			}
			else
			{
				ConsoleWrite($"{tag1,-10}", ConsoleColor.Cyan);
				ConsoleWrite($"--> {tag2,-10}", ConsoleColor.Green);
			}

			Console.WriteLine();
		}


		private string Untabify(string text)
		{
			if (text == null)
				return string.Empty;

			if (text.Length == 0 || !char.IsWhiteSpace(text[0]))
				return text;

			var builder = new StringBuilder();

			int i = 0;

			while ((i < text.Length) && (text[i] == ' ' || text[i] == '\t'))
			{
				if (text[i] == ' ')
				{
					builder.Append(Space);
				}
				else if (text[i] == '\t')
				{
					do
					{
						builder.Append(Space);
					}
					while (builder.Length % 4 != 0);
				}

				i++;
			}

			while (i < text.Length)
			{
				builder.Append(text[i]);
				i++;
			}

			var t1 = text.Replace(' ', '.').Replace('\t', '_');
			var t2 = builder.ToString().Replace(Space, '·');
			ConsoleWriteLine($"... untabified [{t1}] to [{t2}]", ConsoleColor.DarkGray);

			return builder.ToString();
		}


		private string TranslateElementName(string xname, XmlReader reader = null)
		{
			string name;

			switch (xname)
			{
				case "InlineUIContainer":
				case "Span":
					name = "span";
					break;

				case "Run":
					name = "span";
					break;

				case "Bold":
					name = "b";
					break;

				case "Italic":
					name = "i";
					break;

				case "Paragraph":
					name = "p";
					break;

				case "BlockUIContainer":
				case "Section":
					name = "div";
					break;

				case "Table":
					name = "table";
					break;

				case "TableColumn":
					name = "col";
					break;

				case "TableRowGroup":
					name = "tbody";
					break;

				case "TableRow":
					name = "tr";
					break;

				case "TableCell":
					name = "td";
					break;

				case "List":
					switch (reader.GetAttribute("MarkerStyle"))
					{
						case null:
						case "None":
						case "Disc":
						case "Circle":
						case "Square":
						case "Box":
							name = "ul";
							break;

						default:
							name = "ol";
							break;
					}
					break;

				case "ListItem":
					name = "li";
					break;

				case "Hyperlink":
					name = "a";
					break;

				default:
					// ignore
					name = null;
					break;
			}

			return name;
		}


		private void TranslateAttributes(XmlReader reader, XmlTextWriter writer)
		{
			var styles = new StringBuilder();

			while (reader.MoveToNextAttribute())
			{
				switch (reader.Name)
				{
					// character formatting

					case "Background":
						styles.Append($"background-color:{ConvertColor(reader.Value)};");
						break;

					case "FontFamily":
						styles.Append($"font-family:'{reader.Value}';");
						break;

					case "FontStyle":
						styles.Append($"font-style:{reader.Value.ToLower()};");
						break;

					case "FontWeight":
						styles.Append($"font-weight:{reader.Value.ToLower()};");
						break;

					case "FontSize":
						styles.Append($"font-size:{ConvertSize(reader.Value, "pt")};");
						break;

					case "Foreground":
						styles.Append($"color:{ConvertColor(reader.Value)};");
						break;

					case "TextDecorations":
						if (reader.Value.ToLower() == "strikethrough")
							styles.Append("text-decoration:line-through;");
						else
							styles.Append("text-decoration:underline;");
						break;

					// paragraph formatting

					case "Padding":
						styles.Append($"padding:{ConvertSize(reader.Value, "px")};");
						break;

					case "Margin":
						styles.Append($"margin:{ConvertSize(reader.Value, "px")};");
						break;

					case "BorderThickness":
						styles.Append($"border-width:{ConvertSize(reader.Value, "px")};");
						break;

					case "BorderBrush":
						styles.Append($"border-color:{ConvertColor(reader.Value)};");
						break;

					case "TextIndent":
						styles.Append($"text-indent:{reader.Value};");
						break;

					case "TextAlignment":
						styles.Append($"text-align:{reader.Value.ToLower()};");
						break;

					// hyperlink attributes

					case "NavigateUri":
						writer.WriteAttributeString("href", reader.Value);
						break;

					case "TargetName":
						writer.WriteAttributeString("target", reader.Value);
						break;

					// table attributes

					case "Width":
						styles.Append($"width:{reader.Value};");
						break;

					case "ColumnSpan":
						writer.WriteAttributeString("colspan", reader.Value);
						break;

					case "RowSpan":
						writer.WriteAttributeString("rowspan", reader.Value);
						break;
				}
			}

			if (styles.Length > 0)
			{
				writer.WriteAttributeString("style", styles.ToString());
			}

			// move back to element
			reader.MoveToElement();
		}


		private string ConvertColor(string color)
		{
			// Xaml colors are /#[A-F0-9]{8}/
			if (color.Length == 9 && color.StartsWith("#"))
			{
				return "#" + color.Substring(3);
			}

			return color;
		}


		private string ConvertSize(string size, string units = null)
		{
			var parts = size.Split(',');

			for (int i = 0; i < parts.Length; i++)
			{
				if (double.TryParse(parts[i], out var value))
				{
					parts[i] = Math.Ceiling(value).ToString();
				}
				else
				{
					parts[i] = "0";
				}
			}

			var builder = new StringBuilder();
			for (int i = 0; i < parts.Length; i++)
			{
				builder.Append(parts[i]);
				builder.Append(units);

				if (i < parts.Length - 1)
				{
					builder.Append(" ");
				}
			}

			return builder.ToString();
		}


		public string AddHtmlPreamble(string html)
		{
			/*
			 * https://docs.microsoft.com/en-us/windows/win32/dataxchg/html-clipboard-format
			 * 
			 * StartHTML:00071
			 * EndHTML:00170
			 * StartFragment:00140
			 * EndFragment:00160
			 * <html>
			 * <body>
			 * <!--StartFragment--> ... <!--EndFragment-->
			 * </body>
			 * </html>
			 */

			var builder = new StringBuilder();
			builder.AppendLine("Version:0.9");
			builder.AppendLine("StartHTML:0000000000");
			builder.AppendLine("EndHTML:1111111111");
			builder.AppendLine("StartFragment:2222222222");
			builder.AppendLine("EndFragment:3333333333");

			// calculate offsets, accounting for Unicode no-break space chars

			builder.Replace("0000000000", builder.Length.ToString("D10"));

			int start = html.IndexOf("<!--StartFragment-->");
			int spaces = 0;
			for (int i = 0; i < start; i++)
			{
				if (html[i] == Space)
				{
					spaces++;
				}
			}
			builder.Replace("2222222222", (builder.Length + start + 20 + spaces).ToString("D10"));

			int end = html.IndexOf("<!--EndFragment-->");
			for (int i = start + 20; i < end; i++)
			{
				if (html[i] == Space)
				{
					spaces++;
				}
			}
			spaces--;
			builder.Replace("3333333333", (builder.Length + end + spaces).ToString("D10"));
			builder.Replace("1111111111", (builder.Length + html.Length + spaces).ToString("D10"));

			builder.AppendLine(html);
			return builder.ToString();
		}


		internal static void ConsoleWrite(string text, ConsoleColor color)
		{
			var save = Console.ForegroundColor;
			Console.ForegroundColor = color;
			Console.Write(text);
			Console.ForegroundColor = save;
		}

		internal static void ConsoleWriteLine(string text, ConsoleColor color)
		{
			var save = Console.ForegroundColor;
			Console.ForegroundColor = color;
			Console.WriteLine(text);
			Console.ForegroundColor = save;
		}
	}
}
