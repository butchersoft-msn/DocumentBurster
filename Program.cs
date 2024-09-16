using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Path = System.IO.Path;
using Body = DocumentFormat.OpenXml.Wordprocessing.Body;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DocumentSplitter
{
    public class Program
    {
        // Extract Bullet Points
        public static string fileNameIn = string.Empty;
        public static string filePathIn = string.Empty;
        public static string filePathOut = string.Empty;

        public static List<Paragraph> tocParagraphs = new List<Paragraph>();
        public static List<string> tocStyles = new List<string>();

        //Regex rx = new Regex(@"[A-Z][0-9].*\W", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static void Main(string[] args)
        {

            if (args.Length > 0)                
            {
                fileNameIn = args[0];

                if (File.Exists(fileNameIn))
                {                 
                    filePathIn = Path.GetDirectoryName(fileNameIn);                    
                    filePathOut = $"{filePathIn}\\output";

                    Directory.CreateDirectory(filePathOut);
                }

                Console.WriteLine($"Apply Dcument Burster to {fileNameIn}");
                ExtractContents(fileNameIn);
            }
            else
            {
                Console.WriteLine($"file not found, try including the fullpath {fileNameIn}");
            }                
        }

        public static void ExtractContents(string fileNameIn)
        {
            string filePath = fileNameIn;
            string paragraphOutputFile = $"{filePathOut}\\{Path.GetFileNameWithoutExtension(fileNameIn)}.html"; // File to save paragraph text

            // Ensure the directory exists for saving output files
            Directory.CreateDirectory($"{filePathOut}\\paragraphs");
            Directory.CreateDirectory($"{filePathOut}\\images");
            Directory.CreateDirectory($"{filePathOut}\\tables");

            using (StreamWriter htmlWriter = new StreamWriter(paragraphOutputFile, false)) // Prepare to write paragraphs to file
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
            {
                Body body = wordDoc.MainDocumentPart.Document.Body;

                int imageCount = 0;
                int currentLevel = 0;
                Stack<int> listStack = new Stack<int>();

                // Extract document properties
                var properties = wordDoc.PackageProperties;

                // Convert Styling to styles.css
                ConvertStylesToCSS(wordDoc.MainDocumentPart.StyleDefinitionsPart);

                // Start writing the HTML document with <head> and the CSS link
                htmlWriter.WriteLine("<html>");
                htmlWriter.WriteLine("<head>");
                htmlWriter.WriteLine("<link href=\"https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&family=Merriweather:wght@400;700&display=swap\" rel=\"stylesheet\">");
                htmlWriter.WriteLine($"<link rel=\"stylesheet\" type=\"text/css\" href=\"style.css\">"); // Link to external CSS
                htmlWriter.WriteLine($"<meta name=\"title\" content=\"{properties.Title}\">");
                htmlWriter.WriteLine($"<meta name=\"author\" content=\"{properties.Creator}\">");
                htmlWriter.WriteLine($"<meta name=\"subject\" content=\"{properties.Subject}\">");
                htmlWriter.WriteLine($"<meta name=\"keywords\" content=\"{properties.Keywords}\">");
                htmlWriter.WriteLine($"<meta name=\"last-modified-by\" content=\"{properties.LastModifiedBy}\">");
                htmlWriter.WriteLine($"<meta name=\"created\" content=\"{properties.Created}\">");
                htmlWriter.WriteLine($"<meta name=\"modified\" content=\"{properties.Modified}\">");
                htmlWriter.WriteLine($"<meta name=\"source\" content=\"{fileNameIn}\">");
                htmlWriter.WriteLine("</head>");
                htmlWriter.WriteLine("<body style=\"font-family: sans-serif;font-size:11px\">");

                // Add container for two-column layout
                htmlWriter.WriteLine("<div class=\"container\">");
                htmlWriter.WriteLine("<div class=\"toc-column\" style=\"flex: 0 1 75%; width:24%\">");

                // Add a button to download the original document
                htmlWriter.WriteLine("<div>");
                htmlWriter.WriteLine($"<button class=\"btn\" type=\"submit\" onclick=\"window.open('{fileNameIn.Replace("\\", "/")}')\">Download Document</button>");
                htmlWriter.WriteLine("</div>");

                //PARSE TABLE OF CONTENTS
                htmlWriter.WriteLine("<h2>Table of Contents</h2>");
                
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph paragraph)
                    {
                        string pStyle = GetParagraphStyle(paragraph);
                        string indentation = GetIndentation(paragraph);
                        var texts = paragraph.Descendants<Text>().Select(t => t.Text).ToList();
                        //Remove trailing page numbers
                        string paragraphText = string.Join(" ", texts).TrimEnd(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }); ;

                        if (!string.IsNullOrWhiteSpace(paragraphText))
                        {
                            if (pStyle.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                            {
                                int level = Convert.ToInt32(pStyle.Remove(0, 3));
                                // Handle list hierarchy
                                if (level > currentLevel)
                                {
                                    for (int i = currentLevel; i < level; i++)
                                    {
                                        htmlWriter.WriteLine("<ol style=\"list-style-type: none;padding-left: 10;\">");
                                        listStack.Push(i);
                                    }
                                }
                                else if (level < currentLevel)
                                {
                                    for (int i = currentLevel; i > level; i--)
                                    {
                                        htmlWriter.WriteLine("</ol>");
                                        listStack.Pop();
                                    }
                                }

                                var refLink = GetHyperLink(paragraph);
                                string refAttribute = refLink != null ? refLink : string.Empty;

                                htmlWriter.WriteLine($"<li><a {refAttribute} target=\"i_bookmarks\" >{paragraphText}</a></li>");

                                // Write the list item
                                currentLevel = level;
                            }
                        }
                    }
                }

                // Close any remaining open lists
                while (listStack.Count > 0)
                {
                    htmlWriter.WriteLine("</ol>");
                    listStack.Pop();
                }

                // Close TOC column
                htmlWriter.WriteLine("</div>"); 

                // Start content column
                htmlWriter.WriteLine("<div class=\"content-column\" style=\"position: absolute; left: 25%; top: 0; bottom: 0; right: 0; overflow: auto;\";  >");

                //#####################################################################################################################
                //Export BookMark Sections as HTML pages
                var bookmarks = GetBookmarks(body);
                var bookmarkSelect = String.Empty;
                var bookmarkPos = 0;
                List<BookmarkInfo> bookmarkElements = new List<BookmarkInfo>();
                List<OpenXmlElement> sectionElements = new List<OpenXmlElement>();

                
                //Loop through all Elements of Document and Group by Bookmark Name
                foreach (var element in body.Elements())
                {
                    if (element is Paragraph paragraph)
                    {
                        // Find the first BookmarkStart element within the paragraph
                        BookmarkStart bookmarkStart = paragraph.Descendants<BookmarkStart>().LastOrDefault();

                        if (bookmarkStart != null)                            
                        {
                            if (bookmarkStart.Name == bookmarks[bookmarkPos].ToString())
                            {
                                if (bookmarkSelect != null)
                                {                                    
                                    var bookmarkInfo = new BookmarkInfo() { BookmarkName = bookmarkStart.Name, Elements = sectionElements };
                                    bookmarkElements.Add(bookmarkInfo);
                                }

                                bookmarkSelect = bookmarkStart.Name;
                                bookmarkPos++;
                                sectionElements = new List<OpenXmlElement>();
                            }
                                
                        }    
                    }

                    if(bookmarkSelect != null)
                    {
                        sectionElements.Add(element);
                    }

                }

                //Loop through all bookmarked sections
                foreach (var bookmarkElement in bookmarkElements)
                {
                    //Export bookmarkElements to html
                    StringBuilder bookmarkContentBuilder = new StringBuilder();

                    // Start writing the HTML document with <head> and the CSS link
                    bookmarkContentBuilder.AppendLine("<html>");
                    bookmarkContentBuilder.AppendLine("<head>");
                    bookmarkContentBuilder.AppendLine("<link href=\"https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&family=Merriweather:wght@400;700&display=swap\" rel=\"stylesheet\">");
                    bookmarkContentBuilder.AppendLine($"<link rel=\"stylesheet\" type=\"text/css\" href=\"style.css\">"); // Link to external CSS
                    bookmarkContentBuilder.AppendLine("</head>");
                    bookmarkContentBuilder.AppendLine("<body>");

                    //WRITE BOOKMARKS TO PAGES
                    foreach (var element in bookmarkElement.Elements)
                    {
                        // Convert paragraph to HTML
                        if (element is Paragraph paragraph)
                        {
                            var pStyle = GetParagraphStyle(paragraph);
                            if (!pStyle.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
                            {
                                // Write paragraph HTML to the bookmark content
                                var texts = paragraph.Descendants<Text>().Select(t => t.Text).ToList();
                                string paragraphText = string.Join(" ", texts);
                                var paragraphHtml = $"<p class=\"{pStyle}\">{paragraphText}</p>";
                                bookmarkContentBuilder.AppendLine(paragraphHtml);
                            }

                            var drawingElements = paragraph.Descendants<Drawing>();
                            foreach (var drawing in drawingElements)
                            {
                                imageCount++;
                                string imageFileName = ExtractImageFromDrawing(drawing, wordDoc, imageCount);
                                if (!string.IsNullOrEmpty(imageFileName))
                                {
                                    string imgTag = $"<img src=\"../images/{imageFileName}\" alt=\"Image\" />";
                                    bookmarkContentBuilder.AppendLine(imgTag);
                                }
                            }
                        }

                        // Convert table to HTML
                        if (element is Table table)
                        {
                            string tableHtml = ConvertTableToHtml(table);
                            bookmarkContentBuilder.AppendLine(tableHtml);
                        }

                        // Convert images (if applicable)
                        if (element is Drawing image)
                        {
                            var drawingElements = element.Descendants<Drawing>();
                            imageCount = 0;
                            foreach (var drawing in drawingElements)
                            {
                                imageCount++;
                                string imageFileName = ExtractImageFromDrawing(drawing, wordDoc, imageCount);
                                if (!string.IsNullOrEmpty(imageFileName))
                                {
                                    string imageHtml = $"<img src=\"images/{imageFileName}\" alt=\"Image\" />";
                                    bookmarkContentBuilder.AppendLine(imageHtml);
                                }
                            }

                        }
    
                    }

                    bookmarkContentBuilder.AppendLine("</body>");
                    bookmarkContentBuilder.AppendLine("</html>");

                    File.WriteAllText($"{filePathOut}\\paragraphs\\{bookmarkElement.BookmarkName}.html", bookmarkContentBuilder.ToString());

                }

                Console.WriteLine("HTML extraction completed.");

                htmlWriter.WriteLine($"<iframe src=\"paragraphs\\{bookmarkElements[0].BookmarkName}.html\" name=\"i_bookmarks\" width=\"100%\" style=\"height: -webkit-fill-available ;\"></iframe>");
                htmlWriter.WriteLine("</div>"); // Close content column
                htmlWriter.WriteLine("</div>"); // Close container

                htmlWriter.WriteLine("</body>");
                htmlWriter.WriteLine("</html>");

            }
        }

        private static string ConvertLinksToHtml(string text)
        {
            // Regular expression to match HTTP and HTTPS URLs
            string urlPattern = @"(http[s]?:\/\/[^\s]+)";

            // Replace all matches with an HTML anchor tag
            return Regex.Replace(text, urlPattern, match =>
            {
                string url = match.Value;
                return $"<a href=\"{url}\">{url}</a>";
            });
        }


        /// <summary>
        /// List of Elements found in each bookmark
        /// </summary>
        public class BookmarkInfo
        {
            public string BookmarkName { get; set; }
            public List<OpenXmlElement> Elements { get; set; }
        }


        /// Extract the style from Paragraph properties
        public static string GetParagraphStyle(Paragraph paragraph)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null && paragraphProperties.ParagraphStyleId != null)
            {
                return paragraphProperties.ParagraphStyleId.Val;
            }
            return "normal"; // Default to "normal" if no style is found
        }

        // Extract indentation from Paragraph properties
        public static string GetIndentation(Paragraph paragraph)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties != null)
            {
                var indentation = paragraphProperties.Indentation;
                if (indentation != null)
                {
                    // Convert indentation values to CSS 'text-indent' value
                    // Default values
                    int indentLeft = 0;
                    int indentRight = 0;

                    if (indentation.Left != null)
                    {
                        indentLeft = (int)(Convert.ToInt32(indentation.Left.Value) / 240); // Convert from EMUs to mm
                    }
                    if (indentation.Right != null)
                    {
                        indentRight = (int)(Convert.ToInt32(indentation.Right.Value) / 240); // Convert from EMUs to mm
                    }

                    return $"{indentLeft}px"; // You may adjust this based on your requirements
                }
            }
            return "0px"; // No indentation by default
        }

        // Extract images from drawing
        public static string ExtractImageFromDrawing(Drawing drawing, WordprocessingDocument wordDoc, int imageCount)
        {
            var blip = drawing.Descendants<Blip>().FirstOrDefault();
            if (blip != null)
            {
                string embed = blip.Embed;
                var imagePart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(embed);

                if (imagePart != null)
                {
                    string imageFileName = $"image_{imageCount}.png"; // Image name in output folder
                    string imageFilePath = Path.Combine($"{filePathIn}\\output\\images", imageFileName);

                    using (var stream = imagePart.GetStream())
                    {
                        using (var fileStream = new FileStream(imageFilePath, FileMode.Create))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }

                    Console.WriteLine($"Image saved as {imageFileName}");
                    return imageFileName; // Return image filename to insert into HTML
                }
            }
            return null;
        }

        // Convert table to HTML and add bookmark IDs to <td> tags if available
        public static string ConvertTableToHtml(Table table)
        {
            string html = "<table border='1' style='border-collapse: collapse;'>\n";

            foreach (var row in table.Elements<TableRow>())
            {
                html += "<tr>\n";

                foreach (var cell in row.Elements<TableCell>())
                {
                    // Extract the pStyle from the cell's first paragraph, if available
                    var firstParagraph = cell.Descendants<Paragraph>().FirstOrDefault();
                    string cellStyle = firstParagraph != null ? GetParagraphStyle(firstParagraph) : "normal";

                    // Check if a bookmark exists within this cell and use it as the id

                    string cellText = cell.InnerText.Replace("\\s", "").Replace("\\t", "").Replace("\"", "").Replace("AutoTextList", "").Replace("NoStyle", "");

                    // Generate CSS for the cell from the paragraph (if any)
                    string inlineStyle = GetCellInlineStyle(firstParagraph);

                    // Extract cell background color
                    string backgroundColor = GetCellBackgroundColor(cell);
                    if (!string.IsNullOrEmpty(backgroundColor))
                    {
                        inlineStyle += $"background-color: {backgroundColor}; ";
                    }

                    var bookMark = GetBookmarkName(cell.Descendants<Paragraph>().FirstOrDefault());
                    string idAttribute = bookMark != null ? $" id=\"{bookMark}\"" : string.Empty;

                    html += $"<td {idAttribute} class=\"{cellStyle}\"{idAttribute} style=\"{inlineStyle}\">{cellText}</td>\n";
                }

                html += "</tr>\n";
            }

            html += "</table>\n";
            return html;
        }


        private static string GetCellInlineStyle(Paragraph paragraph)
        {
            if (paragraph == null) return string.Empty;

            // Initialize a style string
            string inlineStyle = "";

            // Check for font colur
            if (paragraph.Descendants<RunProperties>().Any(rp => rp.Color != null && rp.Color.Val != null))
            {
                var fontColor = paragraph.Descendants<RunProperties>().Select(rp => rp.Color.Val)?.FirstOrDefault();
                inlineStyle += $" color: {fontColor}; ";
            }

            // Check for bold text
            if (paragraph.Descendants<RunProperties>().Any(rp => rp.Bold != null && rp.Bold.Val != null))
            {
                inlineStyle += "font-weight: bold; ";
            }

            // Check for italic text
            if (paragraph.Descendants<RunProperties>().Any(rp => rp.Italic != null && rp.Italic.Val !=null))
            {
                inlineStyle += "font-style: italic; ";
            }

            // Check for font size
            var fontSize = paragraph.Descendants<RunProperties>().Select(rp => rp.FontSize?.Val).FirstOrDefault();
            if (fontSize != null)
            {
                inlineStyle += $"font-size: {fontSize}pt; ";
            }

            // Handle text alignment
            var alignment = paragraph.ParagraphProperties?.Justification?.Val.ToString().ToLower();
            if (alignment != null)
            {

                string textAlign = alignment switch
                {
                    "center" => "center",
                    "left" => "left",
                    "right" => "right",
                    "both" => "justify",
                    _ => "left"
                };
                inlineStyle += $"  text-align: {textAlign};";
            }

            return inlineStyle;
        }

        /// <summary>
        /// Return the background color of a table cell
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellBackgroundColor(TableCell cell)
        {
            // Check if the cell has a background color (shading)
            var cellShading = cell.TableCellProperties?.Shading;
            if (cellShading != null && cellShading.Fill != null)
            {
                // Extract the hex value of the background color
                return $"#{cellShading.Fill}";
            }

            // Return an empty string if no background color is found
            return string.Empty;
        }

        /// <summary>
        /// Look for bookmark in Pararaph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public static string GetBookmarkName(Paragraph paragraph)
        {
            var bookmarkStart = paragraph.Descendants<BookmarkStart>().LastOrDefault();

            if (bookmarkStart != null)
            {
                return bookmarkStart.Name; // Returns the name of the bookmark
            }

            return null; // No bookmark found
        }

        /// <summary>
        /// Extract Hyperlink from TOC Fields
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public static string GetHyperLink(Paragraph paragraph)
        {
            // Extract hyperlink if present
            var hyperlink = paragraph.Descendants<Hyperlink>().FirstOrDefault();
            string hyperlinkId = string.Empty;
            if (hyperlink != null)
            {
                var hyperlinkRef = hyperlink.InnerText.Split("PAGEREF");
                return $"href=\"paragraphs\\{hyperlinkRef[1].Split(" ")[1]}.html\"";

            }
            return null;
        }

        /// <summary>
        /// Extract all bookmarks in the document
        /// </summary>
        /// <param name="body"></param>
        /// <returns> list of bookmarks key values</returns>
        public static List<string> GetBookmarks(Body body)
        {
            List<string> bookmarks = new List<string>();

            var bookmarkAnchors = body.Descendants<BookmarkStart>().Where(p => p.Name.ToString().StartsWith("_T"));

            foreach (var bookmarkStart in bookmarkAnchors)
            {
                if (!string.IsNullOrWhiteSpace(bookmarkStart.Name))
                {
                    bookmarks.Add(bookmarkStart.Name);                    
                }
            }

            return bookmarks;
        }


        /// <summary>
        /// Counts the number of parents to decide how many levels of indent the heading has
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public static int GetParagraphDepth(Paragraph paragraph)
        {
            int depth = 0;
            var current = paragraph.Parent;

            // Traverse up the hierarchy until the top parent (Parent == null)
            while (current.Parent != null)
            {
                current = current.Parent;
                depth++;
            }

            return depth;
        }


        public static List<OpenXmlElement> GetBookMarkElements(Body body, string bookmarkName)
        {
            // Collect elements between BookmarkStart and BookmarkEnd
            List<OpenXmlElement> elementsBetween = new List<OpenXmlElement>();

            // Find the BookmarkStart element
            BookmarkStart bookmarkStart = body.Descendants<BookmarkStart>()
                                                .FirstOrDefault(b => b.Name == bookmarkName);

            // Find the BookmarkEnd element with matching Id
            if (bookmarkStart != null)
            {
                BookmarkEnd bookmarkEnd = body.Descendants<BookmarkEnd>()
                                                .FirstOrDefault(be => be.Id == bookmarkStart.Id);

                if (bookmarkEnd != null)
                {

                    bool isWithinBookmark = false;

                    foreach (var element in body.Elements())
                    {
                        if (element == bookmarkStart)
                        {
                            isWithinBookmark = true;
                            continue;
                        }

                        if (element == bookmarkEnd)
                        {
                            break;
                        }

                        if (isWithinBookmark)
                        {
                            elementsBetween.Add(element);
                        }
                    }

                    // Do something with the collected elements
                    foreach (var element in elementsBetween)
                    {
                        Console.WriteLine(element.OuterXml);
                    }
                }
            }

            return elementsBetween;
        }

        /// <summary>
        /// Extract all Text styles and Formatting and convert to CSS
        /// </summary>
        /// <param name="styleDefinitionsPart"></param>
        public static void ConvertStylesToCSS(StyleDefinitionsPart styleDefinitionsPart)
        {
            if (styleDefinitionsPart == null)
            {
                return; // No styles part, exit
            }


            var styles = styleDefinitionsPart.Styles.Elements<Style>();

            // Prepare the CSS output
            StringBuilder cssBuilder = new StringBuilder();

            var btn = ".btn {  background-color: DodgerBlue;border: none;color: white;padding: 12px 30px;cursor: pointer;font-size: 14px;}";
            cssBuilder.AppendLine(btn);
            var btnHover = ".btn:hover { background-color: RoyalBlue;}";
            cssBuilder.AppendLine(btnHover);
            var hyperlinkVist = "a, a:visited, a:hover, a:active { color: inherit; }";

            cssBuilder.AppendLine(hyperlinkVist);

            var oldFontName = "sans-serif";
            // Iterate over each style in the document
            foreach (var style in styles)
            {
                var styleId = style.StyleId;
                var styleName = style.StyleName?.Val;

                if (char.IsDigit(styleName,0)) {
                    styleName.Value = styleName.Value.Remove(0, 1); 
                }

                // Skip styles that don't have a name
                if (styleName == null) continue;

                // Create CSS class name based on style ID
                var cssClass = $".{styleId}";

                // Start CSS rule for this style
                cssBuilder.AppendLine($"{cssClass} {{");

                // Extract paragraph properties (e.g., alignment, spacing)
                var paragraphProperties = style.StyleParagraphProperties;
                if (paragraphProperties != null)
                {
                    // Handle text alignment
                    var alignment = paragraphProperties.Justification?.Val.ToString().ToLower();
                    if (alignment != null)
                    {

                        string textAlign = alignment switch
                        {
                            "center" => "center",
                            "left" => "left",
                            "right" => "right",
                            "both" => "justify",
                            _ => "left"
                        };
                        cssBuilder.AppendLine($"  text-align: {textAlign};");
                    }

                    // Handle spacing (before and after)
                    var spacing = paragraphProperties.SpacingBetweenLines;
                    if (spacing != null)
                    {

                        //TODO 
                        if (spacing.Before != null)
                        {
                            //cssBuilder.AppendLine($"  margin-top: {spacing.Before.Value / 20}px;"); // Convert to points
                        }
                        if (spacing.After != null)
                        {
                            //cssBuilder.AppendLine($"  margin-bottom: {spacing.After.Value / 20}px;"); // Convert to points
                        }
                    }
                }

                // Extract run properties (e.g., font, bold, italic)
                string oldfontSize = "22";
                var runProperties = style.StyleRunProperties;
                if (runProperties != null)
                {
                    // Handle font weight
                    if (runProperties.Bold != null)
                    {
                        cssBuilder.AppendLine("  font-weight: bold;");
                    }
                    if (runProperties.Italic != null)
                    {
                        cssBuilder.AppendLine("  font-style: italic;");
                    }

                    // Handle font size
                    
                    if (runProperties.FontSize != null)
                    {
                        var fontSize = runProperties.FontSize.Val.Value == null ? oldfontSize : runProperties.FontSize.Val.Value;
                        cssBuilder.AppendLine($"  font-size: {int.Parse(fontSize) / 2}px;"); // Word font size is in half-points
                        oldfontSize = fontSize;
                    }

                    // Handle font color
                    if (runProperties.Color != null && runProperties.Color.Val.ToString() !="auto")
                    {                        
                        cssBuilder.AppendLine($"  color: #{runProperties.Color.Val};");
                    }

                    

                    // Handle font family
                    if (runProperties.RunFonts != null)
                    {
                        var fontName = oldFontName;
                        if (runProperties.RunFonts?.Ascii != null)
                        {
                            fontName = runProperties.RunFonts.Ascii;
                        }

                        cssBuilder.AppendLine($" font-family: \"{fontName}\",\"Arial\";");
                        oldFontName = fontName;
                    }
                }

                // Close the CSS rule
                cssBuilder.AppendLine("}");
            }

            // Write the CSS output to a file
            File.WriteAllText($"{filePathOut}\\style.css", cssBuilder.ToString());
        }
    }
}



