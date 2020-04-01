using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;

namespace HCUR
{
    class Program
    {
        static void Main(string[] args)
        {
            //WordOpenXmlCreateDocument();
            //WordOpenXmlCreateDocumentFromStream();
            //WordOpenXmlOpenWordDocument();
            //WordOpenXmlOpenAndAddTextToDocument();
            //WordOpenXmlFindPropertiesDocument();
            //WordOpenXmlAddOneCustomPropertyDocument();
            //WordOpenXmlJoinTwoDocuments();
            //WordOpenXmlFontForRun();
            //WordOpenXmlFindStyles();
            //WordOpenXmlCreateAndApplyStyleParagraph();
            //WordOpenXmlCreateHeader();
            //WordOpenXmlCreateFooter();
            //WordOpenXmlRemoveHeadersAndFooters();
            //WordOpenXmlAddComment();
            //WordOpenXmlFindCommentsinDocument();
            //WordOpenXmlDeleteComments();
            //WordOpenXmlAddTableUsingData();
            //WordOpenXmlAddTableUsingDirectData();
            //WordOpenXmlModifyTextInCell();
            //WordOpenXmlDeleteRow();
            //WordOpenXmlAddPictureInDocument();
            //WordOpenXmlDeleteAllPicturesFromDocument();
            //WordOpenXmlAcceptAllChages();
            //WordOpenXmlRemoveHiddenText();
            //WordOpenXmlValidator();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        public static void WordOpenXmlCreateDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Create(@"C:\Temporary\WordDoc01.docx",
                                                    WordprocessingDocumentType.Document))
            {
                MainDocumentPart docMainPart = myWordDoc.AddMainDocumentPart();

                docMainPart.Document = new Document();
                Body docBody = docMainPart.Document.AppendChild(new Body());
                Paragraph docParagraph = docBody.AppendChild(new Paragraph());
                Run docRun = docParagraph.AppendChild(new Run());
                docRun.AppendChild(new Text("Text in the document"));
            }
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        public static void WordOpenXmlCreateDocumentFromStream()
        {
            FileStream myStream = File.Create(@"C:\Temporary\WordDoc01.docx");

            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Create(myStream,
                                                    WordprocessingDocumentType.Document))
            {
                MainDocumentPart docMainPart = myWordDoc.AddMainDocumentPart();

                docMainPart.Document = new Document();
                Body docBody = docMainPart.Document.AppendChild(new Body());
                Paragraph docParagraph = docBody.AppendChild(new Paragraph());
                Run docRun = docParagraph.AppendChild(new Run());
                docRun.AppendChild(new Text("Text from stream document"));
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        public static void WordOpenXmlOpenDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", false))
            {
                Body docBody = myWordDoc.MainDocumentPart.Document.Body;

                Console.WriteLine(docBody.InnerText);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        public static void WordOpenXmlOpenAndAddTextToDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                Body docBody = myWordDoc.MainDocumentPart.Document.Body;

                Paragraph docParagraph = docBody.AppendChild(new Paragraph());
                Run docRun = docParagraph.AppendChild(new Run());
                docRun.AppendChild(new Text("Text added in the document"));
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        public static void WordOpenXmlFindPropertiesDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", false))
            {
                ExtendedFilePropertiesPart docPart = myWordDoc.ExtendedFilePropertiesPart;
                DocumentFormat.OpenXml.ExtendedProperties.Properties myProps =
                                                                    docPart.Properties;

                PropertyInfo[] myPropsInfo = myProps.GetType().GetProperties();
                foreach (var oneProp in myPropsInfo)
                {
                    string propName = oneProp.Name;
                    try
                    {
                        var propValue = oneProp.GetValue(myProps) ?? "(null)";
                        Console.WriteLine(propName + " - " + propValue.ToString());
                    }
                    catch { }
                }
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        public static void WordOpenXmlAddOneCustomPropertyDocument()
        {
            string propName = "myCustomProperty";
            string returnValue = string.Empty;

            CustomDocumentProperty newProp = new CustomDocumentProperty();
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propName;

            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                newProp.VTLPWSTR = new VTLPWSTR("This is the value of the property");

                var customPropsPart = myWordDoc.CustomFilePropertiesPart;
                if (customPropsPart == null)
                {
                    customPropsPart = myWordDoc.AddCustomFilePropertiesPart();
                    customPropsPart.Properties =
                        new DocumentFormat.OpenXml.CustomProperties.Properties();
                }

                var customProps = customPropsPart.Properties;
                if (customProps != null)
                {
                    var oneProp = customProps.Where(
                        prp => ((CustomDocumentProperty)prp).Name.Value
                                            == propName).FirstOrDefault();

                    if (oneProp != null)
                    {
                        returnValue = oneProp.InnerText;
                        oneProp.Remove();
                    }

                    customProps.AppendChild(newProp);
                    int pid = 2;
                    foreach (CustomDocumentProperty item in customProps)
                    {
                        item.PropertyId = pid++;
                    }
                    customProps.Save();
                }
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        public static void WordOpenXmlJoinTwoDocuments()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                string altChunkId = "ID" + Guid.NewGuid().ToString();
                MainDocumentPart mainPart = myWordDoc.MainDocumentPart;

                AlternativeFormatImportPart myChunk =
                    mainPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                using (FileStream fileStream =
                            File.Open(@"C:\Temporary\WordDoc02.docx", FileMode.Open))

                    myChunk.FeedData(fileStream);
                AltChunk altChunk = new AltChunk
                {
                    Id = altChunkId
                };
                mainPart.Document.Body
                    .InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>()
                    .Last());
                mainPart.Document.Save();
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        public static void WordOpenXmlFontForRun()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                RunProperties runProps = new RunProperties(
                    new RunFonts() { Ascii = "Broadway" },
                    new FontSize() { Val = "28" },
                    new Bold(),
                    new Color() { Val = "FF0000" });

                Run myRun = myWordDoc.MainDocumentPart.Document.Descendants<Run>().First();
                myRun.PrependChild<RunProperties>(runProps);

                myWordDoc.MainDocumentPart.Document.Save();
            }
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        public static void WordOpenXmlFindStyles()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", false))
            {
                MainDocumentPart docPart = myWordDoc.MainDocumentPart;
                StylesPart stylesPart = docPart.StyleDefinitionsPart;

                using (XmlReader reader = XmlNodeReader.Create(
                                stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    XDocument docStyles = XDocument.Load(reader);

                    Console.WriteLine(docStyles);
                }
            }
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        public static void WordOpenXmlCreateAndApplyStyleParagraph()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;
                StyleDefinitionsPart stylePart = docMainPart.StyleDefinitionsPart;

                Style myStyle = new Style
                {
                    Type = StyleValues.Paragraph,
                    CustomStyle = true,
                    StyleId = "MyHeading1"
                };

                RunProperties runProps = new RunProperties(
                    new Name() { Val = "My Heading 1" },
                    new BasedOn() { Val = "Heading1" },
                    new NextParagraphStyle() { Val = "Normal" }
                );

                stylePart.Styles = new Styles();
                stylePart.Styles.Append(myStyle);
                stylePart.Styles.Save();

                Body docBody = docMainPart.Document.Body;
                Paragraph docParagraph = docBody.AppendChild(new Paragraph());

                ParagraphProperties paragraphProps = new ParagraphProperties();
                paragraphProps.ParagraphStyleId = new ParagraphStyleId()
                { Val = "MyHeading1" };
                docParagraph.Append(paragraphProps);

                Run docRun = docParagraph.AppendChild(new Run());
                docRun.AppendChild(new Text("This is the Heading with style"));
            }
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        public static void WordOpenXmlCreateHeader()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;

                docMainPart.DeleteParts(docMainPart.HeaderParts);
                HeaderPart headerPart = docMainPart.AddNewPart<HeaderPart>();

                GenerateHeaderPartContentFirst(headerPart);

                string headerPartId = docMainPart.GetIdOfPart(headerPart);

                IEnumerable<SectionProperties> allSections =
                                docMainPart.Document.Body.Elements<SectionProperties>();

                foreach (var oneSection in allSections)
                {
                    oneSection.RemoveAllChildren<HeaderReference>();
                    oneSection.RemoveAllChildren<TitlePage>();

                    oneSection.PrependChild<HeaderReference>(new HeaderReference()
                    { Id = headerPartId, Type = HeaderFooterValues.First });
                    oneSection.PrependChild<TitlePage>(new TitlePage());
                }
                docMainPart.Document.Save();
            }
        }

        public static void GenerateHeaderPartContentFirst(HeaderPart firstPart)
        {
            Header hFirst = new Header()
            {
                MCAttributes = new MarkupCompatibilityAttributes()
                { Ignorable = "w14 wp14" }
            };
            hFirst.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            hFirst.AddNamespaceDeclaration("mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006");
            hFirst.AddNamespaceDeclaration("o",
                "urn:schemas-microsoft-com:office:office");
            hFirst.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            hFirst.AddNamespaceDeclaration("m",
                "http://schemas.openxmlformats.org/officeDocument/2006/math");
            hFirst.AddNamespaceDeclaration("v",
                "urn:schemas-microsoft-com:vml");
            hFirst.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            hFirst.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            hFirst.AddNamespaceDeclaration("w10",
                "urn:schemas-microsoft-com:office:word");
            hFirst.AddNamespaceDeclaration("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            hFirst.AddNamespaceDeclaration("w14",
                "http://schemas.microsoft.com/office/word/2010/wordml");
            hFirst.AddNamespaceDeclaration("wpg",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            hFirst.AddNamespaceDeclaration("wpi",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            hFirst.AddNamespaceDeclaration("wne",
                "http://schemas.microsoft.com/office/word/2006/wordml");
            hFirst.AddNamespaceDeclaration("wps",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph pFirst = new Paragraph()
            {
                RsidParagraphAddition = "00164C17",
                RsidRunAdditionDefault = "00164C17"
            };
            ParagraphProperties pPropFirst = new ParagraphProperties();
            ParagraphStyleId pStyleIdFirst = new ParagraphStyleId() { Val = "Header" };

            pPropFirst.Append(pStyleIdFirst);

            Run runFirst = new Run();
            Text textFirst = new Text();
            textFirst.Text = "Header in the first page";

            runFirst.Append(textFirst);
            pFirst.Append(pPropFirst);
            pFirst.Append(runFirst);
            hFirst.Append(pFirst);
            firstPart.Header = hFirst;
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        public static void WordOpenXmlCreateFooter()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;

                docMainPart.DeleteParts(docMainPart.FooterParts);
                FooterPart defaultFooterPart = docMainPart.AddNewPart<FooterPart>();

                GenerateFooterPartContent(defaultFooterPart);

                string defaultFooterPartId = docMainPart.GetIdOfPart(defaultFooterPart);

                IEnumerable<SectionProperties> allSections =
                                docMainPart.Document.Body.Elements<SectionProperties>();

                foreach (var oneSection in allSections)
                {
                    oneSection.RemoveAllChildren<FooterReference>();
                    oneSection.RemoveAllChildren<TitlePage>();

                    oneSection.PrependChild<FooterReference>(new FooterReference()
                    { Id = defaultFooterPartId, Type = HeaderFooterValues.Default });
                    oneSection.PrependChild<TitlePage>(new TitlePage());
                }
                docMainPart.Document.Save();
            }
        }

        public static void GenerateFooterPartContent(FooterPart defaultPart)
        {
            Footer fDefault = new Footer()
            {
                MCAttributes = new MarkupCompatibilityAttributes()
                { Ignorable = "w14 wp14" }
            };
            fDefault.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            fDefault.AddNamespaceDeclaration("mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fDefault.AddNamespaceDeclaration("o",
                "urn:schemas-microsoft-com:office:office");
            fDefault.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fDefault.AddNamespaceDeclaration("m",
                "http://schemas.openxmlformats.org/officeDocument/2006/math");
            fDefault.AddNamespaceDeclaration("v",
                "urn:schemas-microsoft-com:vml");
            fDefault.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            fDefault.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            fDefault.AddNamespaceDeclaration("w10",
                "urn:schemas-microsoft-com:office:word");
            fDefault.AddNamespaceDeclaration("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fDefault.AddNamespaceDeclaration("w14",
                "http://schemas.microsoft.com/office/word/2010/wordml");
            fDefault.AddNamespaceDeclaration("wpg",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            fDefault.AddNamespaceDeclaration("wpi",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            fDefault.AddNamespaceDeclaration("wne",
                "http://schemas.microsoft.com/office/word/2006/wordml");
            fDefault.AddNamespaceDeclaration("wps",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph pDefault = new Paragraph()
            {
                RsidParagraphAddition = "00164C17",
                RsidRunAdditionDefault = "00164C17"
            };

            ParagraphProperties pPropDefault = new ParagraphProperties();
            ParagraphStyleId pStyleIdDefault = new ParagraphStyleId() { Val = "Footer" };

            pPropDefault.Append(pStyleIdDefault);

            Run runDefault = new Run();
            Text textDefault = new Text();
            textDefault.Text = "Footer in all pages [default]";

            runDefault.Append(textDefault);
            pDefault.Append(pPropDefault);
            pDefault.Append(runDefault);
            fDefault.Append(pDefault);
            defaultPart.Footer = fDefault;
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        public static void WordOpenXmlRemoveHeadersAndFooters()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;

                if (docMainPart.HeaderParts.Count() > 0 ||
                    docMainPart.FooterParts.Count() > 0)
                {
                    docMainPart.DeleteParts(docMainPart.HeaderParts);
                    docMainPart.DeleteParts(docMainPart.FooterParts);

                    Document myDocument = docMainPart.Document;

                    var myHeaders = myDocument.Descendants<HeaderReference>().ToList();
                    foreach (HeaderReference oneHeader in myHeaders)
                    {
                        oneHeader.Remove();
                    }

                    var myFooters = myDocument.Descendants<FooterReference>().ToList();
                    foreach (FooterReference oneFooter in myFooters)
                    {
                        oneFooter.Remove();
                    }

                    myDocument.Save();
                }
            }
            //gavdcodeend 13
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        public static void WordOpenXmlAddComment()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                // Locate the first paragraph in the document.
                Paragraph firstParagraph =
                    myWordDoc.MainDocumentPart.Document.Descendants<Paragraph>().First();
                Comments myComments = null;
                string id = "0";

                if (myWordDoc.MainDocumentPart.
                            GetPartsOfType<WordprocessingCommentsPart>().Count() > 0)
                {
                    myComments =
                        myWordDoc.MainDocumentPart.WordprocessingCommentsPart.Comments;
                    if (myComments.HasChildren)
                    {
                        id = myComments.Descendants<Comment>().Select
                                                                (e => e.Id.Value).Max();
                    }
                }
                else
                {
                    WordprocessingCommentsPart commentPart =
                     myWordDoc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    myComments = commentPart.Comments;
                }

                Paragraph myParagraph = new Paragraph(
                                        new Run(new Text("This is a new comment")));
                Comment newComment =
                    new Comment()
                    {
                        Id = id,
                        Author = "OtherAuthor",
                        Initials = "ioa",
                        Date = DateTime.Now
                    };
                newComment.AppendChild(myParagraph);
                myComments.AppendChild(newComment);
                myComments.Save();

                firstParagraph.InsertBefore(new CommentRangeStart()
                { Id = id }, firstParagraph.GetFirstChild<Run>());

                var commentEnd = firstParagraph.InsertAfter(new CommentRangeEnd()
                { Id = id }, firstParagraph.Elements<Run>().Last());

                firstParagraph.InsertAfter(new Run(
                                        new CommentReference() { Id = id }), commentEnd);
            }
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        public static void WordOpenXmlFindCommentsinDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", false))
            {
                WordprocessingCommentsPart commentsPart =
                    myWordDoc.MainDocumentPart.WordprocessingCommentsPart;

                if (commentsPart != null && commentsPart.Comments != null)
                {
                    IEnumerable<Comment> allComments =
                                            commentsPart.Comments.Elements<Comment>();
                    foreach (Comment oneComment in allComments)
                    {
                        Console.WriteLine(oneComment.InnerText);
                    }
                }
            }
        }
        //gavdcodeend 15

        //gavdcodebegin 16
        public static void WordOpenXmlDeleteComments()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                WordprocessingCommentsPart commentPart =
                    myWordDoc.MainDocumentPart.WordprocessingCommentsPart;

                if (commentPart == null)
                {
                    return;
                }

                List<Comment> commentsToDelete =
                                    commentPart.Comments.Elements<Comment>().ToList();
                if (!String.IsNullOrEmpty("OtherAuthor"))
                {
                    commentsToDelete = commentsToDelete.
                    Where(aut => aut.Author == "OtherAuthor").ToList();
                }
                IEnumerable<string> commentIds =
                    commentsToDelete.Select(del => del.Id.Value);

                foreach (Comment cmt in commentsToDelete)
                {
                    cmt.Remove();
                }

                commentPart.Comments.Save();

                Document myDoc = myWordDoc.MainDocumentPart.Document;

                List<CommentRangeStart> commentRangeStartToDelete =
                    myDoc.Descendants<CommentRangeStart>().
                    Where(cmt => commentIds.Contains(cmt.Id.Value)).ToList();
                foreach (CommentRangeStart oneComment in commentRangeStartToDelete)
                {
                    oneComment.Remove();
                }

                List<CommentRangeEnd> commentRangeEndToDelete =
                    myDoc.Descendants<CommentRangeEnd>().
                    Where(cmt => commentIds.Contains(cmt.Id.Value)).ToList();
                foreach (CommentRangeEnd oneComment in commentRangeEndToDelete)
                {
                    oneComment.Remove();
                }

                List<CommentReference> commentRangeReferenceToDelete =
                    myDoc.Descendants<CommentReference>().
                    Where(cmt => commentIds.Contains(cmt.Id.Value)).ToList();
                foreach (CommentReference oneComment in commentRangeReferenceToDelete)
                {
                    oneComment.Remove();
                }

                myDoc.Save();
            }
        }
        //gavdcodeend 16

        //gavdcodebegin 17
        public static void WordOpenXmlAddTableUsingData()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                var myDocument = myWordDoc.MainDocumentPart.Document;

                Table newTable = new Table();  // A new table

                TableProperties tableProps = new TableProperties(  // Table properties
                    new TableBorders(
                        new TopBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new BottomBorder
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 6
                        }));
                newTable.AppendChild<TableProperties>(tableProps);

                string[,] tableData = new string[,] { { "aa", "bb" }, { "cc", "dd" } };
                for (var cntrH = 0; cntrH <= tableData.GetUpperBound(0); cntrH++)
                {
                    var tblRow = new TableRow();  //New row
                    for (var cntrV = 0; cntrV <= tableData.GetUpperBound(1); cntrV++)
                    {
                        var tblCell = new TableCell();  // New cell
                        tblCell.Append(new Paragraph(new Run(
                                                    new Text(tableData[cntrH, cntrV]))));

                        tblCell.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tblRow.Append(tblCell);
                    }
                    newTable.Append(tblRow);
                }
                myDocument.Body.Append(newTable);
                myDocument.Save();
            }
        }
        //gavdcodeend 17

        //gavdcodebegin 18
        public static void WordOpenXmlAddTableUsingDirectData()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                var myDocument = myWordDoc.MainDocumentPart.Document;

                Table newTable = new Table();  // A new table

                TableProperties tableProp = new TableProperties(  // Table properties
                    new TableBorders(
                        new InsideHorizontalBorder()
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 12
                        },
                        new InsideVerticalBorder()
                        {
                            Val = new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 6
                        }));
                newTable.AppendChild<TableProperties>(tableProp);

                TableRow tblRow = new TableRow();  //New row

                TableCell tblCell01 = new TableCell();  // New cell
                tblCell01.Append(new TableCellProperties(
                    new TableCellWidth()
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = "2400"
                    }));
                tblCell01.Append(new Paragraph(new Run(new Text("My content"))));
                tblRow.Append(tblCell01); // Append cell to row

                TableCell tblCell02 = new TableCell(tblCell01.OuterXml);
                tblRow.Append(tblCell02);

                newTable.Append(tblRow);  // Append row to table

                myDocument.Body.Append(newTable);
                myDocument.Save();
            }
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        public static void WordOpenXmlModifyTextInCell()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                Table myTable =  // 1st table in the doc
                    myWordDoc.MainDocumentPart.Document.Body.Elements<Table>().First();

                TableRow tblRow = myTable.Elements<TableRow>().ElementAt(0); // 1st row
                TableCell tblCell = tblRow.Elements<TableCell>().ElementAt(0); // 1st cell

                Paragraph myParagraph = tblCell.Elements<Paragraph>().First(); // 1st par
                Run myRun = myParagraph.Elements<Run>().First();  // First run

                Text myText = myRun.Elements<Text>().First();
                myText.Text = "New text abc";
            }
        }
        //gavdcodeend 19

        //gavdcodebegin 20
        public static void WordOpenXmlDeleteRow()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                Table myTable =  // 1st table in the doc
                    myWordDoc.MainDocumentPart.Document.Body.Elements<Table>().First();

                TableRow tblRow = myTable.Elements<TableRow>().ElementAt(0); // 1st row
                tblRow.Remove();
            }
        }
        //gavdcodeend 20

        //gavdcodebegin 21
        public static void WordOpenXmlAddPictureInDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;

                ImagePart docImagePart = docMainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream imageStream = new
                            FileStream(@"C:\Temporary\MyImage.gif", FileMode.Open))
                {
                    docImagePart.FeedData(imageStream);
                }

                // Picture definition. See ISO standard 29500-1:2016
                //     https://www.iso.org/standard/71691.html
                Drawing imageElement =
                     new Drawing(
                         new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                             new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent()
                             { Cx = 990000L, Cy = 792000L },
                             new DocumentFormat.OpenXml.Drawing.Wordprocessing.
                                                                        EffectExtent()
                             {
                                 LeftEdge = 0L,
                                 TopEdge = 0L,
                                 RightEdge = 0L,
                                 BottomEdge = 0L
                             },
                             new DocumentFormat.OpenXml.Drawing.Wordprocessing.
                                                                        DocProperties()
                             {
                                 Id = (UInt32Value)1U,
                                 Name = "Picture 1"
                             },
                             new DocumentFormat.OpenXml.Drawing.Wordprocessing.
                                                NonVisualGraphicFrameDrawingProperties(
                                 new DocumentFormat.OpenXml.Drawing.
                                        GraphicFrameLocks()
                                 { NoChangeAspect = true }),
                             new DocumentFormat.OpenXml.Drawing.Graphic(
                                 new DocumentFormat.OpenXml.Drawing.GraphicData(
                                     new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                         new DocumentFormat.OpenXml.Drawing.Pictures.
                                                            NonVisualPictureProperties(
                                             new DocumentFormat.OpenXml.Drawing.Pictures.
                                                            NonVisualDrawingProperties()
                                             {
                                                 Id = (UInt32Value)0U,
                                                 Name = "New Bitmap Image.jpg"
                                             },
                                             new DocumentFormat.OpenXml.Drawing.Pictures.
                                                    NonVisualPictureDrawingProperties()),
                                         new DocumentFormat.OpenXml.Drawing.Pictures.
                                                                                BlipFill(
                                             new DocumentFormat.OpenXml.Drawing.Blip(
                                                 new DocumentFormat.OpenXml.Drawing.
                                                                    BlipExtensionList(
                                                     new DocumentFormat.OpenXml.Drawing.
                                                                        BlipExtension()
                                                     {
                                                         Uri =
                                                 "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                     })
                                             )
                                             {
                                                 Embed = docMainPart.GetIdOfPart(
                                                                            docImagePart),
                                                 CompressionState =
                                                 DocumentFormat.OpenXml.Drawing.
                                                            BlipCompressionValues.Print
                                             },
                                             new DocumentFormat.OpenXml.Drawing.Stretch(
                                                 new DocumentFormat.OpenXml.Drawing.
                                                                        FillRectangle())),
                                         new DocumentFormat.OpenXml.Drawing.Pictures.
                                                                        ShapeProperties(
                                             new DocumentFormat.OpenXml.Drawing.
                                                                        Transform2D(
                                                 new DocumentFormat.OpenXml.Drawing.
                                                            Offset()
                                                 { X = 0L, Y = 0L },
                                                 new DocumentFormat.OpenXml.Drawing.
                                                            Extents()
                                                 {
                                                     Cx = 990000L,
                                                     Cy = 792000L
                                                 }),
                                             new DocumentFormat.OpenXml.Drawing.
                                                            PresetGeometry(
                                                 new DocumentFormat.OpenXml.Drawing.
                                                            AdjustValueList()
                                             )
                                             {
                                                 Preset = DocumentFormat.OpenXml.Drawing.
                                                        ShapeTypeValues.Rectangle
                                             }))
                                 )
                                 {
                                     Uri =
                            "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                 })
                         )
                         {
                             DistanceFromTop = (UInt32Value)0U,
                             DistanceFromBottom = (UInt32Value)0U,
                             DistanceFromLeft = (UInt32Value)0U,
                             DistanceFromRight = (UInt32Value)0U,
                             EditId = "50D07946"
                         });

                myWordDoc.MainDocumentPart.Document.Body.AppendChild(new
                                                    Paragraph(new Run(imageElement)));
            }
        }
        //gavdcodeend 21

        //gavdcodebegin 22
        public static void WordOpenXmlDeleteAllPicturesFromDocument()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                MainDocumentPart docMainPart = myWordDoc.MainDocumentPart;

                IEnumerable<ImagePart> docImageParts = docMainPart.ImageParts;
                List<ImagePart> listImageParts = new List<ImagePart>();
                List<Drawing> listDrwParts = new List<Drawing>(docMainPart.RootElement.
                                                                Descendants<Drawing>());
                List<Drawing> listDrwdDeleteParts = new List<Drawing>();

                foreach (ImagePart oneImage in docImageParts)
                {
                    listImageParts.Add(oneImage);
                    IEnumerable<Drawing> allDrawings = listDrwParts.Where(dp =>
                        dp.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().
                        Any(pc => pc.BlipFill.Blip.Embed == docMainPart.
                        GetIdOfPart(oneImage)));

                    foreach (Drawing oneDrawing in allDrawings)
                    {
                        if (oneDrawing != null && !listDrwdDeleteParts.Contains(
                                                                            oneDrawing))
                            listDrwdDeleteParts.Add(oneDrawing);
                    }
                }

                foreach (Drawing oneDrwdDeletePart in listDrwdDeleteParts)
                {
                    oneDrwdDeletePart.Remove();
                }

                docMainPart.DeleteParts(listImageParts);
            }
        }
        //gavdcodeend 22

        //gavdcodebegin 23
        public static void WordOpenXmlAcceptAllChages()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                Body docBocy = myWordDoc.MainDocumentPart.Document.Body;

                // Accept formatting changes
                List<OpenXmlElement> myChanges =
                                docBocy.Descendants<ParagraphPropertiesChange>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList();

                foreach (OpenXmlElement oneChange in myChanges)
                {

                    oneChange.Remove();
                }

                // Accept the deletions
                List<OpenXmlElement> myDeletions =
                                docBocy.Descendants<Deleted>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList();

                myDeletions.AddRange(docBocy.Descendants<DeletedRun>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList());

                myDeletions.AddRange(docBocy.Descendants<DeletedMathControl>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList());

                foreach (OpenXmlElement oneDeletion in myDeletions)
                {
                    oneDeletion.Remove();
                }

                // Accept the insertions
                List<OpenXmlElement> myInsertions =
                                docBocy.Descendants<Inserted>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList();

                myInsertions.AddRange(docBocy.Descendants<InsertedRun>()
                                .Where(aut => aut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList());

                myInsertions.AddRange(docBocy.Descendants<InsertedMathControl>()
                                .Where(aaut => aaut.Author.Value == "OtherAuthor").
                                Cast<OpenXmlElement>().ToList());

                foreach (OpenXmlElement oneInsertion in myInsertions)
                {
                    foreach (Run oneRun in oneInsertion.Elements<Run>())
                    {
                        if (oneRun == oneInsertion.FirstChild)
                        {
                            oneInsertion.InsertAfterSelf(new Run(oneRun.OuterXml));
                        }
                        else
                        {
                            oneInsertion.NextSibling().InsertAfterSelf(new Run(oneRun.OuterXml));
                        }
                    }
                    oneInsertion.RemoveAttribute("rsidR",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    oneInsertion.RemoveAttribute("rsidRPr",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                    oneInsertion.Remove();
                }
            }
        }
        //gavdcodeend 23

        //gavdcodebegin 24
        public static void WordOpenXmlRemoveHiddenText()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                NameTable myNameTable = new NameTable();
                XmlNamespaceManager nameManager = new XmlNamespaceManager(myNameTable);
                nameManager.AddNamespace("w",
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                XmlDocument myXmlDoc = new XmlDocument(myNameTable);
                myXmlDoc.Load(myWordDoc.MainDocumentPart.GetStream());
                XmlNodeList allHiddenNodes = myXmlDoc.SelectNodes(
                                                            "//w:vanish", nameManager);
                foreach (XmlNode oneHiddenNode in allHiddenNodes)
                {
                    XmlNode topNode = oneHiddenNode.ParentNode.ParentNode;
                    XmlNode topParentNode = topNode.ParentNode;

                    topParentNode.RemoveChild(topNode);

                    if (topParentNode.HasChildNodes == false)
                    {
                        topParentNode.ParentNode.RemoveChild(topParentNode);
                    }
                }

                myXmlDoc.Save(myWordDoc.MainDocumentPart.GetStream(
                                                    FileMode.Create, FileAccess.Write));
            }
        }
        //gavdcodeend 24

        //gavdcodebegin 25
        public static void WordOpenXmlValidator()
        {
            using (WordprocessingDocument myWordDoc =
                WordprocessingDocument.Open(@"C:\Temporary\WordDoc01.docx", true))
            {
                OpenXmlValidator myValidator = new OpenXmlValidator();
                foreach (ValidationErrorInfo oneError in
                                                    myValidator.Validate(myWordDoc))
                {
                    Console.WriteLine("Error ");
                    Console.WriteLine("Description: " + oneError.Description);
                    Console.WriteLine("ErrorType: " + oneError.ErrorType);
                    Console.WriteLine("Node: " + oneError.Node);
                    Console.WriteLine("Path: " + oneError.Path.XPath);
                    Console.WriteLine("Part: " + oneError.Part.Uri);
                    Console.WriteLine(Environment.NewLine);
                }

                myWordDoc.Close();
            }
        }
        //gavdcodeend 25
    }
}