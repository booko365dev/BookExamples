using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace BHAZ
{
    class Program
    {
        static void Main(string[] args)
        {
            //PowerPointOpenXmlCreatePresentation();
            //PowerPointOpenXmlFindTextInSlide();
            //PowerPointOpenXmlFindAllTextInOneSlide();
            //PowerPointOpenXmlCopyTheme();
            //PowerPointOpenXmlInsertNewSlide();
            //PowerPointOpenXmlFindAllSlideTitles();
            //PowerPointOpenXmlFindNumberOfSlides();
            //PowerPointOpenXmlMoveSlide();
            //PowerPointOpenXmlDeleteOneSlide();
            //PowerPointOpenXmlAddCommentToSlide();
            //PowerPointOpenXmlRemoveAllCommentsAuthor();
            //PowerPointOpenXmlAddNotesToSlide();
            //PowerPointOpenXmlFindTextInNotes();
            //PowerPointOpenXmlAddImageToSlide();
            PowerPointOpenXmlAddShapeToSlide();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        public static void PowerPointOpenXmlCreatePresentation()
        {
            using (PresentationDocument myPowerPointDoc =
                    PresentationDocument.Create(@"C:\Temporary\PowerPointDoc01.pptx",
                    PresentationDocumentType.Presentation))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.AddPresentationPart();
                myPresentationPart.Presentation = new Presentation();

                SlideMasterIdList mySlideMasterIdList = new SlideMasterIdList(
                                new SlideMasterId()
                                {
                                    Id = (UInt32Value)2147483648U,
                                    RelationshipId = "rId1"
                                });
                SlideIdList mySlideIdList = new SlideIdList(
                                new SlideId()
                                {
                                    Id = (UInt32Value)256U,
                                    RelationshipId = "rId2"
                                });
                SlideSize mySlideSize = new SlideSize()
                {
                    Cx = 9144000,
                    Cy = 6858000,
                    Type = SlideSizeValues.Screen4x3
                };
                NotesSize myNotesSize = new NotesSize()
                {
                    Cx = 6858000,
                    Cy = 9144000
                };
                DefaultTextStyle myDefaultTextStyle = new DefaultTextStyle();

                myPresentationPart.Presentation.Append(mySlideMasterIdList,
                                                mySlideIdList, mySlideSize,
                                                myNotesSize, myDefaultTextStyle);

                PPCreateDefaultSlide(myPresentationPart);
            }
        }

        private static void PPCreateDefaultSlide(PresentationPart PPPresentationPart)
        {
            SlidePart mySlidePart;
            SlideLayoutPart mySlideLayoutPart;
            SlideMasterPart mySlideMasterPart;
            ThemePart myThemePart;

            mySlidePart = PPCreateSlidePart(PPPresentationPart);
            mySlideLayoutPart = PPCreateSlideLayoutPart(mySlidePart);
            mySlideMasterPart = PPCreateSlideMasterPart(mySlideLayoutPart);
            myThemePart = PPCreateTheme(mySlideMasterPart);

            mySlideMasterPart.AddPart(mySlideLayoutPart, "rId1");
            PPPresentationPart.AddPart(mySlideMasterPart, "rId1");
            PPPresentationPart.AddPart(myThemePart, "rId5");
        }

        private static SlidePart PPCreateSlidePart(PresentationPart PPPresentationPart)
        {
            SlidePart mySlidePart = PPPresentationPart.AddNewPart<SlidePart>("rId2");
            mySlidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)1U,
                                    Name = ""
                                },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)2U,
                                        Name = "My Title"
                                    },
                                    new P.NonVisualShapeDrawingProperties(
                                            new ShapeLocks()
                                            {
                                                NoGrouping = true
                                            }),
                                    new ApplicationNonVisualDrawingProperties(
                                            new PlaceholderShape())),
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new EndParagraphRunProperties()
                                    {
                                        Language = "en-US"
                                    }))))),
                    new ColorMapOverride(new MasterColorMapping()));

            return mySlidePart;
        }

        private static SlideLayoutPart PPCreateSlideLayoutPart(SlidePart PPSlidePart)
        {
            SlideLayoutPart mySlideLayoutPart =
                                        PPSlidePart.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout mySlideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = ""
                },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new TransformGroup()),
                new P.Shape(
                new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)2U,
                    Name = ""
                },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks()
                {
                    NoGrouping = true
                }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                new P.ShapeProperties(),
                new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            mySlideLayoutPart.SlideLayout = mySlideLayout;

            return mySlideLayoutPart;
        }

        private static SlideMasterPart PPCreateSlideMasterPart(
                                                    SlideLayoutPart PPSlideLayoutPart)
        {
            SlideMasterPart mySlideMasterPart =
                                PPSlideLayoutPart.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster mySlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = ""
                },
                new P.NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new TransformGroup()),
                new P.Shape(
                new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)2U,
                    Name = "My Placeholder Title"
                },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks()
                {
                    NoGrouping = true
                }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape()
                {
                    Type = PlaceholderValues.Title
                })),
                new P.ShapeProperties(),
                new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap()
            {
                Background1 = D.ColorSchemeIndexValues.Light1,
                Text1 = D.ColorSchemeIndexValues.Dark1,
                Background2 = D.ColorSchemeIndexValues.Light2,
                Text2 = D.ColorSchemeIndexValues.Dark2,
                Accent1 = D.ColorSchemeIndexValues.Accent1,
                Accent2 = D.ColorSchemeIndexValues.Accent2,
                Accent3 = D.ColorSchemeIndexValues.Accent3,
                Accent4 = D.ColorSchemeIndexValues.Accent4,
                Accent5 = D.ColorSchemeIndexValues.Accent5,
                Accent6 = D.ColorSchemeIndexValues.Accent6,
                Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
            },
            new SlideLayoutIdList(new SlideLayoutId()
            {
                Id = (UInt32Value)2147483649U,
                RelationshipId = "rId1"
            }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            mySlideMasterPart.SlideMaster = mySlideMaster;

            return mySlideMasterPart;
        }

        private static ThemePart PPCreateTheme(SlideMasterPart PPSlideMasterPart)
        {
            ThemePart myThemePart = PPSlideMasterPart.AddNewPart<ThemePart>("rId5");
            D.Theme myTheme = new D.Theme()
            {
                Name = "My Theme"
            };

            D.ThemeElements myThemeElements = new D.ThemeElements(
            new D.ColorScheme(
                new D.Dark1Color(new D.SystemColor()
                {
                    Val = D.SystemColorValues.WindowText,
                    LastColor = "000000"
                }),
                new D.Light1Color(new D.SystemColor()
                {
                    Val = D.SystemColorValues.Window,
                    LastColor = "FFFFFF"
                }),
                new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
                new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
                new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
                new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
                new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
                new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
                new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
                new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
                new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
                new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
                new D.FontScheme(
                new D.MajorFont(
                new D.LatinFont() { Typeface = "Calibri" },
                new D.EastAsianFont() { Typeface = "" },
                new D.ComplexScriptFont() { Typeface = "" }),
                new D.MinorFont(
                new D.LatinFont() { Typeface = "Calibri" },
                new D.EastAsianFont() { Typeface = "" },
                new D.ComplexScriptFont() { Typeface = "" }))
                { Name = "Office" },
                new D.FormatScheme(
                new D.FillStyleList(
                new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                    new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                    new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                new D.NoFill(),
                new D.PatternFill(),
                new D.GroupFill()),
                new D.LineStyleList(
                new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                    new D.Shade() { Val = 95000 },
                    new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                {
                    Width = 9525,
                    CapType = D.LineCapValues.Flat,
                    CompoundLineType = D.CompoundLineValues.Single,
                    Alignment = D.PenAlignmentValues.Center
                },
                new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                    new D.Shade() { Val = 95000 },
                    new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                {
                    Width = 9525,
                    CapType = D.LineCapValues.Flat,
                    CompoundLineType = D.CompoundLineValues.Single,
                    Alignment = D.PenAlignmentValues.Center
                },
                new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                    new D.Shade() { Val = 95000 },
                    new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                {
                    Width = 9525,
                    CapType = D.LineCapValues.Flat,
                    CompoundLineType = D.CompoundLineValues.Single,
                    Alignment = D.PenAlignmentValues.Center
                }),
                new D.EffectStyleList(
                new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                    new D.RgbColorModelHex(
                    new D.Alpha() { Val = 38000 })
                    { Val = "000000" })
                {
                    BlurRadius = 40000L,
                    Distance = 20000L,
                    Direction = 5400000,
                    RotateWithShape = false
                })),
                new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                    new D.RgbColorModelHex(
                    new D.Alpha() { Val = 38000 })
                    { Val = "000000" })
                {
                    BlurRadius = 40000L,
                    Distance = 20000L,
                    Direction = 5400000,
                    RotateWithShape = false
                })),
                new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                    new D.RgbColorModelHex(
                    new D.Alpha() { Val = 38000 })
                    { Val = "000000" })
                {
                    BlurRadius = 40000L,
                    Distance = 20000L,
                    Direction = 5400000,
                    RotateWithShape = false
                }))),
                new D.BackgroundFillStyleList(
                new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                    new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                    { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                    new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                    { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                    new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                    { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                    new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                    { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                    new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                    { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
                { Name = "Office" });

            myTheme.Append(myThemeElements);
            myTheme.Append(new D.ObjectDefaults());
            myTheme.Append(new D.ExtraColorSchemeList());

            myThemePart.Theme = myTheme;

            return myThemePart;
        }

        public static void PowerPointOpenXmlFindTextInSlide()
        {
            using (PresentationDocument myPowerPointDoc = 
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", false))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                OpenXmlElementList mySlideIdList = myPresentationPart.Presentation.
                                                            SlideIdList.ChildElements;
                string myRelationshipId = (mySlideIdList[0] as SlideId).RelationshipId;
                myRelationshipId = (mySlideIdList[0] as SlideId).RelationshipId;

                SlidePart mySlidePart = (SlidePart)myPresentationPart.GetPartById(
                                                                    myRelationshipId);

                IEnumerable<D.Text> allText = mySlidePart.Slide.Descendants<D.Text>();
                foreach (D.Text oneText in allText)
                {
                    Console.WriteLine(oneText.Text);
                }
            }
        }

        public static void PowerPointOpenXmlFindAllTextInOneSlide()
        {
            int toFindTextInSlidePositionIndex = 1;

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", false))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                if (myPresentationPart != null && 
                    myPresentationPart.Presentation != null)
                {
                    Presentation myPresentation = myPresentationPart.Presentation;

                    if (myPresentation.SlideIdList != null)
                    {
                        DocumentFormat.OpenXml.OpenXmlElementList allSlideIds =
                                                myPresentation.SlideIdList.ChildElements;

                        if (toFindTextInSlidePositionIndex < allSlideIds.Count)
                        {
                            string mySlidePartRelationshipId = 
                                (allSlideIds[toFindTextInSlidePositionIndex] as SlideId).
                                                                        RelationshipId;

                            SlidePart mySlidePart =
                                (SlidePart)myPresentationPart.GetPartById(
                                                            mySlidePartRelationshipId);

                            if (mySlidePart.Slide != null)
                            {
                                foreach (D.Paragraph oneParagraph in
                                            mySlidePart.Slide.Descendants<D.Paragraph>())
                                {
                                    foreach (D.Text oneText in
                                                    oneParagraph.Descendants<D.Text>())
                                    {
                                        Console.WriteLine(oneText.Text);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public static void PowerPointOpenXmlCopyTheme()
        {
            using (PresentationDocument sourceThemePowerPointDoc = 
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc02.pptx", false))
            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                SlideMasterPart mySlideMasterPart = myPresentationPart.SlideMasterParts.
                                                                        ElementAt(0);
                string myRelationshipId = myPresentationPart.
                                                        GetIdOfPart(mySlideMasterPart);
                SlideMasterPart newSlideMasterPart = sourceThemePowerPointDoc.
                                        PresentationPart.SlideMasterParts.ElementAt(0);

                myPresentationPart.DeletePart(myPresentationPart.ThemePart);
                myPresentationPart.DeletePart(mySlideMasterPart);

                newSlideMasterPart = myPresentationPart.AddPart(newSlideMasterPart, 
                                                                    myRelationshipId);
                myPresentationPart.AddPart(newSlideMasterPart.ThemePart);

                Dictionary<string, SlideLayoutPart> newSlideLayouts = 
                                            new Dictionary<string, SlideLayoutPart>();
                string layoutType = null;

                foreach (SlideLayoutPart oneSlideLayoutPart in 
                                                    newSlideMasterPart.SlideLayoutParts)
                {
                    layoutType = oneSlideLayoutPart.SlideLayout.CommonSlideData.Name;
                    if(string.IsNullOrEmpty(layoutType) == false)
                        newSlideLayouts.Add(layoutType, oneSlideLayoutPart);
                }

                SlideLayoutPart newLayoutPart = null;
                string defaultLayoutType = "Title and Content";

                foreach (SlidePart oneSlidePart in myPresentationPart.SlideParts)
                {
                    layoutType = null;

                    if (oneSlidePart.SlideLayoutPart != null)
                    {
                        layoutType = oneSlidePart.SlideLayoutPart.SlideLayout.
                                                                    CommonSlideData.Name;
                        oneSlidePart.DeletePart(oneSlidePart.SlideLayoutPart);
                    }

                    if (layoutType != null && 
                            newSlideLayouts.TryGetValue(layoutType, out newLayoutPart))
                    {
                        oneSlidePart.AddPart(newLayoutPart);
                    }
                    else
                    {
                        newLayoutPart = newSlideLayouts[defaultLayoutType];
                        oneSlidePart.AddPart(newLayoutPart);
                    }
                }
            }
        }

        public static void PowerPointOpenXmlInsertNewSlide()
        {
            int newSlidePositionIndex = 1;

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart presentationPart = myPowerPointDoc.PresentationPart;

                Slide newSlide = new Slide(new CommonSlideData(new ShapeTree()));
                uint drawingObjectId = 1;

                P.NonVisualGroupShapeProperties myNonVisualProperties = 
                            newSlide.CommonSlideData.ShapeTree.AppendChild(
                                new P.NonVisualGroupShapeProperties());
                myNonVisualProperties.NonVisualDrawingProperties = 
                                new P.NonVisualDrawingProperties() {
                                    Id = 1, Name = "" };
                myNonVisualProperties.NonVisualGroupShapeDrawingProperties = 
                                new P.NonVisualGroupShapeDrawingProperties();
                myNonVisualProperties.ApplicationNonVisualDrawingProperties = 
                                new ApplicationNonVisualDrawingProperties();

                newSlide.CommonSlideData.ShapeTree.AppendChild(
                                new GroupShapeProperties());
                P.Shape myTitleShape = newSlide.CommonSlideData.ShapeTree.AppendChild(
                                new P.Shape());
                drawingObjectId++;

                myTitleShape.NonVisualShapeProperties = new P.NonVisualShapeProperties
                            (new P.NonVisualDrawingProperties() {
                                Id = drawingObjectId,
                                Name = "My Title" },
                    new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() {
                                NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {
                                Type = PlaceholderValues.Title }));
                myTitleShape.ShapeProperties = new P.ShapeProperties();

                myTitleShape.TextBody = new P.TextBody(new D.BodyProperties(),
                        new D.ListStyle(),
                        new D.Paragraph(new D.Run(new D.Text() {
                                Text = "My New Slide" })));

                P.Shape myBodyShape = newSlide.CommonSlideData.ShapeTree.AppendChild(
                        new P.Shape());
                drawingObjectId++;

                myBodyShape.NonVisualShapeProperties = 
                        new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() {
                            Id = drawingObjectId,
                            Name = "My Content Placeholder" },
                        new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() {
                            NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {
                            Index = 1 }));
                myBodyShape.ShapeProperties = new P.ShapeProperties();

                myBodyShape.TextBody = new P.TextBody(new D.BodyProperties(),
                        new D.ListStyle(),
                        new D.Paragraph());

                SlidePart mySlidePart = presentationPart.AddNewPart<SlidePart>();

                newSlide.Save(mySlidePart);

                SlideIdList mySlideIdList = presentationPart.Presentation.SlideIdList;
                uint maxSlideId = 1;
                SlideId prevSlideId = null;

                foreach (SlideId oneSlideId in mySlideIdList.ChildElements)
                {
                    if (oneSlideId.Id > maxSlideId)
                    {
                        maxSlideId = oneSlideId.Id;
                    }

                    newSlidePositionIndex--;
                    if (newSlidePositionIndex == 0)
                    {
                        prevSlideId = oneSlideId;
                    }
                }
                maxSlideId++;

                SlidePart lastSlidePart;
                if (prevSlideId != null)
                {
                    lastSlidePart = (SlidePart)presentationPart.GetPartById(
                                                        prevSlideId.RelationshipId);
                }
                else
                {
                    lastSlidePart = (SlidePart)presentationPart.GetPartById((
                                (SlideId)(mySlideIdList.ChildElements[0])).RelationshipId);
                }

                if (null != lastSlidePart.SlideLayoutPart)
                {
                    mySlidePart.AddPart(lastSlidePart.SlideLayoutPart);
                }

                SlideId newSlideId = mySlideIdList.InsertAfter(new SlideId(), prevSlideId);
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = presentationPart.GetIdOfPart(mySlidePart);

                presentationPart.Presentation.Save();
            }
        }

        public static void PowerPointOpenXmlFindAllSlideTitles()
        {
            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                if (myPresentationPart != null &&
                    myPresentationPart.Presentation != null)
                {
                    Presentation myPresentation = myPresentationPart.Presentation;

                    if (myPresentation.SlideIdList != null)
                    {
                        List<string> titlesList = new List<string>();

                        foreach (SlideId oneSlideId in myPresentation.SlideIdList.
                                                                    Elements<SlideId>())
                        {
                            SlidePart mySlidePart = myPresentationPart.GetPartById(
                                                oneSlideId.RelationshipId) as SlidePart;

                            if (mySlidePart.Slide != null)
                            {
                                IEnumerable<P.Shape> allShapes = from shape 
                                             in mySlidePart.Slide.Descendants<P.Shape>()
                                             where IsTitleShape(shape)
                                             select shape;

                                foreach (P.Shape oneShape in allShapes)
                                {
                                    foreach (D.Paragraph oneParagraph in oneShape.
                                                    TextBody.Descendants<D.Paragraph>())
                                    {
                                        foreach (D.Text oneText in oneParagraph.
                                                                Descendants<D.Text>())
                                        {
                                            Console.WriteLine(oneText.Text);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private static bool IsTitleShape(P.Shape ShapeToSearch)
        {
            PlaceholderShape myPlaceholderShape = ShapeToSearch.NonVisualShapeProperties.
                ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();
            if (myPlaceholderShape != null && 
                myPlaceholderShape.Type != null && 
                myPlaceholderShape.Type.HasValue)
            {
                switch ((PlaceholderValues)myPlaceholderShape.Type)
                {
                    case PlaceholderValues.Title:

                    case PlaceholderValues.CenteredTitle:
                        return true;

                    default:
                        return false;
                }
            }

            return false;
        }

        public static void PowerPointOpenXmlFindNumberOfSlides(bool includeHidden = true)
        {
            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", false))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                if (myPresentationPart != null)
                {
                    Console.WriteLine("Including hidden slides: " +
                                    myPresentationPart.SlideParts.Count().ToString());

                    IEnumerable<SlidePart> allSlides = 
                                myPresentationPart.SlideParts.Where(
                                              (sl) => (sl.Slide != null) &&
                                              ((sl.Slide.Show == null) || 
                                              (sl.Slide.Show.HasValue &&
                                              sl.Slide.Show.Value)));
                    Console.WriteLine("Only visible slides: " + 
                                                        allSlides.Count().ToString());
                }
            }
        }

        public static void PowerPointOpenXmlMoveSlide()
        {
            int slideIndexFrom = 1; 
            int slideIndexTo = 2;

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                int numberOflides = myPresentationPart.SlideParts.Count();

                Presentation myPresentation = myPresentationPart.Presentation;
                SlideIdList mySlideIdList = myPresentation.SlideIdList;

                SlideId sourceSlide = mySlideIdList.ChildElements[slideIndexFrom] as 
                                                                                SlideId;
                SlideId targetSlide = null;

                if (slideIndexTo == 0)
                {
                    targetSlide = null;
                }
                else if (slideIndexFrom < slideIndexTo)
                {
                    targetSlide = mySlideIdList.ChildElements[slideIndexTo] as SlideId;
                }
                else
                {
                    targetSlide = mySlideIdList.ChildElements[slideIndexTo - 1] as 
                                                                                SlideId;
                }

                sourceSlide.Remove();
                mySlideIdList.InsertAfter(sourceSlide, targetSlide);

                myPresentation.Save();
            }
        }

        public static void PowerPointOpenXmlDeleteOneSlide()
        {
            int slideIndex = 1;

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                Presentation myPresentation = myPresentationPart.Presentation;
                SlideIdList mySlideIdList = myPresentation.SlideIdList;
                SlideId mySlideId = mySlideIdList.ChildElements[slideIndex] as SlideId;
                string mySlideRelationshipId = mySlideId.RelationshipId;

                mySlideIdList.RemoveChild(mySlideId);

                if (myPresentation.CustomShowList != null)
                {
                    foreach (CustomShow oneCustomShow in myPresentation.CustomShowList.
                                                            Elements<CustomShow>())
                    {
                        if (oneCustomShow.SlideList != null)
                        {
                            LinkedList<SlideListEntry> allSlideListEntries = 
                                                new LinkedList<SlideListEntry>();
                            foreach (SlideListEntry oneSlideListEntry in 
                                                oneCustomShow.SlideList.Elements())
                            {
                                if (oneSlideListEntry.Id != null && 
                                    oneSlideListEntry.Id == mySlideRelationshipId)
                                {
                                    allSlideListEntries.AddLast(oneSlideListEntry);
                                }
                            }

                            foreach (SlideListEntry oneSlideListEntry in 
                                                                    allSlideListEntries)
                            {
                                oneCustomShow.SlideList.RemoveChild(oneSlideListEntry);
                            }
                        }
                    }
                }

                myPresentation.Save();

                SlidePart mySlidePart = myPresentationPart.GetPartById(
                                                mySlideRelationshipId) as SlidePart;
                myPresentationPart.DeletePart(mySlidePart);
            }
        }

        public static void PowerPointOpenXmlAddCommentToSlide()
        {
            string myAuthorInitials = "GAVD";
            string myAuthorName = "Guitaca";
            string myCommentText = "Comment from Guitaca";

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                SlideId mySlideId = myPresentationPart.Presentation.SlideIdList.
                                                            GetFirstChild<SlideId>();
                string myRelationshipId = mySlideId.RelationshipId;
                SlidePart mySlidePart = (SlidePart)myPresentationPart.
                                                            GetPartById(myRelationshipId);

                CommentAuthorsPart myCommentAuthorsPart;
                if (myPowerPointDoc.PresentationPart.CommentAuthorsPart == null)
                {
                    myCommentAuthorsPart = myPowerPointDoc.PresentationPart.
                                                    AddNewPart<CommentAuthorsPart>();
                }
                else
                {
                    myCommentAuthorsPart = myPowerPointDoc.PresentationPart.
                                                    CommentAuthorsPart;
                }

                if (myCommentAuthorsPart.CommentAuthorList == null)
                {
                    myCommentAuthorsPart.CommentAuthorList = new CommentAuthorList();
                }

                uint myAuthorId = 0;
                CommentAuthor myCommentAuthor = null;

                if (myCommentAuthorsPart.CommentAuthorList.HasChildren)
                {
                    var allCommentAuthors = myCommentAuthorsPart.CommentAuthorList.
                                                Elements<CommentAuthor>().
                                                Where(au => au.Name == myAuthorName && 
                                                      au.Initials == myAuthorInitials);
                    if (allCommentAuthors.Any())
                    {
                        myCommentAuthor = allCommentAuthors.First();
                        myAuthorId = myCommentAuthor.Id;
                    }

                    if (myCommentAuthor == null)
                    {
                        myAuthorId = myCommentAuthorsPart.CommentAuthorList.
                                                Elements<CommentAuthor>().
                                                Select(au => au.Id.Value).Max();
                    }
                }

                if (myCommentAuthor == null)
                {
                    myAuthorId++;
                    myCommentAuthor = myCommentAuthorsPart.CommentAuthorList.
                            AppendChild<CommentAuthor> (
                                new CommentAuthor()
                                {
                                    Id = myAuthorId,
                                    Name = myAuthorName,
                                    Initials = myAuthorInitials,
                                    ColorIndex = 0
                                });
                }

                SlideCommentsPart mySlideCommentsPart;
                if (mySlidePart.GetPartsOfType<SlideCommentsPart>().Count() == 0)
                {
                    mySlideCommentsPart = mySlidePart.AddNewPart<SlideCommentsPart>();
                }
                else
                {
                    mySlideCommentsPart = mySlidePart.
                                        GetPartsOfType<SlideCommentsPart>().First();
                }

                if (mySlideCommentsPart.CommentList == null)
                {
                    mySlideCommentsPart.CommentList = new CommentList();
                }

                uint myCommentIndex = myCommentAuthor.LastIndex == 
                                                null ? 1 : myCommentAuthor.LastIndex + 1;
                myCommentAuthor.LastIndex = myCommentIndex;

                Comment myComment = mySlideCommentsPart.CommentList.AppendChild<Comment>(
                    new Comment()
                    {
                        AuthorId = myAuthorId,
                        Index = myCommentIndex,
                        DateTime = DateTime.Now
                    });

                myComment.Append(
                    new P.Position() {
                        X = 100,
                        Y = 200 },
                    new P.Text() {
                        Text = myCommentText });

                myCommentAuthorsPart.CommentAuthorList.Save();
                mySlideCommentsPart.CommentList.Save();
            }
        }

        public static void PowerPointOpenXmlRemoveAllCommentsAuthor()
        {
            string myAuthorName = "Guitaca";

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                IEnumerable<CommentAuthor> myCommentAuthors =
                    myPresentationPart.CommentAuthorsPart.CommentAuthorList.
                            Elements<CommentAuthor>().
                            Where(au => au.Name.Value.Equals(myAuthorName));

                foreach (CommentAuthor oneCommentAuthor in myCommentAuthors)
                {
                    UInt32Value oneAuthorId = oneCommentAuthor.Id;

                    foreach (SlidePart oneSlidePart in 
                                        myPowerPointDoc.PresentationPart.SlideParts)
                    {
                        SlideCommentsPart mySlideCommentsPart = 
                                                        oneSlidePart.SlideCommentsPart;
                        if (mySlideCommentsPart != null && 
                            oneSlidePart.SlideCommentsPart.CommentList != null)
                        {
                            IEnumerable<Comment> allCommentList =
                                mySlideCommentsPart.CommentList.
                                        Elements<Comment>().
                                        Where(cm => cm.AuthorId == oneAuthorId.Value);
                            List<Comment> allComments = new List<Comment>();
                            allComments = allCommentList.ToList<Comment>();

                            foreach (Comment oneComment in allComments)
                            {
                                mySlideCommentsPart.CommentList.
                                                    RemoveChild<Comment>(oneComment);
                            }

                            if (mySlideCommentsPart.CommentList.ChildElements.Count == 0)
                                oneSlidePart.DeletePart(mySlideCommentsPart);
                        }
                    }

                    myPresentationPart.CommentAuthorsPart.CommentAuthorList.
                                        RemoveChild<CommentAuthor>(oneCommentAuthor);
                }
            }
        }

        public static void PowerPointOpenXmlAddNotesToSlide()
        {
            string myNoteText = "Note from Guitaca";

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                SlideId mySlideId = myPresentationPart.Presentation.SlideIdList.
                                                           GetFirstChild<SlideId>();
                string myRelationshipId = mySlideId.RelationshipId;
                SlidePart mySlidePart = (SlidePart)myPresentationPart.
                                                           GetPartById(myRelationshipId);

                NotesSlidePart myNotesSlidePart = mySlidePart.NotesSlidePart;
                if (myNotesSlidePart == null)
                {
                    myNotesSlidePart = mySlidePart.
                                           AddNewPart<NotesSlidePart>(myRelationshipId);
                }
                NotesSlide myNotesSlide = new NotesSlide(
                       new CommonSlideData(new ShapeTree(
                         new P.NonVisualGroupShapeProperties(
                           new P.NonVisualDrawingProperties() {
                                        Id = (UInt32Value)1U,
                                        Name = "" },
                           new P.NonVisualGroupShapeDrawingProperties(),
                           new ApplicationNonVisualDrawingProperties()),
                           new GroupShapeProperties(new D.TransformGroup()),
                           new P.Shape(
                               new P.NonVisualShapeProperties(
                                   new P.NonVisualDrawingProperties() {
                                        Id = (UInt32Value)2U,
                                        Name = "SlideImage01" },
                                   new P.NonVisualShapeDrawingProperties(
                                       new D.ShapeLocks() {
                                           NoGrouping = true,
                                           NoRotation = true,
                                           NoChangeAspect = true }),
                                   new ApplicationNonVisualDrawingProperties(
                                       new PlaceholderShape() {
                                           Type = PlaceholderValues.SlideImage })),
                               new P.ShapeProperties()),
                           new P.Shape(
                               new P.NonVisualShapeProperties(
                                   new P.NonVisualDrawingProperties() {
                                        Id = (UInt32Value)3U,
                                       Name = "NotesPlaceholder01" },
                                   new P.NonVisualShapeDrawingProperties(
                                       new D.ShapeLocks() {
                                           NoGrouping = true }),
                                   new ApplicationNonVisualDrawingProperties(
                                       new PlaceholderShape() {
                                           Type = PlaceholderValues.Body,
                                           Index = (UInt32Value)1U })),
                               new P.ShapeProperties(),
                               new P.TextBody(
                                   new D.BodyProperties(),
                                   new D.ListStyle(),
                                   new D.Paragraph(
                                       new D.Run(
                                           new D.RunProperties() {
                                               Language = "en-US",
                                               Dirty = false },
                                           new D.Text() {
                                               Text = myNoteText }),
                                       new D.EndParagraphRunProperties() {
                                           Language = "en-US",
                                           Dirty = false }))
                                   ))),
                                new ColorMapOverride(
                                    new D.MasterColorMapping()));

                myNotesSlidePart.NotesSlide = myNotesSlide;
            }
        }

        public static void PowerPointOpenXmlFindTextInNotes()
        {
            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", false))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;
                OpenXmlElementList mySlideIdList = myPresentationPart.Presentation.
                                                            SlideIdList.ChildElements;
                string myRelationshipId = (mySlideIdList[0] as SlideId).RelationshipId;
                myRelationshipId = (mySlideIdList[0] as SlideId).RelationshipId;

                SlidePart mySlidePart = (SlidePart)myPresentationPart.GetPartById(
                                                                    myRelationshipId);

                IEnumerable<D.Text> allText = mySlidePart.NotesSlidePart.NotesSlide.
                                                                Descendants<D.Text>();
                foreach (D.Text oneText in allText)
                {
                    Console.WriteLine(oneText.Text);
                }
            }
        }

        public static void PowerPointOpenXmlAddImageToSlide()
        {
            string myImage = @"C:\Temporary\MyPicture.jpg";

            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                Presentation myPresentation = myPresentationPart.Presentation;
                SlidePart mySlidePart = myPresentation.PresentationPart.
                                                                    SlideParts.First();
                ImagePart myImagePart = mySlidePart.AddImagePart(ImagePartType.Png);

                using (FileStream imageStream = File.OpenRead(myImage))
                {
                    myImagePart.FeedData(imageStream);
                }

                ShapeTree myShapeTree = mySlidePart.Slide.Descendants<P.ShapeTree>().
                                                                                First();

                P.Picture myPicture = new P.Picture();

                myPicture.NonVisualPictureProperties = new P.NonVisualPictureProperties();
                myPicture.NonVisualPictureProperties.Append(
                            new P.NonVisualDrawingProperties {
                                Name = "My Picture",
                                Id = (UInt32)myShapeTree.ChildElements.Count - 1 });

                P.NonVisualPictureDrawingProperties myNonVisualPictureDrawingProp = 
                                            new P.NonVisualPictureDrawingProperties();
                myNonVisualPictureDrawingProp.Append(new D.PictureLocks() {
                            NoChangeAspect = true });
                myPicture.NonVisualPictureProperties.Append(
                                        myNonVisualPictureDrawingProp);
                myPicture.NonVisualPictureProperties.Append(
                                        new P.ApplicationNonVisualDrawingProperties());

                P.BlipFill myBlipFill = new P.BlipFill();
                D.Blip myBlip = new D.Blip() {
                            Embed = mySlidePart.GetIdOfPart(myImagePart) };
                D.BlipExtensionList myBlipExtensionList = new D.BlipExtensionList();
                D.BlipExtension myBlipExtension = new D.BlipExtension() {
                    Uri = "{12345678-ABCD-9876-EFAB-123456789ABC}" };
                var myUseLocalDpi = new DocumentFormat.OpenXml.Office2010.
                                            Drawing.UseLocalDpi() {
                                                        Val = false };
                myUseLocalDpi.AddNamespaceDeclaration("a14", 
                            "http://schemas.microsoft.com/office/drawing/2010/main");
                myBlipExtension.Append(myUseLocalDpi);
                myBlipExtensionList.Append(myBlipExtension);
                myBlip.Append(myBlipExtensionList);
                D.Stretch myStretch = new D.Stretch();
                myStretch.Append(new D.FillRectangle());
                myBlipFill.Append(myBlip);
                myBlipFill.Append(myStretch);
                myPicture.Append(myBlipFill);

                myPicture.ShapeProperties = new P.ShapeProperties();
                myPicture.ShapeProperties.Transform2D = new D.Transform2D();
                myPicture.ShapeProperties.Transform2D.Append(new D.Offset {
                                X = 100,
                                Y = 50 });
                myPicture.ShapeProperties.Transform2D.Append(new D.Extents {
                                Cx = 1000000,
                                Cy = 1000000 });
                myPicture.ShapeProperties.Append(new D.PresetGeometry {
                    Preset = D.ShapeTypeValues.Rectangle });

                myShapeTree.Append(myPicture);
            }
        }

        public static void PowerPointOpenXmlAddShapeToSlide()
        {
            using (PresentationDocument myPowerPointDoc =
                 PresentationDocument.Open(@"C:\Temporary\PowerPointDoc01.pptx", true))
            {
                PresentationPart myPresentationPart = myPowerPointDoc.PresentationPart;

                Presentation myPresentation = myPresentationPart.Presentation;
                ShapeTree myShapeTree = myPresentation.PresentationPart.SlideParts.
                            ElementAt(0).Slide.Descendants<P.ShapeTree>().First();

                P.Shape myShape = new P.Shape();
                myShape.NonVisualShapeProperties = new P.NonVisualShapeProperties();
                myShape.NonVisualShapeProperties.Append(new P.NonVisualDrawingProperties {
                            Name = "My Shape",
                            Id = (UInt32)myShapeTree.ChildElements.Count - 1 });
                myShape.NonVisualShapeProperties.Append(
                            new P.NonVisualShapeDrawingProperties());
                myShape.NonVisualShapeProperties.Append(
                            new P.ApplicationNonVisualDrawingProperties());

                myShape.ShapeProperties = new P.ShapeProperties();
                myShape.ShapeProperties.Transform2D = new D.Transform2D();
                myShape.ShapeProperties.Transform2D.Append(new D.Offset {
                            X = 0,
                            Y = 0 });
                myShape.ShapeProperties.Transform2D.Append(new D.Extents {
                            Cx = 1000000,
                            Cy = 1000000 });
                myShape.ShapeProperties.Append(new PresetGeometry {
                            Preset = ShapeTypeValues.Rectangle });
                myShape.ShapeProperties.Append(new SolidFill {
                            SchemeColor = new SchemeColor {
                                Val = SchemeColorValues.Accent2 }});
                myShape.ShapeProperties.Append(new Outline(new NoFill()));

                myShapeTree.AppendChild(myShape);
            }
        }
    }
}

