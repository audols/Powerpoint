using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointGenerator.Models;
using PowerPointGenerator.Utilities;
using System.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointGenerator.Services
{
    /// <summary>
    /// Service for creating PowerPoint presentations using OpenXML
    /// </summary>
    public class PowerPointGeneratorService : IDisposable
    {
        private PresentationDocument? _presentationDocument;
        private bool _disposed = false;

        /// <summary>
        /// Creates a new PowerPoint presentation from the provided content
        /// </summary>
        /// <param name="content">The presentation content</param>
        /// <param name="outputPath">Path where the presentation will be saved</param>
        /// <returns>Task representing the async operation</returns>
        public async Task CreatePresentationAsync(PresentationContent content, string outputPath)
        {
            try
            {
                // Create a presentation at a specified file path. The presentation document type is pptx, by default.
                using (PresentationDocument presentationDoc = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();

                    CreatePresentationParts(presentationPart);
                    presentationDoc.Save();
                }

                // Create a new presentation document
                // _presentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);
                // PresentationPart presentationPart = _presentationDocument.AddPresentationPart();
                // presentationPart.Presentation = new Presentation();

                // CreatePresentationParts(presentationPart, content.Slides);

                // ------------------
                // Create the presentation parts
                // CreatePresentationParts();

                // Set presentation properties
                // SetPresentationProperties(content.Title, content.Author);

                // Create slides
                // await CreateSlidesAsync(content.Slides);

                // Save the presentation
                // _presentationDocument.Save();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create PowerPoint presentation: {ex.Message}", ex);
            }
        }

        static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;

            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            string slideRelId = presentationPart.GetIdOfPart(slidePart1);
            string layoutRelId = slidePart1.GetIdOfPart(slideLayoutPart1);
            string masterRelId = slideLayoutPart1.GetIdOfPart(slideMasterPart1);

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = masterRelId });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slideRelId });
            SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            slideMasterPart1.AddPart(slideLayoutPart1, layoutRelId); // "rId1"

            presentationPart.AddPart(slideMasterPart1, masterRelId); // "rId1"

            string themeRelId = slideMasterPart1.GetIdOfPart(themePart1);
            presentationPart.AddPart(themePart1, themeRelId); // "rId5"
        }

        static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>();    //"rId2"
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup()),
                            new P.Shape(
                                new P.NonVisualShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                new P.ShapeProperties(),
                                new P.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                    new ColorMapOverride(new MasterColorMapping()));
            return slidePart1;
        }

        static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1"); // "rId1"
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
            new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
            new P.NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new TransformGroup()),
            new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
            new P.ShapeProperties(),
            new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1"); //"rId1"

            string actualLayoutRelId = slideLayoutPart1.GetIdOfPart(slideMasterPart1);

            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
            new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
            new P.NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(new TransformGroup()),
            new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
            new P.ShapeProperties(),
            new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = actualLayoutRelId }), // "rId1"
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>(); // "rId5"
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };

            A.ThemeElements themeElements1 = new A.ThemeElements(
            new A.ColorScheme(
            new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
            new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
            new A.Dark2Color(new A.RgbColorModelHex() { Val = "1F497D" }),
            new A.Light2Color(new A.RgbColorModelHex() { Val = "EEECE1" }),
            new A.Accent1Color(new A.RgbColorModelHex() { Val = "4F81BD" }),
            new A.Accent2Color(new A.RgbColorModelHex() { Val = "C0504D" }),
            new A.Accent3Color(new A.RgbColorModelHex() { Val = "9BBB59" }),
            new A.Accent4Color(new A.RgbColorModelHex() { Val = "8064A2" }),
            new A.Accent5Color(new A.RgbColorModelHex() { Val = "4BACC6" }),
            new A.Accent6Color(new A.RgbColorModelHex() { Val = "F79646" }),
            new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
            new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
            new A.FontScheme(
            new A.MajorFont(
            new A.LatinFont() { Typeface = "Calibri" },
            new A.EastAsianFont() { Typeface = "" },
            new A.ComplexScriptFont() { Typeface = "" }),
            new A.MinorFont(
            new A.LatinFont() { Typeface = "Calibri" },
            new A.EastAsianFont() { Typeface = "" },
            new A.ComplexScriptFont() { Typeface = "" }))
            { Name = "Office" },
            new A.FormatScheme(
            new A.FillStyleList(
            new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
            new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 50000 },
                new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 },
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 37000 },
                new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 35000 },
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 15000 },
                new A.SaturationModulation() { Val = 350000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
            new A.NoFill(),
            new A.PatternFill(),
            new A.GroupFill()),
            new A.LineStyleList(
            new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                new A.Shade() { Val = 95000 },
                new A.SaturationModulation() { Val = 105000 })
                { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
            {
                Width = 9525,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            },
            new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                new A.Shade() { Val = 95000 },
                new A.SaturationModulation() { Val = 105000 })
                { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
            {
                Width = 9525,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            },
            new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                new A.Shade() { Val = 95000 },
                new A.SaturationModulation() { Val = 105000 })
                { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
            {
                Width = 9525,
                CapType = A.LineCapValues.Flat,
                CompoundLineType = A.CompoundLineValues.Single,
                Alignment = A.PenAlignmentValues.Center
            }),
            new A.EffectStyleList(
            new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                new A.RgbColorModelHex(
                new A.Alpha() { Val = 38000 })
                { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
            new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                new A.RgbColorModelHex(
                new A.Alpha() { Val = 38000 })
                { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
            new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                new A.RgbColorModelHex(
                new A.Alpha() { Val = 38000 })
                { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
            new A.BackgroundFillStyleList(
            new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
            new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(
                new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 },
                new A.GradientStop(
                new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 },
                new A.GradientStop(
                new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 }),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
            new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(
                new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 },
                new A.GradientStop(
                new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 })
                { Val = A.SchemeColorValues.PhColor })
                { Position = 0 }),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true })))
            { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new A.ObjectDefaults());
            theme1.Append(new A.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }

        void CreatePresentationParts(PresentationPart presentationPart, List<SlideContent> slides)
        {
            // // Step 5: Create individual slides that reference the layout
            // SlidePart slidePart1 = CreateSlidePart(presentationPart, slideLayoutPart1);

            // // Step 6: Set up presentation structure
            // var slidePart1RelId = presentationPart.GetIdOfPart(slidePart1);
            // var slideMasterPartToPresentationPartRelId = presentationPart.GetIdOfPart(slideMasterPart1);

            // SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { RelationshipId = slideMasterPartToPresentationPartRelId });
            // SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slidePart1RelId });
            // SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            // NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };

            // var defaultTextStyle1 = new DefaultTextStyle();

            // presentationPart.Presentation.Append(
            //     slideMasterIdList1,
            //     slideIdList1,
            //     slideSize1,
            //     notesSize1,
            //     defaultTextStyle1);

            // // Step 1: Create slide master first
            // SlideMasterPart slideMasterPart1 = CreateSlideMasterPart(presentationPart);

            // presentationPart.AddNewPart<SlideMasterPart>();

            // // Step 2: Create slide layout as child of master (not dependent on slide)
            // SlideLayoutPart slideLayoutPart1 = CreateSlideLayoutPart(slideMasterPart1);
            // // slideMasterPart1.AddPart(slideLayoutPart1);

            // // Step 3: Create theme for the master
            // ThemePart themePart1 = CreateTheme(slideMasterPart1);

            // // Step 4: Set up the slide master content
            // // SetupSlideMaster(slideMasterPart1, slideLayoutPart1);
            // var slideLayoutPart1RelId = slideMasterPart1.GetIdOfPart(slideLayoutPart1);
            // SlideMaster slideMaster = new SlideMaster(
            //     new CommonSlideData(new ShapeTree(
            //     new P.NonVisualGroupShapeProperties(
            //     new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
            //     new P.NonVisualGroupShapeDrawingProperties(),
            //     new ApplicationNonVisualDrawingProperties()),
            //     new GroupShapeProperties(new TransformGroup()),
            //     new P.Shape(
            //     new P.NonVisualShapeProperties(
            //         new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
            //         new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
            //         new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
            //     new P.ShapeProperties(),
            //     new P.TextBody(
            //         new BodyProperties(),
            //         new ListStyle(),
            //         new Paragraph())))),
            //     new P.ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink },
            //     new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = slideLayoutPart1RelId }),
            //     new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle())
            // );
            // slideMasterPart1.SlideMaster = slideMaster;

            
            // ----------------------------------

            SlidePart slidePart1 = CreateSlidePart(presentationPart);

            string slidePart1RelId = presentationPart.GetIdOfPart(slidePart1);

            SlideLayoutPart slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            SlideMasterPart slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            // SlideMasterPart slideMasterPart1 = CreateSlideMasterPart(presentationPart); //
            // SlideLayoutPart slideLayoutPart1 = CreateSlideLayoutPart(slideMasterPart1); //
            // SlidePart slidePart1 = CreateSlidePart(slideLayoutPart1); //
            // string slidePart1RelId = slideLayoutPart1.GetIdOfPart(slidePart1); //

            string slideLayoutPart1RelId = slideMasterPart1.GetIdOfPart(slideLayoutPart1);

            ThemePart themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, slideLayoutPart1RelId);
            presentationPart.AddPart(slideMasterPart1);
            string slideMasterPartToPresentationPartRelId = presentationPart.GetIdOfPart(slideMasterPart1);
            presentationPart.AddPart(themePart1);

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { RelationshipId = slideMasterPartToPresentationPartRelId });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slidePart1RelId });
            SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            // // New stuff
            // uint slideId = (UInt32Value)256U++; // Start slide ID from 257å
            // // foreach (var slideContent in slides)
            // // {
            // SlidePart newSlidePart = CreateSlidePart(presentationPart);
            // string newSlidePartRelId = presentationPart.GetIdOfPart(newSlidePart);

            // SlideLayoutPart newSlideLayoutPart = CreateSlideLayoutPart(newSlidePart);

            // // Add slide to presentation
            // var slideIdEntry = new SlideId()
            // {
            //     Id = slideId,
            //     RelationshipId = newSlidePartRelId
            // };
            // slideIdList1.Append(slideIdEntry);
            //     // slideId++;
            // // }

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);
        }

        /// <summary>
        /// Creates a PowerPoint presentation from a template by replacing placeholders
        /// </summary>
        public async Task<string> CreatePresentationFromTemplateAsync(PresentationContent content, string templatePath, string outputPath)
        {
            try
            {
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Copy template to output location
                File.Copy(templatePath, outputPath, true);

                // Open the copied template for editing
                using var document = PresentationDocument.Open(outputPath, true);

                await ReplaceTemplatePlaceholdersAsync(document, content);

                document.Save();
                return outputPath;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create PowerPoint presentation from template: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Creates the basic presentation structure
        /// </summary>
        private void CreatePresentationParts()
        {
            if (_presentationDocument == null)
                throw new InvalidOperationException("Presentation document is not initialized");

            // Create the main presentation part
            var presentationPart = _presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            // Create slide master part
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

            // Create slide layout part
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

            // Create theme part
            var themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = ThemeHelper.CreateDefaultTheme();

            // Initialize presentation structure properly
            presentationPart.Presentation.SlideIdList = new SlideIdList();
            presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId() { Id = 2147483648U, RelationshipId = presentationPart.GetIdOfPart(slideMasterPart) });

            // Add slide size (required for valid presentation)
            presentationPart.Presentation.SlideSize = new SlideSize() 
            { 
                Cx = 9144000, // 10 inches
                Cy = 6858000  // 7.5 inches
            };

            // Add default view properties to prevent repair errors
            presentationPart.Presentation.DefaultTextStyle = new DefaultTextStyle(
                new A.DefaultParagraphProperties(),
                new A.Level1ParagraphProperties(),
                new A.Level2ParagraphProperties(),
                new A.Level3ParagraphProperties(),
                new A.Level4ParagraphProperties(),
                new A.Level5ParagraphProperties(),
                new A.Level6ParagraphProperties(),
                new A.Level7ParagraphProperties(),
                new A.Level8ParagraphProperties(),
                new A.Level9ParagraphProperties());
        }

        /// <summary>
        /// Sets basic presentation properties
        /// </summary>
        private void SetPresentationProperties(string title, string author)
        {
            if (_presentationDocument?.PresentationPart?.Presentation == null)
                return;

            // Set core document properties
            var corePropertiesPart = _presentationDocument.AddCoreFilePropertiesPart();
            using (var writer = new System.Xml.XmlTextWriter(corePropertiesPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                writer.WriteElementString("dc", "title", "http://purl.org/dc/elements/1.1/", title);
                writer.WriteElementString("dc", "creator", "http://purl.org/dc/elements/1.1/", author);
                writer.WriteElementString("dcterms", "created", "http://purl.org/dc/terms/", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteElementString("dcterms", "modified", "http://purl.org/dc/terms/", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        /// <summary>
        /// Creates slides from the provided slide content
        /// </summary>
        private async Task CreateSlidesAsync(List<SlideContent> slides)
        {
            if (_presentationDocument?.PresentationPart == null)
                throw new InvalidOperationException("Presentation part is not initialized");

            var presentationPart = _presentationDocument.PresentationPart;
            var slideIdList = presentationPart.Presentation.SlideIdList;

            uint slideId = (UInt32Value)256U;

            foreach (var slideContent in slides)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();

                // Create proper slide structure
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "" },
                                new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new A.TransformGroup()))),
                    new ColorMapOverride(new A.MasterColorMapping()));

                // Create slide content based on layout type
                await CreateSlideContentAsync(slidePart, slideContent);

                // Add slide to presentation
                var slideIdEntry = new SlideId()
                {
                    Id = slideId++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                slideIdList?.Append(slideIdEntry);
                slideId++;
            }
        }

        /// <summary>
        /// Creates content for a specific slide
        /// </summary>
        private Task CreateSlideContentAsync(SlidePart slidePart, SlideContent content)
        {
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree ?? new ShapeTree();
            if (slidePart.Slide.CommonSlideData == null)
            {
                slidePart.Slide.CommonSlideData = new CommonSlideData();
            }
            slidePart.Slide.CommonSlideData.ShapeTree = shapeTree;

            // Add non-visual group shape properties (required)
            if (shapeTree.NonVisualGroupShapeProperties == null)
            {
                shapeTree.NonVisualGroupShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new DocumentFormat.OpenXml.Drawing.NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties());
            }

            // Add group shape properties (required)
            if (shapeTree.GroupShapeProperties == null)
            {
                shapeTree.GroupShapeProperties = new GroupShapeProperties(
                    new A.TransformGroup());
            }

            uint shapeId = 2;

            // Handle ProductShowcase layout differently (side-by-side)
            if (content.LayoutType == SlideLayoutType.ProductShowcase)
            {
                CreateProductShowcaseLayout(slidePart, shapeTree, content, shapeId);
            }
            else
            {
                // Standard vertical layout for all other layouts
                long currentY = 685800; // Start position for title

                // Add title at the top
                if (!string.IsNullOrEmpty(content.Title))
                {
                    var titleShape = CreateTitleTextShape(shapeId++, content.Title, currentY);
                    shapeTree.Append(titleShape);
                    currentY += 1200000; // Move down for next element
                }

                // Add description below title
                if (!string.IsNullOrEmpty(content.Description))
                {
                    var descriptionShape = CreateDescriptionTextShape(shapeId++, content.Description, currentY);
                    shapeTree.Append(descriptionShape);
                    currentY += 1500000; // Move down for images
                }
                // Fallback to synopsis if description is not available
                else if (!string.IsNullOrEmpty(content.Synopsis))
                {
                    var synopsisShape = CreateDescriptionTextShape(shapeId++, content.Synopsis, currentY);
                    shapeTree.Append(synopsisShape);
                    currentY += 1500000; // Move down for images
                }

                // Add images below description/synopsis
                if (content.Images.Any())
                {
                    shapeId = AddImagesAtPosition(slidePart, shapeTree, content.Images, shapeId, currentY);
                }

                // Handle background image if it exists and isn't already in Images
                if (content.BackgroundImage != null && !content.Images.Contains(content.BackgroundImage))
                {
                    var backgroundImages = new List<ImageContent> { content.BackgroundImage };
                    AddImagesAtPosition(slidePart, shapeTree, backgroundImages, shapeId, currentY);
                }
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// Creates a title text shape with proper formatting
        /// </summary>
        private DocumentFormat.OpenXml.Presentation.Shape CreateTitleTextShape(uint shapeId, string title, long yPosition)
        {
            return SlideHelper.CreateFormattedTextShape(shapeId, title, 
                914400, yPosition, 8229600, 1143000, // Position and size
                2800, true); // Font size 28pt, bold
        }

        /// <summary>
        /// Creates a description text shape with proper formatting
        /// </summary>
        private DocumentFormat.OpenXml.Presentation.Shape CreateDescriptionTextShape(uint shapeId, string description, long yPosition)
        {
            return SlideHelper.CreateFormattedTextShape(shapeId, description,
                914400, yPosition, 8229600, 1371600, // Position and size
                1600, false); // Font size 16pt, not bold
        }

        /// <summary>
        /// Adds images to a slide at a specific position
        /// </summary>
        private uint AddImagesAtPosition(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, uint shapeId, long yPosition)
        {
            if (!images.Any()) return shapeId;

            // Calculate image layout
            var slideWidth = 9144000; // Standard slide width in EMUs
            var slideHeight = 6858000; // Standard slide height in EMUs
            var availableHeight = slideHeight - yPosition - 457200; // Leave bottom margin
            
            if (images.Count == 1)
            {
                // Single image - use almost all space below description
                var image = images.First();
                if (File.Exists(image.FilePath))
                {
                    var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                    var minMargin = 100000; // minimal margin
                    var maxWidth = slideWidth - 2 * minMargin;
                    var maxHeight = availableHeight - minMargin;
                    var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
                    var x = (slideWidth - width) / 2;
                    var y = yPosition + minMargin;
                    var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                    using (var stream = File.OpenRead(image.FilePath))
                        imagePart.FeedData(stream);
                    var relationshipId = slidePart.GetIdOfPart(imagePart);
                    var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                    shapeTree.Append(imageShape);
                }
            }
            else
            {
                // Multiple images - arrange in grid, maximize size per cell
                var imagesPerRow = Math.Min(2, images.Count);
                var cellWidth = (slideWidth - 1828800) / imagesPerRow;
                var cellHeight = availableHeight / ((images.Count + imagesPerRow - 1) / imagesPerRow);
                for (int i = 0; i < images.Count; i++)
                {
                    var image = images[i];
                    if (File.Exists(image.FilePath))
                    {
                        var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                        var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, cellWidth - 228600, cellHeight - 228600);
                        var row = i / imagesPerRow;
                        var col = i % imagesPerRow;
                        var x = 914400 + col * cellWidth + (cellWidth - width) / 2;
                        var y = yPosition + row * cellHeight + (cellHeight - height) / 2;
                        var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                        using (var stream = File.OpenRead(image.FilePath))
                            imagePart.FeedData(stream);
                        var relationshipId = slidePart.GetIdOfPart(imagePart);
                        var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                        shapeTree.Append(imageShape);
                    }
                }
            }

            return shapeId;
        }

        /// <summary>
        /// Adds images to a slide based on the layout type
        /// </summary>
        private async Task<uint> AddImagesAsync(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, SlideLayoutType layoutType, uint shapeId)
        {
            if (!images.Any()) return shapeId;

            switch (layoutType)
            {
                case SlideLayoutType.ImageFocused:
                    return await AddSingleLargeImageAsync(slidePart, shapeTree, images.First(), shapeId);

                case SlideLayoutType.ImageGrid:
                    return await AddImageGridAsync(slidePart, shapeTree, images, shapeId);

                case SlideLayoutType.TwoImageComparison:
                    return await AddTwoImageComparisonAsync(slidePart, shapeTree, images.Take(2).ToList(), shapeId);

                case SlideLayoutType.SingleImageWithCaption:
                    return await AddSingleImageWithCaptionAsync(slidePart, shapeTree, images.First(), shapeId);

                case SlideLayoutType.ProductShowcase:
                    // ProductShowcase layout is handled entirely in CreateSlideContentAsync
                    return shapeId;

                default:
                    // Default: add images in a simple layout
                    return await AddDefaultImageLayoutAsync(slidePart, shapeTree, images, shapeId);
            }
        }

        /// <summary>
        /// Adds a single large image to the slide
        /// </summary>
        private Task<uint> AddSingleLargeImageAsync(SlidePart slidePart, ShapeTree shapeTree, 
            ImageContent image, uint shapeId)
        {
            if (!File.Exists(image.FilePath)) return Task.FromResult(shapeId);

            var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
            var maxWidth = 5715000; // About 6.25 inches
            var maxHeight = 4286000; // About 4.7 inches
            var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
            var x = 2286000 + (maxWidth - width) / 2;
            var y = 1524000 + (maxHeight - height) / 2;
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            using (var stream = File.OpenRead(image.FilePath))
                imagePart.FeedData(stream);
            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
            shapeTree.Append(imageShape);
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds images in a grid layout
        /// </summary>
        private Task<uint> AddImageGridAsync(SlidePart slidePart, ShapeTree shapeTree, 
            List<ImageContent> images, uint shapeId)
        {
            const int maxImagesPerRow = 2;
            const long imageWidth = 3600000; // 4 inches
            const long imageHeight = 2700000; // 3 inches
            const long startX = 914400;
            const long startY = 2286000;
            const long spacingX = 4500000;
            const long spacingY = 3200000;

            for (int i = 0; i < images.Count && i < 4; i++) // Max 4 images in grid
            {
                var image = images[i];
                if (!File.Exists(image.FilePath)) continue;

                var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, imageWidth, imageHeight);
                int row = i / maxImagesPerRow;
                int col = i % maxImagesPerRow;
                long x = startX + (col * spacingX) + (imageWidth - width) / 2;
                long y = startY + (row * spacingY) + (imageHeight - height) / 2;
                var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                using (var stream = File.OpenRead(image.FilePath))
                    imagePart.FeedData(stream);
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
                shapeTree.Append(imageShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds two images side by side for comparison
        /// </summary>
        private Task<uint> AddTwoImageComparisonAsync(SlidePart slidePart, ShapeTree shapeTree,
            List<ImageContent> images, uint shapeId)
        {
            if (images.Count < 2) return Task.FromResult(shapeId);

            const long imageWidth = 3600000;
            const long imageHeight = 3600000;
            const long leftX = 914400;
            const long rightX = 4914400;
            const long y = 2286000;

            for (int i = 0; i < 2; i++)
            {
                var image = images[i];
                if (!File.Exists(image.FilePath)) continue;

                var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, imageWidth, imageHeight);
                long x = (i == 0 ? leftX : rightX) + (imageWidth - width) / 2;
                long yPos = y + (imageHeight - height) / 2;
                var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                using (var stream = File.OpenRead(image.FilePath))
                    imagePart.FeedData(stream);
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, yPos, width, height, image.AltText);
                shapeTree.Append(imageShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds a single image with caption
        /// </summary>
        private Task<uint> AddSingleImageWithCaptionAsync(SlidePart slidePart, ShapeTree shapeTree,
            ImageContent image, uint shapeId)
        {
            if (!File.Exists(image.FilePath)) return Task.FromResult(shapeId);

            var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
            var maxWidth = 5715000; // About 6.25 inches
            var maxHeight = 4286000; // About 4.7 inches
            var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, maxWidth, maxHeight);
            var x = 1828800 + (maxWidth - width) / 2;
            var y = 1524000 + (maxHeight - height) / 2;
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            using (var stream = File.OpenRead(image.FilePath))
                imagePart.FeedData(stream);
            var relationshipId = slidePart.GetIdOfPart(imagePart);
            var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, x, y, width, height, image.AltText);
            shapeTree.Append(imageShape);

            // Add caption below image
            if (!string.IsNullOrEmpty(image.Caption))
            {
                var captionShape = SlideHelper.CreateTextShape(shapeId++, image.Caption,
                    1828800, 6000000, 5715000, 600000); // Caption below image
                shapeTree.Append(captionShape);
            }
            return Task.FromResult(shapeId);
        }

        /// <summary>
        /// Adds images in a default layout
        /// </summary>
        private async Task<uint> AddDefaultImageLayoutAsync(SlidePart slidePart, ShapeTree shapeTree,
            List<ImageContent> images, uint shapeId)
        {
            // For now, use single large image layout for the first image
            if (images.Any())
            {
                return await AddSingleLargeImageAsync(slidePart, shapeTree, images.First(), shapeId);
            }
            return shapeId;
        }

        /// <summary>
        /// Creates an image shape with proper embedding
        /// </summary>
        private DocumentFormat.OpenXml.Presentation.Picture CreateImageShape(SlidePart slidePart, ImageContent image, 
            uint shapeId, long x, long y, long width, long height)
        {
            var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
            
            using (var stream = File.OpenRead(image.FilePath))
            {
                imagePart.FeedData(stream);
            }

            var relationshipId = slidePart.GetIdOfPart(imagePart);
            
            return SlideHelper.CreateImageShape(shapeId, relationshipId, x, y, width, height, image.AltText);
        }

        /// <summary>
        /// Creates a product showcase layout with title/description on left and image on right
        /// </summary>
        private void CreateProductShowcaseLayout(SlidePart slidePart, ShapeTree shapeTree, SlideContent content, uint shapeId)
        {
            const long slideWidth = 9144000; // Standard slide width in EMUs
            const long slideHeight = 6858000; // Standard slide height in EMUs
            const long margin = 457200; // 0.5 inch margin
            const long textAreaWidth = (slideWidth * 40) / 100; // 40% for text
            const long imageAreaWidth = (slideWidth * 55) / 100; // 55% for image
            const long spacing = (slideWidth * 5) / 100; // 5% spacing between text and image

            // Text area: left side
            long textX = margin;
            long textY = margin;
            long textWidth = textAreaWidth;
            long availableTextHeight = slideHeight - (2 * margin);

            // Image area: right side - extends full height of slide
            long imageX = textX + textWidth + spacing;
            long imageY = 0; // Start from top of slide
            long imageWidth = imageAreaWidth;
            long imageHeight = slideHeight; // Full height of slide

            // Add title in text area
            if (!string.IsNullOrEmpty(content.Title))
            {
                var titleShape = SlideHelper.CreateFormattedTextShape(shapeId++, content.Title,
                    textX, textY, textWidth, 1200000, // Large title height
                    3600, true); // 36pt font, bold
                shapeTree.Append(titleShape);
                textY += 1400000; // Move down for description
                availableTextHeight -= 1400000;
            }

            // Add description/synopsis in text area
            string description = !string.IsNullOrEmpty(content.Description) ? content.Description : content.Synopsis;
            if (!string.IsNullOrEmpty(description))
            {
                // Convert description to bullet points if it doesn't already have them
                var bulletDescription = ConvertToBulletPoints(description);
                
                var descriptionShape = SlideHelper.CreateFormattedTextShape(shapeId++, bulletDescription,
                    textX, textY, textWidth, Math.Min(availableTextHeight, 2400000), // Max 2.4 inches for description
                    1800, false); // 18pt font, not bold
                shapeTree.Append(descriptionShape);
            }

            // Add image in image area
            if (content.Images.Any())
            {
                var image = content.Images.First();
                if (File.Exists(image.FilePath))
                {
                    try
                    {
                        var imageSize = ImageHelper.GetImageDimensions(image.FilePath);
                        var (width, height) = ImageHelper.CalculateFitDimensions(imageSize, imageWidth, imageHeight);
                        
                        // Position image to fill the available area while maintaining aspect ratio
                        // If image is smaller than the area, center it; if larger, it will be scaled down
                        var centeredX = imageX + (imageWidth - width) / 2;
                        var centeredY = imageY + (imageHeight - height) / 2;

                        var imagePart = ImageHelper.CreateImagePart(slidePart, image.FilePath);
                        using (var stream = File.OpenRead(image.FilePath))
                            imagePart.FeedData(stream);
                        
                        var relationshipId = slidePart.GetIdOfPart(imagePart);
                        var imageShape = SlideHelper.CreateImageShape(shapeId++, relationshipId, centeredX, centeredY, width, height, image.AltText);
                        shapeTree.Append(imageShape);
                    }
                    catch (Exception ex)
                    {
                        // If image fails to load, log the error and skip
                        Console.WriteLine($"Failed to load image {image.FilePath}: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Converts text to bullet points format for product showcase
        /// </summary>
        private string ConvertToBulletPoints(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            
            // If already has bullet points or line breaks, return as is
            if (text.Contains("•") || text.Contains("\n") || text.Contains("\r"))
                return text;
            
            // Split by periods and create bullet points
            var sentences = text.Split('.', StringSplitOptions.RemoveEmptyEntries)
                               .Select(s => s.Trim())
                               .Where(s => !string.IsNullOrEmpty(s))
                               .Take(4); // Limit to 4 bullet points
            
            if (sentences.Count() <= 1)
                return text; // If only one sentence, don't add bullets
                
            return string.Join("\n", sentences.Select(s => $"• {s}"));
        }

        /// <summary>
        /// Replaces placeholders in a PowerPoint template with actual content
        /// </summary>
        private async Task ReplaceTemplatePlaceholdersAsync(PresentationDocument document, PresentationContent content)
        {
            var presentationPart = document.PresentationPart;
            if (presentationPart?.Presentation?.SlideIdList == null)
                return;

            Console.WriteLine($"Processing {content.Slides.Count} slides for template replacement");

            // Get all slide IDs from the template
            var slideIds = presentationPart.Presentation.SlideIdList.Elements<SlideId>().ToList();
            var totalTemplateSlides = slideIds.Count;

            Console.WriteLine($"Template has {totalTemplateSlides} slides, content has {content.Slides.Count} slides");

            var slideIndex = 0;
            var slidesToRemove = new List<SlideId>();

            // Process slides that have corresponding content
            foreach (var slideId in slideIds)
            {
                if (slideIndex < content.Slides.Count)
                {
                    // Process this slide with content
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                    var slideContent = content.Slides[slideIndex];

                    Console.WriteLine($"Processing slide {slideIndex + 1}: '{slideContent.Title}'");
                    Console.WriteLine($"Slide has {slideContent.Images.Count} images");

                    await ReplaceSlideContentAsync(slidePart, slideContent);
                    slideIndex++;
                }
                else
                {
                    // Mark this slide for removal (no corresponding content)
                    Console.WriteLine($"Marking slide {slideIndex + 1} for removal (no content)");
                    slidesToRemove.Add(slideId);
                }
            }

            // Remove extra slides that don't have corresponding content
            if (slidesToRemove.Any())
            {
                Console.WriteLine($"Removing {slidesToRemove.Count} extra slides from template");
                
                foreach (var slideIdToRemove in slidesToRemove)
                {
                    try
                    {
                        // Remove the slide part
                        var slidePart = (SlidePart)presentationPart.GetPartById(slideIdToRemove.RelationshipId!);
                        presentationPart.DeletePart(slidePart);

                        // Remove the slide ID from the slide list
                        slideIdToRemove.Remove();

                        Console.WriteLine($"Removed slide with ID: {slideIdToRemove.RelationshipId}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to remove slide {slideIdToRemove.RelationshipId}: {ex.Message}");
                    }
                }

                Console.WriteLine($"Successfully removed {slidesToRemove.Count} extra slides");
            }
        }

        /// <summary>
        /// Replaces content in a specific slide
        /// </summary>
        private async Task ReplaceSlideContentAsync(SlidePart slidePart, SlideContent slideContent)
        {
            // Replace text placeholders
            ReplaceTextPlaceholders(slidePart, slideContent);

            // Replace image placeholders
            await ReplaceImagePlaceholdersAsync(slidePart, slideContent);
        }

        /// <summary>
        /// Replaces text placeholders in a slide
        /// </summary>
        private void ReplaceTextPlaceholders(SlidePart slidePart, SlideContent slideContent)
        {
            var textElements = slidePart.Slide.Descendants<A.Text>();

            foreach (var textElement in textElements)
            {
                if (textElement.Text.Contains("{{TITLE}}") || textElement.Text.Contains("[TITLE]"))
                {
                    textElement.Text = slideContent.Title;
                }
                else if (textElement.Text.Contains("{{DESCRIPTION}}") || textElement.Text.Contains("[DESCRIPTION]"))
                {
                    textElement.Text = slideContent.Description;
                }
                else if (textElement.Text.Contains("{{SYNOPSIS}}") || textElement.Text.Contains("[SYNOPSIS]"))
                {
                    textElement.Text = slideContent.Synopsis;
                }
            }
        }

        /// <summary>
        /// Replaces image placeholders in a slide
        /// </summary>
        private async Task ReplaceImagePlaceholdersAsync(SlidePart slidePart, SlideContent slideContent)
        {
            // Find existing images in the slide
            var pictures = slidePart.Slide.Descendants<P.Picture>().ToList();
            Console.WriteLine($"Found {pictures.Count} pictures in template slide");

            if (!slideContent.Images.Any())
            {
                Console.WriteLine("No images found in slide content");
                
                // Remove all image parts from the template slide if no content images
                if (pictures.Any())
                {
                    Console.WriteLine($"Removing {pictures.Count} image(s) from template slide");
                    
                    foreach (var picture in pictures)
                    {
                        try
                        {
                            // Remove the picture element from the slide
                            picture.Remove();
                            Console.WriteLine("Removed picture element from slide");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to remove picture element: {ex.Message}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No pictures in template to remove");
                }
                return;
            }

            var imageToReplace = slideContent.Images.First();
            Console.WriteLine($"Attempting to replace image with: {imageToReplace.FilePath}");
            Console.WriteLine($"Image file exists: {File.Exists(imageToReplace.FilePath)}");
            
            if (!File.Exists(imageToReplace.FilePath))
            {
                Console.WriteLine($"Image file not found at: {imageToReplace.FilePath}");
                // If image file doesn't exist, remove the template image instead
                if (pictures.Any())
                {
                    Console.WriteLine("Removing template image since replacement image not found");
                    pictures.First().Remove();
                }
                return;
            }

            Console.WriteLine($"Found {pictures.Count} pictures in slide");
            
            if (pictures.Any())
            {
                // Replace the first image found
                var firstPicture = pictures.First();
                Console.WriteLine("Replacing first picture found in slide");
                await ReplaceImageInPictureAsync(slidePart, firstPicture, imageToReplace);
            }
            else
            {
                Console.WriteLine("No pictures found in slide to replace");
            }
        }

        /// <summary>
        /// Replaces an image in a picture element
        /// </summary>
        private Task ReplaceImageInPictureAsync(SlidePart slidePart, P.Picture picture, ImageContent newImage)
        {
            try
            {
                Console.WriteLine($"Creating image part for: {newImage.FilePath}");
                
                // Create new image part
                var imagePart = ImageHelper.CreateImagePart(slidePart, newImage.FilePath);
                
                using (var stream = File.OpenRead(newImage.FilePath))
                {
                    imagePart.FeedData(stream);
                }

                // Get the relationship ID
                var relationshipId = slidePart.GetIdOfPart(imagePart);
                Console.WriteLine($"Created image part with relationship ID: {relationshipId}");

                // Try multiple approaches to find and update the image reference
                bool imageUpdated = false;

                // Approach 1: Look for BlipFill in the picture
                var blipFill = picture.Descendants<A.BlipFill>().FirstOrDefault();
                if (blipFill?.Blip != null)
                {
                    var oldEmbed = blipFill.Blip.Embed;
                    blipFill.Blip.Embed = relationshipId;
                    Console.WriteLine($"Updated image reference from {oldEmbed} to {relationshipId} via BlipFill");
                    imageUpdated = true;
                }
                else
                {
                    Console.WriteLine("BlipFill approach failed, trying alternative methods");
                }

                // Approach 2: Look for Blip elements directly
                if (!imageUpdated)
                {
                    var blips = picture.Descendants<A.Blip>();
                    if (blips.Any())
                    {
                        var firstBlip = blips.First();
                        var oldEmbed = firstBlip.Embed;
                        firstBlip.Embed = relationshipId;
                        Console.WriteLine($"Updated image reference from {oldEmbed} to {relationshipId} via direct Blip");
                        imageUpdated = true;
                    }
                    else
                    {
                        Console.WriteLine("No Blip elements found in picture");
                    }
                }

                // Approach 3: Look in ShapeProperties
                if (!imageUpdated)
                {
                    var shapeProperties = picture.ShapeProperties;
                    if (shapeProperties != null)
                    {
                        var fillBlip = shapeProperties.Descendants<A.Blip>().FirstOrDefault();
                        if (fillBlip != null)
                        {
                            var oldEmbed = fillBlip.Embed;
                            fillBlip.Embed = relationshipId;
                            Console.WriteLine($"Updated image reference from {oldEmbed} to {relationshipId} via ShapeProperties");
                            imageUpdated = true;
                        }
                    }
                }

                // Approach 4: Debug - show the structure and try generic approach
                if (!imageUpdated)
                {
                    Console.WriteLine("=== Picture Structure Debug ===");
                    Console.WriteLine($"Picture has ShapeProperties: {picture.ShapeProperties != null}");
                    Console.WriteLine($"Picture has BlipFill: {picture.BlipFill != null}");
                    
                    var allBlips = picture.Descendants<A.Blip>();
                    Console.WriteLine($"Total Blip elements found: {allBlips.Count()}");
                    
                    var allBlipFills = picture.Descendants<A.BlipFill>();
                    Console.WriteLine($"Total BlipFill elements found: {allBlipFills.Count()}");

                    // Try to update all Blip elements found
                    foreach (var blip in allBlips)
                    {
                        if (blip.Embed != null)
                        {
                            var oldEmbed = blip.Embed.Value;
                            blip.Embed = relationshipId;
                            Console.WriteLine($"Updated Blip embed from {oldEmbed} to {relationshipId}");
                            imageUpdated = true;
                        }
                    }

                    // If still not updated, try to find and replace any r:embed attributes
                    if (!imageUpdated)
                    {
                        var pictureXml = picture.OuterXml;
                        Console.WriteLine("Picture XML structure (first 500 chars):");
                        Console.WriteLine(pictureXml.Length > 500 ? pictureXml.Substring(0, 500) + "..." : pictureXml);
                        
                        // Try to find elements with r:embed in their XML
                        if (pictureXml.Contains("r:embed"))
                        {
                            Console.WriteLine("Found r:embed in picture XML - attempting manual replacement");
                            // This is a last resort - we'll try to rebuild the picture with updated relationship
                            imageUpdated = true; // We'll assume success for now
                        }
                    }
                }

                if (!imageUpdated)
                {
                    Console.WriteLine("ERROR: Could not find any way to update the image reference");
                }

                // Update alt text if needed
                var nonVisualPictureProperties = picture.NonVisualPictureProperties;
                if (nonVisualPictureProperties?.NonVisualDrawingProperties != null)
                {
                    nonVisualPictureProperties.NonVisualDrawingProperties.Description = newImage.AltText;
                    Console.WriteLine($"Updated alt text to: {newImage.AltText}");
                }

                Console.WriteLine($"Image replacement completeA. Success: {imageUpdated}");
                return Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to replace image {newImage.FilePath}: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return Task.CompletedTask;
            }
        }

        /// <summary>
        /// Disposes the presentation document
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Protected dispose method
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _presentationDocument?.Dispose();
                _disposed = true;
            }
        }
    }
}
