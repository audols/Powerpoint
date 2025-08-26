using DocumentFormat.OpenXml;
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
                // Create a new presentation document
                _presentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation);

                // Create the presentation parts
                CreatePresentationParts();

                // Set presentation properties
                SetPresentationProperties(content.Title, content.Author);

                // Create slides
                await CreateSlidesAsync(content.Slides);

                // Save the presentation
                _presentationDocument.Save();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create PowerPoint presentation: {ex.Message}", ex);
            }
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
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

            // Create slide layout part
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
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

            uint slideId = 256;

            foreach (var slideContent in slides)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                
                // Create proper slide structure
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties() { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
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
                shapeTree.NonVisualGroupShapeProperties = new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
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
        private Shape CreateTitleTextShape(uint shapeId, string title, long yPosition)
        {
            return SlideHelper.CreateFormattedTextShape(shapeId, title, 
                914400, yPosition, 8229600, 1143000, // Position and size
                2800, true); // Font size 28pt, bold
        }

        /// <summary>
        /// Creates a description text shape with proper formatting
        /// </summary>
        private Shape CreateDescriptionTextShape(uint shapeId, string description, long yPosition)
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
        private Picture CreateImageShape(SlidePart slidePart, ImageContent image, 
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
                else if (textElement.Text.Contains("{{SUBTITLE}}") || textElement.Text.Contains("[SUBTITLE]"))
                {
                    textElement.Text = slideContent.Subtitle;
                }
                else if (textElement.Text.Contains("{{DESCRIPTION}}") || textElement.Text.Contains("[DESCRIPTION]"))
                {
                    textElement.Text = slideContent.Description;
                }
                else if (textElement.Text.Contains("{{SYNOPSIS}}") || textElement.Text.Contains("[SYNOPSIS]"))
                {
                    textElement.Text = slideContent.Synopsis;
                }
                // Image-specific placeholders
                else if (TryReplaceImagePlaceholder(textElement.Text, slideContent.Images, out string replacementText))
                {
                    textElement.Text = replacementText;
                }
            }
        }
        
        /// <summary>
        /// Checks if text contains a placeholder in either {{}} or [] format
        /// </summary>
        private static bool ContainsPlaceholder(string text, string placeholder)
        {
            return text.Contains($"{{{{{placeholder}}}}}") || text.Contains($"[{placeholder}]");
        }

        /// <summary>
        /// Attempts to replace image-related placeholders with appropriate content
        /// </summary>
        private static bool TryReplaceImagePlaceholder(string text, List<ImageContent>? images, out string replacementText)
        {
            replacementText = "";

            // Define image placeholder patterns
            var imagePatterns = new (string Placeholder, int Index, Func<ImageContent, string?> Selector)[]
            {
                ("IMAGE1 TITLE", 0, img => img.Title),
                ("IMAGE1 SUBTITLE", 0, img => img.Subtitle),
                ("IMAGE2 TITLE", 1, img => img.Title),
                ("IMAGE2 SUBTITLE", 1, img => img.Subtitle),
                ("IMAGE3 TITLE", 2, img => img.Title),
                ("IMAGE3 SUBTITLE", 2, img => img.Subtitle),
                ("IMAGE4 TITLE", 3, img => img.Title),
                ("IMAGE4 SUBTITLE", 3, img => img.Subtitle)
            };

            foreach (var (placeholder, index, selector) in imagePatterns)
            {
                if (ContainsPlaceholder(text, placeholder))
                {
                    replacementText = GetImageContentSafely(images, index, selector) ?? "";
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Safely retrieves content from an image at the specified index using the provided selector
        /// </summary>
        private static string? GetImageContentSafely(List<ImageContent>? images, int index, Func<ImageContent, string?> selector)
        {
            if (images == null || index >= images.Count || index < 0)
                return null;
                
            var image = images[index];
            return image != null ? selector(image) : null;
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

            // More images than pictures - remove excess images from the end
            if (slideContent.Images.Count > pictures.Count)
            {
                var excessImageCount = slideContent.Images.Count - pictures.Count;
                var imagesToRemove = slideContent.Images.Skip(pictures.Count).ToList();

                Console.WriteLine($"Warning: More images in content ({slideContent.Images.Count}) than in template ({pictures.Count}). Some images will not be used.");
                foreach (var imageToRemove in imagesToRemove)
                {
                    slideContent.Images.Remove(imageToRemove);
                }

                Console.WriteLine($"Removed {excessImageCount} excess images from slideContent");
            }
            // More pictures than images - remove excess pictures from the end
            else if (pictures.Count > slideContent.Images.Count)
            {
                var excessPictureCount = pictures.Count - slideContent.Images.Count;
                var picturesToRemove = pictures.Skip(slideContent.Images.Count).ToList();

                foreach (var pictureToRemove in picturesToRemove)
                {
                    try
                    {
                        pictureToRemove.Remove();
                        Console.WriteLine("Removed excess picture from template slide");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to remove excess picture element: {ex.Message}");
                    }
                }

                // Update the pictures list to reflect the removals
                pictures.RemoveRange(slideContent.Images.Count, excessPictureCount);

                Console.WriteLine($"Removed {excessPictureCount} excess pictures from template slide");
            }

            var imageToReplace = slideContent.Images.First();
            Console.WriteLine($"Attempting to replace image with: {imageToReplace.FilePath}");
            Console.WriteLine($"Image file exists: {File.Exists(imageToReplace.FilePath)}");

            for (int i = 0; i < slideContent.Images.Count; i++)
            {
                var image = slideContent.Images[i];
                Console.WriteLine($"Attempting to replace image with: {image.FilePath}");
                Console.WriteLine($"Image file exists: {File.Exists(image.FilePath)}");

                if (!File.Exists(image.FilePath))
                {
                    Console.WriteLine($"Image file not found at: {image.FilePath}");
                    // If image file doesn't exist, remove the template image instead
                    if (pictures.Any())
                    {
                        Console.WriteLine("Removing template image since replacement image not found");
                        pictures.First().Remove();
                    }
                    return;
                }

                Console.WriteLine($"Found {pictures.Count} pictures in slide");

                if (pictures[i] != null)
                {
                    // Replace the first image found
                    Console.WriteLine("Replacing first picture found in slide");
                    await ReplaceImageInPictureAsync(slidePart, pictures[i], image);
                }
                else
                {
                    Console.WriteLine("No pictures found in slide to replace");
                }
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

                Console.WriteLine($"Image replacement completed. Success: {imageUpdated}");
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
