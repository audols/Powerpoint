using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using PowerPointGenerator.Models;

namespace PowerPointGenerator.Services
{
    /// <summary>
    /// Parser for JSON slide content format
    /// </summary>
    public static class JsonSlideParser
    {
        /// <summary>
        /// Parses JSON slide content from a file
        /// </summary>
        /// <param name="jsonFilePath">Path to the JSON file</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Parsed presentation content</returns>
        public static PresentationContent ParseFromFile(
            string jsonFilePath,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"JSON file not found: {jsonFilePath}");

            var jsonContent = File.ReadAllText(jsonFilePath);
            return ParseFromString(jsonContent, presentationTitle, author, imageBasePath);
        }

        /// <summary>
        /// Parses JSON slide content from a string
        /// </summary>
        /// <param name="jsonContent">JSON content as string</param>
        /// <param name="presentationTitle">Title of the presentation</param>
        /// <param name="author">Author of the presentation</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Parsed presentation content</returns>
        public static PresentationContent ParseFromString(
            string jsonContent,
            string presentationTitle = "JSON Generated Presentation",
            string author = "AI Assistant",
            string? imageBasePath = null)
        {
            imageBasePath ??= Path.Combine(Environment.CurrentDirectory, "Images");

            // Parse JSON
            var jsonSlideContent = JsonSerializer.Deserialize<JsonSlideContent>(jsonContent, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (jsonSlideContent?.Slides == null || !jsonSlideContent.Slides.Any())
                throw new InvalidOperationException("No slides found in JSON content");

            // Convert to presentation content
            var presentationContent = new PresentationContent
            {
                Title = presentationTitle,
                Author = author
            };

            foreach (var jsonSlide in jsonSlideContent.Slides)
            {
                var slideContent = new SlideContent
                {
                    Title = jsonSlide.Title,
                    Subtitle = jsonSlide.Subtitle,
                    Description = jsonSlide.Description,
                    LayoutType = ParseLayoutType(jsonSlide.Layout),
                    Template = jsonSlide.Template ?? "DETAIL"
                };

                Console.WriteLine($"Processing slide: {slideContent.Title} with layout {slideContent.LayoutType}");
                Console.WriteLine($"Images found: {slideContent.Images.Count}");

                // Parse image from suggested_image field
                if (jsonSlide.Images.Any())
                {
                    foreach (var jsonImage in jsonSlide.Images)
                    {
                        var imagePath = Path.Combine(imageBasePath, jsonImage.FilePath);
                        Console.WriteLine($"Checking image path: {imagePath}");
                        Console.WriteLine($"Image exists: {File.Exists(imagePath)}");
                        if (File.Exists(imagePath))
                        {
                            slideContent.Images.Add(new ImageContent
                            {
                                FilePath = imagePath,
                                Title = jsonImage.Title,
                                Subtitle = jsonImage.Subtitle
                            });
                        }
                        else
                        {
                            Console.WriteLine($"Image file not found: {imagePath}");
                        }
                    }
                }

                presentationContent.Slides.Add(slideContent);
            }

            return presentationContent;
        }

        /// <summary>
        /// Extracts the image filename from the suggested_image field
        /// </summary>
        /// <param name="suggestedImage">The suggested image text (e.g., "Use Image 1: filename.png")</param>
        /// <param name="imageBasePath">Base path for images</param>
        /// <returns>Full path to the image file</returns>
        private static string ExtractImagePathFromSuggestion(string suggestedImage, string imageBasePath)
        {
            // Extract filename from patterns like "Use Image 1: filename.png" or just "filename.png"
            var patterns = new[]
            {
                @"Use Image \d+:\s*""?([^""]+\.(?:png|jpg|jpeg|gif|bmp))""?",  // "Use Image 1: "filename.png""
                @"Use Image \d+:\s*([^""]+\.(?:png|jpg|jpeg|gif|bmp))",        // Use Image 1: filename.png
                @"""([^""]+\.(?:png|jpg|jpeg|gif|bmp))""",                     // "filename.png"
                @"([^""]+\.(?:png|jpg|jpeg|gif|bmp))"                          // filename.png
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(suggestedImage, pattern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var filename = match.Groups[1].Value.Trim();
                    var fullPath = Path.Combine(imageBasePath, filename);
                    return fullPath;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Parses layout type from string
        /// </summary>
        /// <param name="layoutString">Layout string from JSON</param>
        /// <returns>Corresponding SlideLayoutType</returns>
        private static SlideLayoutType ParseLayoutType(string layoutString)
        {
            if (string.IsNullOrWhiteSpace(layoutString))
                return SlideLayoutType.ImageFocused; // Default

            return layoutString.ToLowerInvariant() switch
            {
                "title" => SlideLayoutType.Title,
                "titleandcontent" or "title_and_content" or "title-and-content" => SlideLayoutType.TitleAndContent,
                "imagefocused" or "image_focused" or "image-focused" => SlideLayoutType.ImageFocused,
                "imagegrid" or "image_grid" or "image-grid" => SlideLayoutType.ImageGrid,
                "singleimagewithcaption" or "single_image_with_caption" or "single-image-with-caption" => SlideLayoutType.SingleImageWithCaption,
                "twoimagecomparison" or "two_image_comparison" or "two-image-comparison" => SlideLayoutType.TwoImageComparison,
                "productshowcase" or "product_showcase" or "product-showcase" => SlideLayoutType.ProductShowcase,
                _ => SlideLayoutType.ImageFocused // Default fallback
            };
        }
    }
}
