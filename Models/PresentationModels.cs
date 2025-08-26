namespace PowerPointGenerator.Models
{
    /// <summary>
    /// Represents the input content for generating a PowerPoint presentation
    /// </summary>
    public class PresentationContent
    {
        /// <summary>
        /// The title of the presentation
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// The author of the presentation
        /// </summary>
        public string Author { get; set; } = string.Empty;

        /// <summary>
        /// Collection of slides to be included in the presentation
        /// </summary>
        public List<SlideContent> Slides { get; set; } = new List<SlideContent>();
    }

    /// <summary>
    /// Represents the content for a single slide
    /// </summary>
    public class SlideContent
    {
        /// <summary>
        /// The title of the slide
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// The subtitle of the slide
        /// </summary>
        public string Subtitle { get; set; } = string.Empty;

        /// <summary>
        /// The main content/synopsis/description for the slide
        /// </summary>
        public string Synopsis { get; set; } = string.Empty;

        /// <summary>
        /// Alternative property name for Synopsis (maps to Description in input)
        /// </summary>
        public string Description
        {
            get => Synopsis;
            set => Synopsis = value;
        }

        /// <summary>
        /// Collection of images to be included in the slide
        /// </summary>
        public List<ImageContent> Images { get; set; } = new List<ImageContent>();

        /// <summary>
        /// Primary background image for the slide (maps to "Suggested Image Background")
        /// </summary>
        public ImageContent? BackgroundImage { get; set; }

        /// <summary>
        /// Additional bullet points or text content
        /// </summary>
        public List<string> BulletPoints { get; set; } = new List<string>();

        /// <summary>
        /// Layout type for the slide
        /// </summary>
        public SlideLayoutType LayoutType { get; set; } = SlideLayoutType.TitleAndContent;
    }

    /// <summary>
    /// Represents an image to be included in a slide
    /// </summary>
    public class ImageContent
    {
        /// <summary>
        /// Title of the image file
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// Subtitle of the image file
        /// </summary>
        public string Subtitle { get; set; } = string.Empty;

        /// <summary>
        /// Path to the image file
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Alternative text for the image
        /// </summary>
        public string AltText { get; set; } = string.Empty;

        /// <summary>
        /// Caption for the image
        /// </summary>
        public string Caption { get; set; } = string.Empty;

        /// <summary>
        /// Position and size information for the image
        /// </summary>
        public ImagePlacement Placement { get; set; } = new ImagePlacement();
    }

    /// <summary>
    /// Defines the placement and size of an image on a slide
    /// </summary>
    public class ImagePlacement
    {
        /// <summary>
        /// X coordinate (left position) in EMUs (English Metric Units)
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Y coordinate (top position) in EMUs
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Width in EMUs
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Height in EMUs
        /// </summary>
        public long Height { get; set; }
    }

    /// <summary>
    /// Defines different slide layout types
    /// </summary>
    public enum SlideLayoutType
    {
        /// <summary>
        /// Title slide layout
        /// </summary>
        Title,

        /// <summary>
        /// Title and content layout
        /// </summary>
        TitleAndContent,

        /// <summary>
        /// Image-focused layout with minimal text
        /// </summary>
        ImageFocused,

        /// <summary>
        /// Multiple images in a grid layout
        /// </summary>
        ImageGrid,

        /// <summary>
        /// Single large image with caption
        /// </summary>
        SingleImageWithCaption,

        /// <summary>
        /// Comparison layout with two images side by side
        /// </summary>
        TwoImageComparison,

        /// <summary>
        /// Product showcase layout with title/description on left and large image on right
        /// </summary>
        ProductShowcase
    }
}
