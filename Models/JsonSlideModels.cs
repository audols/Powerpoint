using System.Text.Json.Serialization;

namespace PowerPointGenerator.Models
{
    /// <summary>
    /// Root JSON model for slide content
    /// </summary>
    public class JsonSlideContent
    {
        [JsonPropertyName("slides")]
        public List<JsonSlide> Slides { get; set; } = new List<JsonSlide>();
    }

    /// <summary>
    /// Individual slide in JSON format
    /// </summary>
    public class JsonSlide
    {
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        [JsonPropertyName("subtitle")]
        public string Subtitle { get; set; } = string.Empty;

        [JsonPropertyName("description")]
        public string Description { get; set; } = string.Empty;

        [JsonPropertyName("suggested_image")]
        public string SuggestedImage { get; set; } = string.Empty;

        [JsonPropertyName("layout")]
        public string Layout { get; set; } = string.Empty;

        [JsonPropertyName("images")]
        public List<JsonImage> Images { get; set; } = new List<JsonImage>();
    }

    /// <summary>
    /// Represents an image to be included in a slide
    /// </summary>
    public class JsonImage
    {
        /// <summary>
        /// Path to the image file
        /// </summary>
        [JsonPropertyName("image_file")]
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Title of the image file
        /// </summary>
        [JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// Subtitle of the image file
        /// </summary>
        [JsonPropertyName("subtitle")]
        public string Subtitle { get; set; } = string.Empty;
    }
}
