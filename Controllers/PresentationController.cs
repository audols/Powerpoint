using Microsoft.AspNetCore.Mvc;
using PowerPointGenerator.Models;
using PowerPointGenerator.Services;
using System.Text.Json;

namespace PowerPointGenerator.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PresentationController : ControllerBase
    {
        private readonly ILogger<PresentationController> _logger;
        private readonly string _outputDirectory;
        private readonly string _imageDirectory;
        private readonly string _templatesDirectory;

        public PresentationController(ILogger<PresentationController> logger)
        {
            _logger = logger;
            _outputDirectory = Path.Combine(Environment.CurrentDirectory, "GeneratedPresentations");
            _imageDirectory = Path.Combine(Environment.CurrentDirectory, "Images");
            _templatesDirectory = Path.Combine(Environment.CurrentDirectory, "Templates");
            
            // Ensure directories exist
            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }
            
            if (!Directory.Exists(_imageDirectory))
            {
                Directory.CreateDirectory(_imageDirectory);
            }

            if (!Directory.Exists(_templatesDirectory))
            {
                Directory.CreateDirectory(_templatesDirectory);
            }
        }

        /// <summary>
        /// Creates a PowerPoint presentation from JSON slide content
        /// </summary>
        /// <param name="request">The presentation creation request</param>
        /// <returns>Information about the created presentation</returns>
        [HttpPost("create-from-json")]
        public async Task<ActionResult<PresentationResponse>> CreateFromJson([FromBody] CreatePresentationRequest request)
        {
            try
            {
                _logger.LogInformation("Creating presentation from JSON content");

                // Validate request
                if (request?.JsonContent == null)
                {
                    return BadRequest(new { error = "JSON content is required" });
                }

                // Generate unique filename
                var presentationName = string.IsNullOrWhiteSpace(request.PresentationName) 
                    ? $"Presentation_{DateTime.Now:yyyyMMdd_HHmmss}" 
                    : request.PresentationName;

                var fileName = $"{presentationName}_{Guid.NewGuid():N}.pptx";
                var outputPath = Path.Combine(_outputDirectory, fileName);

                // Ensure Images directory exists
                var imageDirectory = Path.Combine(Environment.CurrentDirectory, "Images");
                if (!Directory.Exists(imageDirectory))
                {
                    Directory.CreateDirectory(imageDirectory);
                }

                // Parse JSON content
                var presentationContent = JsonSlideParser.ParseFromString(
                    request.JsonContent,
                    request.PresentationTitle ?? presentationName,
                    request.Author ?? "API User",
                    imageDirectory
                );

                // Create the presentation
                using var generator = new PowerPointGeneratorService();
                await generator.CreatePresentationAsync(presentationContent, outputPath);

                // Return response with file information
                var response = new PresentationResponse
                {
                    Success = true,
                    FileName = fileName,
                    FilePath = outputPath,
                    PresentationName = presentationName,
                    CreatedAt = DateTime.UtcNow,
                    FileSize = new FileInfo(outputPath).Length,
                    SlideCount = presentationContent.Slides.Count,
                    DownloadUrl = $"/api/presentation/download/{fileName}"
                };

                _logger.LogInformation("Successfully created presentation: {FileName}", fileName);
                return Ok(response);
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Invalid JSON format in request");
                return BadRequest(new { error = "Invalid JSON format", details = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating presentation");
                return StatusCode(500, new { error = "Failed to create presentation", details = ex.Message });
            }
        }

        /// <summary>
        /// Creates a PowerPoint presentation from a template and JSON content
        /// </summary>
        /// <param name="request">The presentation creation request with optional template name</param>
        /// <returns>Information about the created presentation</returns>
        [HttpPost("create-from-template")]
        public async Task<ActionResult<PresentationResponse>> CreateFromTemplate([FromBody] CreatePresentationFromTemplateRequest request)
        {
            try
            {
                _logger.LogInformation("Creating presentation from template");

                // Validate request
                if (request?.JsonContent == null)
                {
                    return BadRequest(new { error = "JSON content is required" });
                }

                // Determine template file to use
                var templateName = string.IsNullOrWhiteSpace(request.TemplateName) 
                    ? "test_template.pptx" 
                    : request.TemplateName;

                // Ensure template has .pptx extension
                if (!templateName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    templateName += ".pptx";
                }

                var templatePath = Path.Combine(_templatesDirectory, templateName);

                // Check if template exists
                if (!System.IO.File.Exists(templatePath))
                {
                    return BadRequest(new { error = $"Template file '{templateName}' not found in Templates folder. Available templates: {string.Join(", ", GetAvailableTemplates())}" });
                }

                // Generate unique filename
                var finalPresentationName = string.IsNullOrWhiteSpace(request.PresentationName)
                    ? $"Presentation_{DateTime.Now:yyyyMMdd_HHmmss}"
                    : request.PresentationName;

                var fileName = $"{finalPresentationName}_{Guid.NewGuid():N}.pptx";
                var outputPath = Path.Combine(_outputDirectory, fileName);

                // Parse JSON content
                var presentationContent = JsonSlideParser.ParseFromString(
                    request.JsonContent,
                    request.PresentationTitle ?? finalPresentationName,
                    request.Author ?? "API User",
                    _imageDirectory
                );

                // Create the presentation from template
                using var generator = new PowerPointGeneratorService();
                await generator.CreatePresentationFromTemplateAsync(presentationContent, templatePath, outputPath);

                // Return response with file information
                var response = new PresentationResponse
                {
                    Success = true,
                    FileName = fileName,
                    FilePath = outputPath,
                    PresentationName = finalPresentationName,
                    CreatedAt = DateTime.UtcNow,
                    FileSize = new FileInfo(outputPath).Length,
                    SlideCount = presentationContent.Slides.Count,
                    DownloadUrl = $"/api/presentation/download/{fileName}"
                };

                _logger.LogInformation("Successfully created presentation from template '{TemplateName}': {FileName}", templateName, fileName);
                return Ok(response);
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Invalid JSON format in request");
                return BadRequest(new { error = "Invalid JSON format", details = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating presentation from template");
                return StatusCode(500, new { error = "Failed to create presentation from template", details = ex.Message });
            }
        }

        /// <summary>
        /// Downloads a generated presentation file
        /// </summary>
        /// <param name="fileName">The name of the file to download</param>
        /// <returns>The presentation file</returns>
        [HttpGet("download/{fileName}")]
        public async Task<IActionResult> DownloadPresentation(string fileName)
        {
            try
            {
                var filePath = Path.Combine(_outputDirectory, fileName);
                
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                var contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                
                return File(fileBytes, contentType, fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading presentation: {FileName}", fileName);
                return StatusCode(500, new { error = "Failed to download presentation" });
            }
        }

        /// <summary>
        /// Gets a list of all generated presentations
        /// </summary>
        /// <returns>List of generated presentations</returns>
        [HttpGet("list")]
        public ActionResult<IEnumerable<PresentationFileInfo>> ListPresentations()
        {
            try
            {
                var files = Directory.GetFiles(_outputDirectory, "*.pptx")
                    .Select(filePath => new PresentationFileInfo
                    {
                        FileName = Path.GetFileName(filePath),
                        CreatedAt = System.IO.File.GetCreationTimeUtc(filePath),
                        FileSize = new FileInfo(filePath).Length,
                        DownloadUrl = $"/api/presentation/download/{Path.GetFileName(filePath)}"
                    })
                    .OrderByDescending(f => f.CreatedAt)
                    .ToList();

                return Ok(files);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error listing presentations");
                return StatusCode(500, new { error = "Failed to list presentations" });
            }
        }

        /// <summary>
        /// Deletes a generated presentation file
        /// </summary>
        /// <param name="fileName">The name of the file to delete</param>
        /// <returns>Deletion result</returns>
        [HttpDelete("delete/{fileName}")]
        public IActionResult DeletePresentation(string fileName)
        {
            try
            {
                var filePath = Path.Combine(_outputDirectory, fileName);
                
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                System.IO.File.Delete(filePath);
                _logger.LogInformation("Deleted presentation: {FileName}", fileName);
                
                return Ok(new { success = true, message = "Presentation deleted successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting presentation: {FileName}", fileName);
                return StatusCode(500, new { error = "Failed to delete presentation" });
            }
        }

        /// <summary>
        /// Uploads an image file for use in presentations
        /// </summary>
        /// <param name="file">The image file to upload</param>
        /// <returns>Information about the uploaded image</returns>
        [HttpPost("upload-image")]
        [Consumes("multipart/form-data")]
        public async Task<ActionResult<ImageUploadResponse>> UploadImage([FromForm] IFormFile file)
        {
            try
            {
                // Validate file
                if (file == null || file.Length == 0)
                {
                    return BadRequest(new ImageUploadResponse
                    {
                        Success = false,
                        ErrorMessage = "No file was uploaded"
                    });
                }

                // Validate file type
                var allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
                var fileExtension = Path.GetExtension(file.FileName);
                
                if (!allowedExtensions.Contains(fileExtension))
                {
                    return BadRequest(new ImageUploadResponse
                    {
                        Success = false,
                        ErrorMessage = $"File type '{fileExtension}' is not supported. Allowed types: {string.Join(", ", allowedExtensions)}"
                    });
                }

                // Validate file size (max 10MB)
                const long maxFileSize = 10 * 1024 * 1024; // 10MB
                if (file.Length > maxFileSize)
                {
                    return BadRequest(new ImageUploadResponse
                    {
                        Success = false,
                        ErrorMessage = $"File size ({file.Length / 1024 / 1024}MB) exceeds maximum allowed size (10MB)"
                    });
                }

                // Use exact filename from request
                var fileName = file.FileName;
                var filePath = Path.Combine(_imageDirectory, fileName);

                // Check if file already exists
                if (System.IO.File.Exists(filePath))
                {
                    _logger.LogInformation("Image already exists, skipping upload: {FileName}", fileName);
                    
                    return Ok(new ImageUploadResponse
                    {
                        Success = true,
                        FileName = fileName,
                        FilePath = filePath,
                        UploadedAt = System.IO.File.GetCreationTimeUtc(filePath),
                        FileSize = new FileInfo(filePath).Length,
                        ImageUrl = $"/api/presentation/image/{fileName}",
                        ErrorMessage = "File already exists, upload skipped"
                    });
                }

                // Save the file
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                _logger.LogInformation("Successfully uploaded image: {FileName}", fileName);

                return Ok(new ImageUploadResponse
                {
                    Success = true,
                    FileName = fileName,
                    FilePath = filePath,
                    UploadedAt = DateTime.UtcNow,
                    FileSize = file.Length,
                    ImageUrl = $"/api/presentation/image/{fileName}"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading image");
                return StatusCode(500, new ImageUploadResponse
                {
                    Success = false,
                    ErrorMessage = "Failed to upload image"
                });
            }
        }

        /// <summary>
        /// Uploads multiple image files for use in presentations
        /// </summary>
        /// <param name="files">The image files to upload</param>
        /// <returns>Information about the uploaded images</returns>
        [HttpPost("upload-images")]
        [Consumes("multipart/form-data")]
        public async Task<ActionResult<IEnumerable<ImageUploadResponse>>> UploadImages([FromForm] List<IFormFile> files)
        {
            var results = new List<ImageUploadResponse>();

            if (files == null || files.Count == 0)
            {
                return BadRequest(new { error = "No files were uploaded" });
            }

            foreach (var file in files)
            {
                try
                {
                    var singleFileResult = await UploadImage(file);
                    if (singleFileResult.Result is OkObjectResult okResult && okResult.Value is ImageUploadResponse response)
                    {
                        results.Add(response);
                    }
                    else if (singleFileResult.Result is BadRequestObjectResult badResult && badResult.Value is ImageUploadResponse errorResponse)
                    {
                        results.Add(errorResponse);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error uploading file: {FileName}", file.FileName);
                    results.Add(new ImageUploadResponse
                    {
                        Success = false,
                        FileName = file.FileName,
                        ErrorMessage = $"Failed to upload {file.FileName}: {ex.Message}"
                    });
                }
            }

            return Ok(results);
        }

        /// <summary>
        /// Gets an uploaded image file
        /// </summary>
        /// <param name="fileName">The name of the image file</param>
        /// <returns>The image file</returns>
        [HttpGet("image/{fileName}")]
        public async Task<IActionResult> GetImage(string fileName)
        {
            try
            {
                var filePath = Path.Combine(_imageDirectory, fileName);
                
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "Image not found" });
                }

                var fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                var contentType = GetContentType(Path.GetExtension(fileName));
                
                return File(fileBytes, contentType, fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving image: {FileName}", fileName);
                return StatusCode(500, new { error = "Failed to retrieve image" });
            }
        }

        /// <summary>
        /// Gets a list of all uploaded images
        /// </summary>
        /// <returns>List of uploaded images</returns>
        [HttpGet("images")]
        public ActionResult<IEnumerable<ImageInfo>> ListImages()
        {
            try
            {
                var imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
                var imageFiles = Directory.GetFiles(_imageDirectory)
                    .Where(f => imageExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                    .Select(filePath =>
                    {
                        var fileName = Path.GetFileName(filePath);
                        var fileInfo = new FileInfo(filePath);
                        
                        return new ImageInfo
                        {
                            FileName = fileName,
                            UploadedAt = fileInfo.CreationTimeUtc,
                            FileSize = fileInfo.Length,
                            ImageUrl = $"/api/presentation/image/{fileName}",
                            Dimensions = GetImageDimensions(filePath)
                        };
                    })
                    .OrderByDescending(i => i.UploadedAt)
                    .ToList();

                return Ok(imageFiles);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error listing images");
                return StatusCode(500, new { error = "Failed to list images" });
            }
        }

        /// <summary>
        /// Deletes an uploaded image file
        /// </summary>
        /// <param name="fileName">The name of the image file to delete</param>
        /// <returns>Deletion result</returns>
        [HttpDelete("image/{fileName}")]
        public IActionResult DeleteImage(string fileName)
        {
            try
            {
                var filePath = Path.Combine(_imageDirectory, fileName);
                
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "Image not found" });
                }

                System.IO.File.Delete(filePath);
                _logger.LogInformation("Deleted image: {FileName}", fileName);
                
                return Ok(new { success = true, message = "Image deleted successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting image: {FileName}", fileName);
                return StatusCode(500, new { error = "Failed to delete image" });
            }
        }

        /// <summary>
        /// Gets a list of available template files
        /// </summary>
        /// <returns>List of available template files</returns>
        [HttpGet("templates")]
        public ActionResult<IEnumerable<string>> ListTemplates()
        {
            try
            {
                var templates = GetAvailableTemplates();
                return Ok(templates);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error listing templates");
                return StatusCode(500, new { error = "Failed to list templates" });
            }
        }

        /// <summary>
        /// Health check endpoint
        /// </summary>
        /// <returns>Service health status</returns>
        [HttpGet("health")]
        public IActionResult HealthCheck()
        {
            return Ok(new 
            { 
                status = "healthy", 
                timestamp = DateTime.UtcNow,
                version = "1.0.0",
                service = "PowerPoint Generator API"
            });
        }

        /// <summary>
        /// Gets the MIME content type for an image file extension
        /// </summary>
        /// <param name="extension">File extension</param>
        /// <returns>MIME content type</returns>
        private static string GetContentType(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                ".webp" => "image/webp",
                _ => "application/octet-stream"
            };
        }

        /// <summary>
        /// Gets image dimensions as a string
        /// </summary>
        /// <param name="imagePath">Path to the image file</param>
        /// <returns>Dimensions string (width x height) or null if cannot be determined</returns>
        private string? GetImageDimensions(string imagePath)
        {
            try
            {
                using var image = System.Drawing.Image.FromFile(imagePath);
                return $"{image.Width} x {image.Height}";
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Gets a list of available template files from the Templates directory
        /// </summary>
        /// <returns>List of template file names</returns>
        private List<string> GetAvailableTemplates()
        {
            try
            {
                if (!Directory.Exists(_templatesDirectory))
                {
                    return new List<string>();
                }

                return Directory.GetFiles(_templatesDirectory, "*.pptx")
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .ToList()!;
            }
            catch
            {
                return new List<string>();
            }
        }
    }
}
