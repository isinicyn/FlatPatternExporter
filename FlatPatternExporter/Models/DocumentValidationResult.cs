using FlatPatternExporter.Enums;
using Inventor;

namespace FlatPatternExporter.Models;

public class DocumentValidationResult
{
    public Document? Document { get; set; }
    public DocumentType DocType { get; set; }
    public bool IsValid { get; set; }
    public string ErrorMessage { get; set; } = "";
    public string DocumentTypeName { get; set; } = "";
}