using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length != 1)
{
    Console.Error.WriteLine("Usage: OpenXmlValidator <presentation.pptx>");
    return 2;
}

var path = Path.GetFullPath(args[0]);
if (!File.Exists(path))
{
    Console.Error.WriteLine($"PPTX file does not exist: {path}");
    return 2;
}

try
{
    using var presentation = PresentationDocument.Open(path, false);
    var errors = new OpenXmlValidator().Validate(presentation).ToList();

    foreach (var error in errors)
    {
        var part = error.Part?.Uri?.ToString() ?? "(unknown part)";
        var location = error.Path?.XPath ?? "(unknown path)";
        Console.Error.WriteLine($"{part} {location}: {error.Description}");
    }

    if (errors.Count > 0)
    {
        Console.Error.WriteLine($"Open XML validation failed: {errors.Count} error(s)");
        return 1;
    }

    Console.WriteLine($"Open XML validation passed: {path}");
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"Could not validate '{path}': {exception.Message}");
    return 2;
}
