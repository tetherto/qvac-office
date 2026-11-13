using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var dir = args.Length > 0 ? args[0] : ".";
var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var failed = 0;
var checkedFiles = 0;

foreach (var path in Directory.GetFiles(dir).OrderBy(p => p))
{
    var ext = Path.GetExtension(path).ToLowerInvariant();
    if (ext is not (".pptx" or ".docx" or ".xlsx") || path.Contains(".orig.")) continue;
    checkedFiles++;
    var errors = Validate(path, ext);
    var original = Path.Combine(dir, Path.GetFileNameWithoutExtension(path) + ".orig" + ext);
    if (File.Exists(original)) errors.ExceptWith(Validate(original, ext));
    if (errors.Count == 0) continue;
    failed++;
    Console.WriteLine($"{Path.GetFileName(path)}: {errors.Count} new error(s)");
    foreach (var error in errors.Take(10)) Console.WriteLine($"  {error}");
}

Console.WriteLine($"validated {checkedFiles} file(s), {failed} with new errors");
return failed == 0 ? 0 : 1;

HashSet<string> Validate(string path, string ext)
{
    try
    {
        using OpenXmlPackage doc = ext switch
        {
            ".pptx" => PresentationDocument.Open(path, false),
            ".docx" => WordprocessingDocument.Open(path, false),
            _ => SpreadsheetDocument.Open(path, false)
        };
        return validator
            .Validate(doc)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}")
            .ToHashSet();
    }
    catch (Exception e)
    {
        return new HashSet<string> { $"cannot open: {e.GetType().Name}: {e.Message}" };
    }
}
