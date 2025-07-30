namespace CvParser.Document.Implementation.ClientCvFormats;

public abstract class PartnerFormatService
{
    protected static bool IsValidValue(string? value)
        => !string.IsNullOrEmpty(value) || !string.IsNullOrWhiteSpace(value);
}