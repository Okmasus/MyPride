using Xceed.Document.NET;

namespace CvParser.Document.Implementation.ClientCvFormats;

internal sealed class ParagraphDecorator : ICloneable
{
    private readonly Container _container;
    private readonly Paragraph? _preview;

    public ParagraphDecorator(Paragraph value, Container container, Paragraph? preview = null)
    {
        _container = container;
        _preview = preview;
        Value = value;
    }

    public Paragraph Value { get; }

    public object Clone()
    {
        var clone = _container.InsertParagraph();
        clone = _preview?.InsertParagraphAfterSelf(clone) ?? clone;

        // Копируем стиль, если есть
        clone.StyleId = Value.StyleId;

        // Копируем текст с его фрагментами (включая форматирование)
        foreach (var textRun in Value.MagicText)
            clone.Append(textRun.text, textRun.formatting);

        // Копируем абзацные свойства
        clone.Alignment = Value.Alignment;
        clone.IndentationBefore = Value.IndentationBefore;
        clone.IndentationAfter = Value.IndentationAfter;
        clone.LineSpacing = Value.LineSpacing;
        clone.LineSpacingAfter = Value.LineSpacingAfter;
        clone.LineSpacingBefore = Value.LineSpacingBefore;
        clone.Direction = Value.Direction;

        return clone;
    }
}