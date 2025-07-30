using CvParser.Document.DTO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace CvParser.Document.Implementation.ClientCvFormats;

public abstract class TemplateDocumentFormatService
{
    /// <summary>
    /// Создать документ.
    /// </summary>
    /// <param name="cvDocumentDTO">Резюме.</param>
    /// <param name="template">Шаблон.</param>
    /// <returns>Документ.</returns>
    public abstract CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO);
    
    public virtual void WriteCollectionBlocks(
        Container container,
        CvDTO cvDocumentDto, 
        string positionInTemplate,
        Action<Paragraph[], Container, Paragraph, CvDTO> collectionBlocksWriter)
    {
        var paragraphs = container
            .Paragraphs
            .SkipWhile(p => p.Text != $"[{positionInTemplate}]")
            .Skip(1)
            .TakeWhile(p => p.Text != $"[{positionInTemplate}]")
            .ToArray();
        var endOfBlockParagraph = paragraphs[^1].NextParagraph;
        
        collectionBlocksWriter.Invoke(paragraphs.ToArray(), container, endOfBlockParagraph, cvDocumentDto);

        RemoveTemplateParagraphs(
            container,
            container
                .Paragraphs
                .Where(p => p.Text == $"[{positionInTemplate}]")
                .Concat(paragraphs)
                .ToList());
    }

    public virtual void WriteData(DocX doc, string positionInTemplate, string text)
    {
        var stringReplaceTextOptions = new StringReplaceTextOptions
        {
            SearchValue = $"{{{positionInTemplate}}}",
            NewValue = text
        };

        doc.ReplaceText(stringReplaceTextOptions);
    }

    public virtual void WriteTable(
        DocX doc,
        CvDTO cvDocumentDto, 
        string positionInTemplate,
        Action<Table, CvDTO> tableWriter)
    {
        var paragraphWithTable = doc
            .Paragraphs
            .SkipWhile(p => p.Text != $"[{positionInTemplate}]")
            .Skip(1)
            .FirstOrDefault();
        var table = paragraphWithTable?.FollowingTables.FirstOrDefault();

        if (table is null)
            return;

        WriteData(doc, positionInTemplate, string.Empty);
        
        tableWriter.Invoke(table, cvDocumentDto);
        
        RemoveTemplateParagraphs(
            doc,
            doc.Paragraphs.Where(p => p.Text == $"[{positionInTemplate}]").ToList());
        
        table.RemoveRow(0);
    }
    
    protected static Paragraph CopyParagraph(Paragraph paragraph, Container container, Paragraph? preview = null)
    {
        var copy = (Paragraph) new ParagraphDecorator(paragraph, container, preview).Clone();    
        
        return copy;
    }

    private static void RemoveTemplateParagraphs(Container container, List<Paragraph> paragraphs)
    {
        while (paragraphs.Count > 0)
        {
            var lastParagraph = paragraphs[^1];
            
            container.RemoveParagraph(lastParagraph);
            paragraphs.Remove(lastParagraph);
        }
    }
}