using System.Globalization;
using CvParser.Document.DTO;
using CvParser.Document.Extensions;
using CvParser.Document.Interfaces.ClientCvFormats;
using CvParser.Services.DTO;
using Microsoft.Extensions.Logging;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace CvParser.Document.Implementation.ClientCvFormats;

public sealed class MosbirjaFormatService : TemplateDocumentFormatService
{
    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        var ms = new MemoryStream();
        
        using var doc = DocX.Load("Templates/Mosbirja.docx");

        WriteData(doc, "FullName", $"{cvDocumentDTO.LastName} {cvDocumentDTO.FirstName} {cvDocumentDTO.SurName}");
        WriteData(doc, "Position", cvDocumentDTO.Stack);
        WriteData(doc, "Grade", cvDocumentDTO.Staff!.Grade!);
        WriteData(doc, "WorkStart", DateTimeOffset.UtcNow.ToString("dd MMMM yyyy", new CultureInfo("ru-RU")));
        WriteData(doc, "Location", cvDocumentDTO.Location);
        WriteData(doc, "Description", (cvDocumentDTO.About ?? TextDto.Empty).Value);
        WriteTable(doc, cvDocumentDTO, "Stacks.Table", SkillsWriter);
        WriteCollectionBlocks(doc, cvDocumentDTO, "Educations", EducationsWriter);
        WriteCollectionBlocks(doc, cvDocumentDTO, "Works", WorksWriter);

        doc.SaveAs(ms);

        return ms;
    }
    
    private static void EducationsWriter (Paragraph[] educationsParagraphs, Container container, Paragraph previous, CvDTO dto)
    {
        for (var icv = 0; icv < dto.Educations.Length; icv++)
            for (var i = educationsParagraphs.Length - 1; i >= 0; i--)
            {
                var education = dto.Educations[icv];
                var paragraph = CopyParagraph(educationsParagraphs[i], container, previous);

                paragraph.ReplaceText("{Education.Organization}", education.Organization);
                paragraph.ReplaceText("{Education.Result}", education.Result);
                paragraph.ReplaceText("{Education.Year}", education.Year.ToString());
                paragraph.ReplaceText("{Education.Name}", education.Title);
            }
    }
        
    private static void WorksWriter (Paragraph[] workParagraphs, Container container, Paragraph previous, CvDTO dto)
    {
        for (var icv = 0; icv < dto.Works.Length; icv++)
            for (var i = workParagraphs.Length - 1; i >= 0; i--)
            {
                var work = dto.Works[icv];
                var paragraph = CopyParagraph(workParagraphs[i], container, previous);
                
                paragraph.ReplaceText("{Work.Name}", work.Title);
                paragraph.ReplaceText("{Work.Start}", work.Start.ToString("MMMM yyyy"));
                paragraph.ReplaceText("{Work.End}", work.End.HasValue ? work.End!.Value.ToString("MMMM yyyy") : "по настоящее время");
                paragraph.ReplaceText("{Work.Tools}", work.Tools.Value);
                paragraph.ReplaceText("{Work.Description}", work.Tasks.Value);
                paragraph.ReplaceText("{Work.Result}", work.Results.Value);
            }
    }
        
    private static void SkillsWriter (Table parameterTable, CvDTO dto)
    {
        var templateRow = parameterTable.Rows[0];

        foreach (var categoryGroup in dto
                     .SkillsDto
                     .GroupBy(skill => skill.Category))
        {
            var row = parameterTable.InsertRow();

            var categoryParagraph = CopyParagraph(templateRow.Cells[0].Paragraphs[0], row.Cells[0]);
            var stacksParagraph = CopyParagraph(templateRow.Cells[1].Paragraphs[0], row.Cells[1]);

            categoryParagraph.ReplaceText(new StringReplaceTextOptions
            {
                SearchValue = "{Category.Name}",
                NewValue = categoryGroup.Key
            });
            stacksParagraph.ReplaceText(new StringReplaceTextOptions
            {
                SearchValue = "{Category.Tools}",
                NewValue = string.Join(", ", categoryGroup.Select(skill => skill.Value))
            });
        }
    }
}