using CvParser.Document.DTO;
using CvParser.Document.Interfaces.ClientCvFormats;
using CvParser.Extensions;
using System.Drawing;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Font = Xceed.Document.NET.Font;

namespace CvParser.Document.Implementation.ClientCvFormats;

public class GoDigitalFormatService : TemplateDocumentFormatService
{
    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        using var ms = new MemoryStream();
        using var doc = DocX.Create(ms);

        doc.SetDefaultFont(new Font("Manrope"), fontSize: 12, Color.FromArgb(60, 70, 83));
        doc.AddHeaders();
        doc.AddFooters();
        doc.DifferentFirstPage = true;
        doc.MarginTop = 10f;
        doc.MarginBottom = 10f;
        doc.MarginLeft = 60f;
        doc.MarginRight = 60f;

        var paragraph = doc.InsertParagraph();
        InsertImageAndTextRow(doc, paragraph, cvDocumentDTO);
        doc.InsertParagraph().SpacingAfter(25);

        if(IsValidValue(cvDocumentDTO.Skills))
        {
            InsertData("Компетенции:", doc, new Formatting() { Bold = true, Size = 16}, spacingAfter: 15);
            InsertData($"{cvDocumentDTO.Skills}", doc, new Formatting() { Size = 12 }, spacingAfter: 15, lineSpacing: 20);
        }

        InsertData($"Опыт", doc, new Formatting()
        {
            Size = 16,
            Bold = true,
        }, spacingAfter: 15);

        InsertExperiences(doc, cvDocumentDTO.Works);

        var actualEducations = cvDocumentDTO.Educations
            .Where(education =>
                   IsValidValue(education.Description?.Value)
                || IsValidValue(education.Title));

        if (actualEducations.Count() > 0)
        {
            InsertData($"Образование", doc, new Formatting()
            {
                Size = 16,
                Bold = true,
            }, spacingAfter: 15);
            foreach (var education in actualEducations)
            {
                InsertData($"{education.Title}/ {education.Description?.Value}/ {education.Year}", doc, new Formatting()
                {
                    Size = 12,
                }, spacingAfter: 15);
            }
        }

        if (IsValidValue(cvDocumentDTO.Languages))
        {
            InsertData("Иностранный язык", doc, new Formatting() { Bold = true, Size = 16 }, spacingAfter: 15);
            InsertData($"{cvDocumentDTO.Languages}", doc, new Formatting() { Size = 12 }, spacingAfter: 15, lineSpacing: 20);
        }

        doc.SaveAs(ms);
        return ms;
    }

    private void InsertExperiences(DocX doc, WorkDTO[] workExperiences)
    {
        foreach (var experience in workExperiences)
        {
            var table = doc.AddTable(5, 2);
            table.Alignment = Alignment.left;
            table.Design = TableDesign.TableGrid;
            table.AutoFit = AutoFit.Window;
            table.SetWidths(new float[] { 40f, 700f });

            foreach (var row in table.Rows)
                foreach (var cell in row.Cells)
                {
                    cell.MarginTop = 5;
                    cell.MarginBottom = 5;
                    cell.VerticalAlignment = VerticalAlignment.Top;
                }

            // Row 0: Должность в компании
            table.Rows[0].MergeCells(0, 1);

            table.Rows[0].Cells[0].Paragraphs[0]
                .Append(IsValidValue(experience.Position) ? $"{experience.Position} в " : "").Bold().FontSize(13)
                .Append(experience.Title).Bold().UnderlineStyle(UnderlineStyle.singleLine).Alignment = Alignment.center;
            table.Rows[0].Cells[0].Paragraphs[0].AppendLine($"{experience.Start.ToString("MMMM yyyy", Extension.Ru)} - {(experience.End.HasValue ? experience.End.Value.ToString("MMMM yyyy", Extension.Ru) : "настоящее время")}").FontSize(11);

            // Row 1: Проект
            if (IsValidValue(experience.Description.Value))
            {
                table.Rows[1].Cells[1].Paragraphs[0]
                    .Append("Проект: ").Bold()
                    .Append(experience.Description.Value).SpacingAfter(5);
            }
            else
                table.RemoveRow();

            // Row 2: Задачи / Результаты
            var cell2 = table.Rows[2].Cells[1];

            if(!IsValidValue(experience.Tasks.Value) || !IsValidValue(experience.Results.Value))
                table.RemoveRow();
            else
            {
                if (IsValidValue(experience.Tasks.Value))
                {
                    cell2.Paragraphs[0]
                        .Append("Задачи:").Bold().SpacingAfter(5)
                        .AppendLine()
                        .Append(experience.Tasks.Value).SpacingAfter(5).AppendLine();
                }
                if (IsValidValue(experience.Results.Value))
                {
                    cell2.Paragraphs[0]
                        .Append("Результаты:").Bold().SpacingAfter(5)
                        .AppendLine()
                        .Append(experience.Results.Value).SpacingAfter(5).AppendLine();
                }
            }

            // Row 3: Team
            if (IsValidValue(experience.Team.Value))
            {
                table.Rows[3].Cells[1].Paragraphs[0]
                    .Append("Команда: ").Bold()
                    .Append(experience.Team.Value).SpacingAfter(5);
            }
            else
                table.RemoveRow();

            // Row 4: Tools
            if (IsValidValue(experience.Tools.Value))
            {
                table.Rows[table.RowCount - 1].Cells[table.ColumnCount - 1].InsertParagraph();
                table.Rows[table.RowCount - 1].Cells[table.ColumnCount - 1].Paragraphs[0].Append("Стек: ").Bold().Append(experience.Tools.Value).SpacingAfter(5);
            }
            else
                table.RemoveRow();

            doc.InsertTable(table);
            doc.InsertParagraph().SpacingAfter(25);
        }
    }

    private static bool IsValidValue(string? value)
        => !string.IsNullOrEmpty(value) || !string.IsNullOrWhiteSpace(value);

    private static Table InsertTable(DocX doc, in int columnsCount)
        => doc.InsertTable(1, columnsCount);

    private void InsertImageAndTextRow(DocX doc, Paragraph paragraph, CvDTO cv)
    {
        var table = InsertTable(doc, 2);
        table.SetWidths(new float[] { 500f, 250f });

        table.SetBorder(TableBorderType.Top, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.SetBorder(TableBorderType.Bottom, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.SetBorder(TableBorderType.Left, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.SetBorder(TableBorderType.Right, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.SetBorder(TableBorderType.InsideH, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.SetBorder(TableBorderType.InsideV, new Border(BorderStyle.Tcbs_none, 0, 0, Color.White));
        table.Alignment = Alignment.both;

        var imageCell = table.Rows[0].Cells[1];
        using var ms = new MemoryStream(Properties.Resources.GoDigitalLogo);
        var logo = doc.AddImage(ms);
        imageCell.Paragraphs[0].AppendPicture(logo.CreatePicture(145f, 160f))
               .Alignment = Alignment.right;

        var textCell = table.Rows[0].Cells[0];
        textCell.Paragraphs[0].SpacingLine(20);


        if (cv.Staff is not null || IsValidValue(cv.Staff?.Stak))
        {
            textCell.Paragraphs[0]
               .Append($"{cv.Staff.Stak}")
               .Font(new Font("Manrope")).Bold().FontSize(16).SpacingAfter(30)
               .AppendLine();
        }

        var nameBlock = $"{cv.LastName} {cv.FirstName} {cv.SurName}";

        if (IsValidValue(nameBlock))
        {
            textCell.Paragraphs[0]
               .Append($"{cv.LastName} {cv.FirstName} {cv.SurName}")
               .Font(new Font("Manrope")).Bold().FontSize(16).SpacingAfter(30)
               .AppendLine();
        }

        if (IsValidValue(cv.AllExperience))
        {
            textCell.Paragraphs[0]
               .Append($"Профессиональный опыт: {cv.AllExperience}").SpacingAfter(20)
               .Font(new Font("Manrope")).FontSize(12)
               .AppendLine();
        }

        if (IsValidValue(cv.Location))
        {
            textCell.Paragraphs[0]
               .Append($"Локация: {cv.Location}").Font(new Font("Manrope")).FontSize(12).SpacingAfter(20)
               .AppendLine();
            
        }

        if (cv.Staff is not null || IsValidValue(cv.Staff?.Grade))
        {
            textCell.Paragraphs[0]
                .Append($"Грейд: {cv.Staff?.Grade}").Font(new Font("Manrope")).FontSize(12).SpacingAfter(20);
        }
    }
    private void InsertData(
        in string? paragraphTextContent,
        Container container,
        Formatting? formatting = null,
        string? headingStyleId = null,
        Alignment alignment = Alignment.left,
        float spacingAfter = 0,
        float lineSpacing = 0)
    {
        if (string.IsNullOrEmpty(paragraphTextContent) || string.IsNullOrWhiteSpace(paragraphTextContent))
            return;

        var paragraph = container.InsertParagraph();
        paragraph.InsertText(paragraphTextContent, formatting: formatting);

        if (!string.IsNullOrEmpty(headingStyleId))
            paragraph.StyleId = headingStyleId;

        paragraph.Alignment = alignment;

        if (spacingAfter > 0)
            paragraph.SpacingAfter(spacingAfter);

        if (lineSpacing > 0)
            paragraph.LineSpacing = lineSpacing;
    }
}
