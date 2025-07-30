using CvParser.Document.DTO;
using CvParser.Document.Interfaces.ClientCvFormats;
using CvParser.Extensions;
using System.Drawing;
using Xceed.Document.NET;
using Xceed.Words.NET;
using CvParser.Services.DTO;
using Font = Xceed.Document.NET.Font;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;

namespace CvParser.Document.Implementation.ClientCvFormats;

/// <summary>
/// Сервис форматирования резюме для компании Digimatics.
/// </summary>
public sealed class DigimaticsFormatService : TemplateDocumentFormatService, IDataInserter
{
    private const string FiraSansFontFamily = "Fira Sans";
    private const string FiraSansLightFontFamily = "Fira Sans Light";

    private static readonly Color NameFontColor = Color.FromArgb(46, 83, 149);
    private static readonly Color OrganizationFontColor = Color.FromArgb(102, 102, 102);

    private const string OrganizationPreview = @"Компания Цифроматика, ведущая проекты по разработке информационных систем, систем поддержки принятия решения, предиктивной аналитики и машинного обучения для промышленности, на транспорте и в государственном секторе.";
    private const string OrganizationAddressContent = "Санкт-Петербург, улица Гончарная 27,\r\nлитер А, офис 209\r\n";
    private const string OrganizationContactsContent = "тел. +7 (812) 408-33-44\r\ninfo@digimatics.ru\r\n";


    private const string LocationTitle = "Location";
    private const string TechnologyStackTitle = "Technology Stack";
    private const string EducationTitle = "Education";
    private const string LanguagesTitle = "Language";
    private const string LastProjectExperienceTitle = "Last Project Experience";


    private static Formatting NameFormatting => new()
    {
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        Size = 24,
        Bold = true,
        FontColor = NameFontColor,
        StyleName = "Normal"
    };
    private static Formatting TitleFormatting => new()
    {
        Size = 16,
        FontFamily = new Xceed.Document.NET.Font(FiraSansLightFontFamily),
        StyleName = "Normal"
    };
    private static Formatting SectionFormatting => new()
    {
        Size = 12,
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        StyleName = "Normal"
    };
    private static Formatting SectionLightFormatting => new()
    {
        Size = 12,
        Bold = true,
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        StyleName = "Normal"
    };
    private static Formatting SectionBoldFormatting => new()
    {
        Size = 12,
        Bold = true,
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        StyleName = "Normal"
    };
    private static Formatting SmallFormatting => new()
    {
        Size = 9,
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        FontColor = OrganizationFontColor,
        StyleName = "Normal"
    };
    private static Formatting HeaderFormatting => new()
    {
        Size = 7,
        FontFamily = new Xceed.Document.NET.Font(FiraSansFontFamily),
        StyleName = "Normal"
    };

    /// <summary>
    /// Создать документ на основе резюме.
    /// </summary>
    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        using var ms = new MemoryStream();
        using var doc = DocX.Create(ms);

        doc.AddHeaders();
        doc.AddFooters();
        doc.Footers.Even.PageNumbers = true;
        doc.Footers.Odd.PageNumbers = true;
        doc.SetDefaultFont(new Font("Calibri"));
        CustomHeader(doc, doc.Headers.Even);
        CustomHeader(doc, doc.Headers.Odd);
        InsertOrganizationPreview(doc);
        InsertEmptyParagraphs(doc, 4);

        // Имя
        InsertData(cvDocumentDTO.FirstName ?? string.Empty, doc, NameFormatting);

        // Должность и грейд
        if (cvDocumentDTO.Staff != null && (IsValidValue(cvDocumentDTO.Staff.Grade) || IsValidValue(cvDocumentDTO.Staff.Stak)))
        {
            InsertData($"{cvDocumentDTO.Staff.Grade}, {cvDocumentDTO.Staff.Stak}", doc, SectionLightFormatting);
        }

        InsertEmptyParagraphs(doc, 2);

        InsertSectionIfValid(doc, LocationTitle, cvDocumentDTO.Location);
        InsertSectionIfValid(doc, TechnologyStackTitle, cvDocumentDTO.Stack);

        // Образование
        if (cvDocumentDTO.Educations != null && cvDocumentDTO.Educations.Length > 0)
        {
            InsertTitle(EducationTitle, doc);
            InsertEducations(doc, cvDocumentDTO.Educations);
        }
        InsertEmptyParagraph(doc);

        // Опыт работы
        if (cvDocumentDTO.Works != null && cvDocumentDTO.Works.Length > 0)
        {
            InsertTitle(LastProjectExperienceTitle, doc);
            InsertWorkExperiences(doc, cvDocumentDTO.Works);
        }

        // Языки
        InsertSectionIfValid(doc, LanguagesTitle, cvDocumentDTO.Languages, isTable: true);

        doc.SaveAs(ms);
        return ms;
    }

    /// <summary>
    /// Вставляет текст в контейнер.
    /// </summary>
    public void InsertData(in string? paragraphTextContent, Container container, Formatting? formatting = null, Alignment alignment = Alignment.left)
    {
        var paragraph = container.InsertParagraph();
        paragraph.Alignment = alignment;
        ConfigurateTextStyle(paragraph);
        paragraph.InsertText(paragraphTextContent, formatting: formatting);
    }

    private void CustomHeader(DocX doc, Header header)
    {
        var headerContent = header.InsertTable(1, 4);
        HideAllBordersOfTable(headerContent);
        var headerContentRow = headerContent.InsertRow();
        InsertImage(doc, headerContentRow.Cells[0]);
        headerContentRow.Cells[1].Width = 51.04;
        InsertData(OrganizationAddressContent, headerContentRow.Cells[2], HeaderFormatting);
        InsertData(OrganizationContactsContent, headerContentRow.Cells[3], HeaderFormatting);
    }

    private static void InsertImage(DocX doc, Cell cell)
    {
        using var ms = new MemoryStream(Properties.Resources.DigimaticsLogo);
        var logo = doc.AddImage(ms);
        cell.InsertParagraph().AppendPicture(logo.CreatePicture(18, 158.7f));
    }

    private void InsertEducations(DocX doc, EducationDTO[] educations)
    {
        var table = doc.InsertTable(1, 2);
        HideAllBordersOfTable(table);
        foreach (var education in educations)
        {
            var row = table.InsertRow();
            InsertData(education.Year.ToString(), row.Cells[0], SectionFormatting);
            InsertData(education.Title, row.Cells[1], SectionBoldFormatting);
            if (education.Description != null && IsValidValue(education.Description.Value))
            {
                InsertData(education.Description.Value, row.Cells[1], SectionFormatting);
            }
        }
    }

    private void InsertLanguages(DocX doc, string languages)
    {
        var table = doc.InsertTable(1, 1);
        HideAllBordersOfTable(table);
        var row = table.InsertRow();
        InsertData(languages, row.Cells[0], SectionFormatting);
    }

    private static void HideAllBordersOfTable(Table table)
    {
        table.SetBorder(TableBorderType.Top, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.Bottom, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.Right, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.Left, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.InsideV, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.InsideH, new Border { Tcbs = BorderStyle.Tcbs_none });
    }

    private static void ConfigurateTextStyle(Paragraph paragraph)
    {
        paragraph.StyleName = "Normal";
    }
    private void InsertWorkExperiences(DocX doc, WorkDTO[] workExperiences)
    {
        foreach (var work in workExperiences.OrderByDescending(w => w.End ?? DateTime.MaxValue))
        {
            InsertData(work.Position, doc, SectionLightFormatting);
            var dateRange = $"{work.Start.ToString("MMMM yyyy", Extension.Ru)} - {(work.End.HasValue ? work.End.Value.ToString("MMMM yyyy", Extension.Ru) : "настоящее время")}";
            InsertData($"Проект ({dateRange}):", doc, SectionLightFormatting);
            InsertWorkSection(doc, work.Description, null);
            InsertWorkSection(doc, work.Tasks, "Задачи:");
            InsertWorkSection(doc, work.Results, "Результат работы:");
            InsertWorkSection(doc, work.Tools, "Стек:");
            InsertEmptyParagraphs(doc, 2);
        }
    }

    private void InsertWorkSection(DocX doc, TextDto textDto, string? header)
    {
        if (textDto != null && IsValidValue(textDto.Value))
        {
            if (!string.IsNullOrEmpty(header))
            {
                InsertEmptyParagraph(doc);
                InsertData(header, doc, SectionBoldFormatting);
            }
            InsertBulletList(doc, textDto.Value);
        }
    }
    private void InsertSectionIfValid(DocX doc, string title, string? value, bool isTable = false)
    {
        if (IsValidValue(value))
        {
            InsertTitle(title, doc);
            if (isTable)
                InsertLanguages(doc, value!);
            else
                InsertData(value!, doc, SectionFormatting);
            InsertEmptyParagraphs(doc, 2);
        }
    }

    private void InsertBulletList(DocX doc, string value)
    {
        foreach (var line in value.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries))
            InsertBulletPoint(doc, line.Trim());
    }

    private void InsertBulletPoint(DocX doc, string text)
    {
        var paragraph = doc.InsertParagraph();
        paragraph.IndentationBefore = 20;
        var bulletRun = paragraph.Append("•");
        bulletRun.Font(new Font("Noto Sans Symbols"));
        paragraph.Append(" " + text, SectionFormatting);
    }

    private void InsertTitle(in string titleName, Container container) =>
        InsertData(in titleName, container, TitleFormatting);

    private void InsertOrganizationPreview(DocX doc)
    {
        var table = doc.InsertTable(1, 1);
        table.SetBorder(TableBorderType.Top, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.Bottom, new Border { Color = OrganizationFontColor });
        table.SetBorder(TableBorderType.Right, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.Left, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.InsideV, new Border { Tcbs = BorderStyle.Tcbs_none });
        table.SetBorder(TableBorderType.InsideH, new Border { Tcbs = BorderStyle.Tcbs_none });
        var cell = table.Rows[0].Cells[0];
        cell.MarginBottom = 10;
        InsertData(
            OrganizationPreview,
            cell,
            SmallFormatting,
            Alignment.both);
    }

    private static void InsertEmptyParagraph(DocX doc) => doc.InsertParagraph();

    private static void InsertEmptyParagraphs(DocX doc, int count)
    {
        for (int i = 0; i < count; i++)
            InsertEmptyParagraph(doc);
    }

    private static bool IsValidValue(string? value) => !string.IsNullOrEmpty(value);
}