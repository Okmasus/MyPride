using CvParser.DataAccess.Models;
using CvParser.Document.DTO;
using CvParser.Document.Interfaces;
using CvParser.Document.Interfaces.ClientCvFormats;
using System.Drawing;
using System.Globalization;
using System.Text;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Font = Xceed.Document.NET.Font;

namespace CvParser.Document.Implementation.ClientCvFormats;

public sealed class SSPSoftFormatService : TemplateDocumentFormatService, IDataInserter
{
    private const string ShortInfoAboutStafferTitle = "Краткая информация о специалисте на соответствие вакансии";
    private const string TechnicalSkillsTitle = "Технические навыки";
    private const string WorkExperienceTitle = "Профессиональный опыт";
    private const string AdditionalDataTitle = "Дополнительная информация";
    private const string EducationsSubTitle = "Образование";
    private const string LanguagesTitle = "Иностранный язык";
    private readonly ISortedKeySkill _sortedKeySkill;
    

    public SSPSoftFormatService(ISortedKeySkill sortedKeySkill)
    {
        _sortedKeySkill = sortedKeySkill;
    }

    private readonly Formatting _boldHeaderFormatting = new()
    {
        Bold = true,
        Size = 14
    };

    private readonly Formatting _headerFontSizeFormatting = new()
    {
        Size = 14
    };

    private const float LeftColumnWidth = 5.45f * 28.35f;
    private const float RightColumnWidth = 10.42f * 28.35f;
    // 1 пункт в XCEED = 28.35, поэтому если кто-то захочет поменять размер колонок, то имейте ввиду(Левое значеине - это то какой размер мы хотим, а правый - константа)

    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        using var ms = new MemoryStream();
        using var doc = DocX.Create(ms);

        doc.SetDefaultFont(new Font("Calibri"));

        doc.AddHeaders();
        doc.AddFooters();

        doc.Footers.Even.PageNumbers = true;
        doc.Footers.Odd.PageNumbers = true;
        doc.DifferentFirstPage = true;

        InsertImage(doc, doc.Headers.First);
        InsertHeaderSection(doc, cvDocumentDTO);
        InsertAboutSection(doc, cvDocumentDTO);
        InsertTechnicalSkillsSection(doc, _sortedKeySkill.FillSortedDto(cvDocumentDTO));
        InsertWorkExperienceSection(doc, cvDocumentDTO);
        InsertAdditionalDataSection(doc, cvDocumentDTO);

        doc.SaveAs(ms);
        return ms;
    }

    public void InsertData(in string? paragraphTextContent, Container container, Formatting? formatting = null, Alignment alignment = Alignment.left) =>
        container.InsertParagraph().InsertText(paragraphTextContent, formatting: formatting);

    private static Table InsertTable(DocX doc, in int columnsCount)
    {
        var table = doc.InsertTable(1, columnsCount);
        HideAllBordersOfTable(table);
        table.SetBorder(TableBorderType.Top, new Border
        {
            Size = BorderSize.five,
            Color = Color.FromArgb(31, 73, 125)
        });

        return table;
    }

    private void InsertHeaderSection(DocX doc, CvDTO cvDocumentDTO)
    {
        InsertData($"{cvDocumentDTO?.LastName} {cvDocumentDTO?.FirstName} {cvDocumentDTO?.SurName}", doc, new Formatting
        {
            Size = 18
        });

        if (cvDocumentDTO.Staff is not null)
            InsertData($"{cvDocumentDTO.Stack} ({cvDocumentDTO.Staff!.Grade})", doc, _headerFontSizeFormatting);

        if (cvDocumentDTO.DateOfBirth is not null)
        {
            var dateOfBirth = DateTime.ParseExact(cvDocumentDTO.DateOfBirth, "yyyy-MM-dd", CultureInfo.InvariantCulture);
            InsertData($"Возраст: {cvDocumentDTO.Age}({dateOfBirth:dd.MM.yyyy}) ", doc, _headerFontSizeFormatting);
        }

        InsertData($"GMT{(cvDocumentDTO.TimeZone >= 0 ? '+' : '-')}{cvDocumentDTO.TimeZone} ({cvDocumentDTO.Location})", doc, _headerFontSizeFormatting);

        InsertEmptyParagraph(doc);
    }

    private void InsertAboutSection(DocX doc, CvDTO cv)
    {
        InsertData(ShortInfoAboutStafferTitle, doc, _boldHeaderFormatting);

        if (!string.IsNullOrWhiteSpace(cv.About?.Value))
        {
            foreach (var paragraph in cv.About.Value.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                var p = doc.InsertParagraph();
                p.Append($"• {paragraph.Trim()}");
            }
            InsertEmptyParagraph(doc);
            InsertEmptyParagraph(doc);
        }
    }

    private void InsertTechnicalSkillsSection(DocX doc, SortedSkillsDTO sortedSkills)
    {
        InsertData(TechnicalSkillsTitle, doc, _boldHeaderFormatting);
        var skillsTable = doc.InsertTable(1, 2);
        ConfigureSkillsTableStyle(skillsTable);
        skillsTable.SetWidths(new float[] { LeftColumnWidth, RightColumnWidth });

        void InsertSkillCategory(string categoryName, List<KeySkill> skills)
        {
            if (skills == null || !skills.Any(s => !string.IsNullOrWhiteSpace(s?.Value)))
                return;

            var skillsText = string.Join(" / ", skills
                .Where(s => !string.IsNullOrWhiteSpace(s.Value))
                .OrderBy(s => s.Order)
                .Select(s => s.Value));

            var row = skillsTable.InsertRow();
            row.Cells[0].FillColor = SetDefaultColor();
            row.Cells[0].Paragraphs.First()
               .Append(categoryName)
               .Bold()
               .FontSize(11);
            row.Cells[1].Paragraphs.First()
               .Append(skillsText)
               .FontSize(11);
        }
        InsertSkillCategory("Методологии", sortedSkills.Methodologies);
        InsertSkillCategory("Инструменты", sortedSkills.Tools);
        InsertSkillCategory("Языки программирования", sortedSkills.Languages);
        InsertSkillCategory("Базы данных", sortedSkills.Databases);
        InsertSkillCategory("Фреймворки и библиотеки", sortedSkills.FrameworksAndLibraries);
        InsertSkillCategory("Нотации и языки моделирования", sortedSkills.ModelingNotationsAndLanguages);
        InsertSkillCategory("Машинное обучение", sortedSkills.ML);
        InsertSkillCategory("Искусственный интеллект", sortedSkills.AI);
        InsertSkillCategory("Другие навыки", sortedSkills.NoCategory);

        skillsTable.RemoveRow(0);
        InsertEmptyParagraph(doc);
        InsertEmptyParagraph(doc);
    }

    private void ConfigureSkillsTableStyle(Table table)
    {
        HideAllBordersOfTable(table);
        table.SetBorder(TableBorderType.Top, new Border
        {
            Size = BorderSize.five,
            Color = Color.FromArgb(31, 73, 125)
        });
    }

    private void InsertRow(Table table, params string[] data)
    {
        var row = table.InsertRow();
        var cells = row.Cells;

        cells[0].FillColor = SetDefaultColor();
        cells[0].SetBorder(TableCellBorderType.TopLeftToBottomRight, new Border
        {
            Color = SetDefaultColor()
        });

        for (int i = 0; i < data.Length; i++)
            InsertData(data[i], cells[i]);
    }

    private static void InsertImage(DocX doc, Header header)
    {
        using var ms = new MemoryStream(Properties.Resources.SSPSoftLogo);
        var logo = doc.AddImage(ms);

        var paragraph = header.InsertParagraph();
        paragraph.Alignment = Alignment.right;
        paragraph.AppendPicture(logo.CreatePicture(50.094f, 193.02f));
    }

    private void InsertWorkExperienceSection(DocX doc, CvDTO cv)
    {
        InsertData(WorkExperienceTitle, doc, _boldHeaderFormatting);

        foreach (var work in cv.Works)
        {
            var workTable = doc.InsertTable(1, 2);
            ConfigureWorkTableStyle(workTable);
            workTable.SetWidths(new float[] { LeftColumnWidth, RightColumnWidth });
            // Левая колонка
            var leftCell = workTable.Rows[0].Cells[0];
            leftCell.FillColor = SetDefaultColor();
            leftCell.Paragraphs.First()
                .Append(work.Title)
                .Bold()
                .FontSize(11)
                .AppendLine()
                .Append($"{work.Start:dd.MM.yyyy} - {work.End?.ToString("dd.MM.yyyy") ?? "настоящее время"} ({work.Duration})")
                .FontSize(11);

            
            //Правая колонка
            var rightCell = workTable.Rows[0].Cells[1];
            var currentParagraph = rightCell.Paragraphs.First();

            currentParagraph.Append("Описание проекта:").Bold().AppendLine()
                          .Append(work.Description?.Value ?? "").AppendLine();

            currentParagraph.Append("Роль в проекте:")
                .Bold()
                .AppendLine()
                .Append(work.Position).AppendLine();

            if (!string.IsNullOrWhiteSpace(work.Tasks?.Value))
            {
                currentParagraph.Append("Обязанности:").Bold().AppendLine();
                foreach (var duty in work.Tasks.Value.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    currentParagraph.Append($"• {duty.Trim()}").AppendLine();
                }
            }

            if (!string.IsNullOrWhiteSpace(work.Results?.Value))
            {
                currentParagraph.Append("Достижения:").Bold().AppendLine();
                foreach (var achievement in work.Results.Value.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    currentParagraph.Append($"• {achievement.Trim()}").AppendLine();
                }
            }

            if (!string.IsNullOrWhiteSpace(work.Tools?.Value))
            {
                currentParagraph.Append("Основные технологии проекта:").Bold().AppendLine()
                          .Append($"{work.Tools?.Value}").AppendLine();
            }

            if (!string.IsNullOrWhiteSpace(work.Team?.Value))
            {
                currentParagraph.Append("Состав команды:").Bold().AppendLine()
                              .Append($"{work.Team.Value}");
            }

            doc.InsertParagraph();
        }
        InsertEmptyParagraph(doc);
    }

    private void ConfigureWorkTableStyle(Table table)
    {
        table.SetBorder(TableBorderType.Top, new Border
        {
            Size = BorderSize.one,
            Color = Color.FromArgb(31, 73, 125)
        });

        var noBorder = new Border(
         BorderStyle.Tcbs_none,
         BorderSize.one,
         0,
         Color.Transparent);
        table.SetBorder(TableBorderType.Bottom, noBorder);
        table.SetBorder(TableBorderType.Left, noBorder);
        table.SetBorder(TableBorderType.Right, noBorder);
        table.SetBorder(TableBorderType.InsideV, noBorder);
        table.SetBorder(TableBorderType.InsideH, noBorder);
    }

    private void InsertAdditionalDataSection(DocX doc, CvDTO cv)
    {
        InsertData(AdditionalDataTitle, doc, _boldHeaderFormatting);

        var additionalDataTable = InsertTable(doc, 2);
        additionalDataTable.SetWidths(new float[] { LeftColumnWidth, RightColumnWidth });

        if (cv.Educations != null && cv.Educations.Length > 0)
        {
            var educationsContent = new StringBuilder();
            foreach (var education in cv.Educations)
            {
                educationsContent.Append($"{education.Title}");

                if (!string.IsNullOrEmpty(education.Description?.Value))
                    educationsContent.Append($", {education.Description.Value}");

                educationsContent.AppendLine();
            }
            InsertRow(additionalDataTable, EducationsSubTitle, educationsContent.ToString().Trim());
        }
        if (!string.IsNullOrWhiteSpace(cv.Languages))
        {
            InsertRow(additionalDataTable, LanguagesTitle, cv.Languages);
        }
        if (additionalDataTable.RowCount > 0 &&
            string.IsNullOrEmpty(additionalDataTable.Rows[0].Cells[0].Paragraphs[0].Text))
        {
            additionalDataTable.RemoveRow(0);
        }
    }

    private static void InsertEmptyParagraph(DocX doc) => doc.InsertParagraph();

    private static Color SetDefaultColor() => Color.FromArgb(243, 243, 243);

    private static void HideAllBordersOfTable(Table table)
    {
        table.SetBorder(TableBorderType.Bottom, new Border
        {
            Tcbs = BorderStyle.Tcbs_none
        });
        table.SetBorder(TableBorderType.Right, new Border
        {
            Tcbs = BorderStyle.Tcbs_none
        });
        table.SetBorder(TableBorderType.Left, new Border
        {
            Tcbs = BorderStyle.Tcbs_none
        });
        table.SetBorder(TableBorderType.InsideV, new Border
        {
            Tcbs = BorderStyle.Tcbs_none
        });
        table.SetBorder(TableBorderType.InsideH, new Border
        {
            Tcbs = BorderStyle.Tcbs_none
        });
    }
}