using CvParser.DataAccess.Models;
using CvParser.Document.DTO;
using CvParser.Document.Interfaces.ClientCvFormats;
using CvParser.Services.DTO;
using System.Drawing;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace CvParser.Document.Implementation.ClientCvFormats;

/// <summary>
/// <inheritdoc/>
/// </summary>
public class ITFBFormatService : TemplateDocumentFormatService
{
    private const string EducationsPlaceName = "Educations";
    private const string AdditionalEducationsPlaceName = "AdditionalEducations";
    private const string EducationsBlockName = "Образование";
    private const string AdditionalEducationsBlockName = "Курсы";
    private static readonly Dictionary<int, string> SectionTitles = new Dictionary<int, string>
{
    { 0, "Обязанности" },
    { 1, "Достижения" },
    { 2, "Состав команды" },
    { 3, "Технологии" }
};
    private static readonly Border _noneBorderStyle = new()
    {
        Tcbs = BorderStyle.Tcbs_none
    };

    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        using var ms = new MemoryStream();

        using var doc = DocX.Load($"Templates/ITFBTemplate.docx");
        try
        {
            DocumentFormat(cvDocumentDTO, doc);
            InsertSomethingEducations(doc, cvDocumentDTO.Educations.Where(education => !education.IsAdditional).ToArray(), EducationsPlaceName, EducationsBlockName);
            InsertSomethingEducations(doc, cvDocumentDTO.Educations.Where(education => education.IsAdditional).ToArray(), AdditionalEducationsPlaceName, AdditionalEducationsBlockName);

            CreateTable(cvDocumentDTO, doc);

            return ms;
        }
        finally
        {
            doc.SaveAs(ms);
        }
    }
    /// <summary>
    /// Приведение документа к формату.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="value">Формат.</param>
    /// <param name="doc">Документ.</param>
    /// <returns></returns>
    private static void DocumentFormat<T>(T value, DocX doc)
    {
        if (value is CvDTO)
        {
            foreach (var prop in value.GetType().GetProperties())
            {
                var valueProp = prop.GetValue(value);

                if (valueProp is string strValue && strValue.Equals("Неизвестно"))
                    prop.SetValue(value, string.Empty);
            }
        }

        var properties = typeof(T)
            .GetProperties()
            .ToDictionary(
                key => key.Name,
                v => (Func<string>)(() => v.GetValue(value)?.ToString() ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

        var regexSimple = new Regex(@"^{([a-zA-Z]*?)}$");        // для {Name}
        var regexWithIndex = new Regex(@"^{([a-zA-Z]*?)\[(\d*)\]}$"); // для {Name[0]}

        var frtd = new FunctionReplaceTextOptions()
        {
            FindPattern = @"{[a-zA-Z]*?\[\d*?\]}", // ищем только плейсхолдеры с индексами
            RegExOptions = RegexOptions.IgnoreCase,
            RegexMatchHandler = (s) =>
            {
                var match = regexWithIndex.Match(s);
                if (!match.Success)
                    return s; // если не подходит под шаблон, оставляем как есть

                var valueName = match.Groups[1].Value;
                if (valueName.Equals("Skills", StringComparison.OrdinalIgnoreCase))
                    return s; // оставляем {Skills[i]} нетронутым

                var index = int.Parse(match.Groups[2].Value);
                if (properties.TryGetValue(valueName, out var func))
                {
                    var result = func();
                    return (string.IsNullOrEmpty(result) || index >= result.Length)
                        ? string.Empty
                        : result[index].ToString();
                }
                return s; // если свойства нет, оставляем плейсхолдер
            }
        };
        doc.ReplaceText(frtd);

        var frt = new FunctionReplaceTextOptions()
        {
            FindPattern = @"{[a-zA-Z]*?}", // ищем плейсхолдеры без индексов
            RegExOptions = RegexOptions.IgnoreCase,
            RegexMatchHandler = (s) =>
            {
                var match = regexSimple.Match(s);
                if (!match.Success)
                    return s;

                var valueName = match.Groups[1].Value;


                if (properties.TryGetValue(valueName, out var func))
                {
                    var result = func();
                    return string.IsNullOrWhiteSpace(result)
                        ? string.Empty // удаляем пустой плейсхолдер
                        : result;
                }

                return s;
            }
        };
        doc.ReplaceText(frt);
        var paragraphs = doc.Paragraphs;

        for (int i = 0; i < paragraphs.Count - 1; i++)
        {
            var current = paragraphs[i];
            var next = paragraphs[i + 1];

            // Если текущий параграф — "Навыки", а следующий — пустой
            if (current.Text.Trim().Equals("Навыки", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(next.Text) || next.Text == "{Skills}")
            {
                // Удаляем оба
                doc.RemoveParagraph(next);
                doc.RemoveParagraph(current);
                break;
            }
        }
    }

    private static string SwapCountryAndCity(string location)
    {
        if (string.IsNullOrEmpty(location))
            return location;

        string[] parts = location.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

        if (parts.Length != 2)
            return location;

        string country = parts[0].Trim();
        string city = parts[1].Trim();

        return $"{city}, {country}";
    }
    private static void InsertSomethingEducations(DocX doc, EducationDTO[] educations, string placeName, string blockName)
    {
        var educationParagraphs = doc.Paragraphs
                        .SkipWhile(paragraph => !paragraph.Text.Contains($"[{placeName}]"))
                        .Skip(1)
                        .TakeWhile(paragraph => !paragraph.Text.Contains($"[{placeName}]"))
                        .ToList();
        if (educations.Length > 0)
        {
            for (var icv = educations.Length - 1; icv >= 1; icv--)
            {
                for (var i = educationParagraphs.Count - 1; i >= 0; i--)
                {
                    var pare = educationParagraphs[^1].InsertParagraphAfterSelf(educationParagraphs[i]);
                    pare.ReplaceText($"{{{placeName}.Title}}", educations[icv].Title);
                    pare.ReplaceText($"{{{placeName}.Description}}", educations[icv].Description.Value);
                }
            }

            for (var i = educationParagraphs.Count - 1; i >= 0; i--)
            {
                var pareFierst = educationParagraphs[i];
                pareFierst.ReplaceText($"{{{placeName}.Title}}", educations[0].Title);
                pareFierst.ReplaceText($"{{{placeName}.Description}}", educations[0].Description.Value);
            }
        }
        else
        {
            doc.Paragraphs
                .Where(p => p.Text == blockName)
                .ToList()
                .ForEach(paragraph => paragraph.Remove(false));
            educationParagraphs.ForEach(paragraph => paragraph.Remove(false));
        }

        doc.Paragraphs
            .Where(paragraph => paragraph.Text == $"[{placeName}]")
            .ToList()
            .ForEach(paragraph => paragraph.Remove(false));
    }

    private static void CreateTable(CvDTO cvDTO, DocX doc)
    {
        doc.ReplaceText("{LocationSpecial}", SwapCountryAndCity(cvDTO.Location));

        Table table = doc.AddTable(cvDTO.Works.Length, 2);

        // Настраиваем стиль таблицы
        table.Design = TableDesign.TableNormal;
        table.Alignment = Alignment.left;
        table.SetColumnWidth(0, 150);
        table.SetColumnWidth(1, 324);


        var DescTitlesfontFormat = new Formatting
        {
            FontFamily = new Xceed.Document.NET.Font("Open Sans"), // Название шрифта
            Size = 11,                // Размер шрифта
            FontColor = Color.FromArgb(4, 131, 248) // Цвет в RGB
        };
        var workExpCellFormat = new Formatting
        {
            FontFamily = new Xceed.Document.NET.Font("Open Sans"),
            Size = 10,
            FontColor = Color.FromArgb(174, 170, 170)
        };
        var companyTitleFormat = new Formatting
        {
            FontFamily = new Xceed.Document.NET.Font("Open Sans"),
            Size = 11,
            Bold = true,
            FontColor = Color.FromArgb(64, 64, 64)
        };
        var descSubTitleFormat = new Formatting()
        {
            FontFamily = new Xceed.Document.NET.Font("Open Sans"),
            Size = 11,
            FontColor = Color.FromArgb(4, 131, 248),
            Bold = true
        };


        var border = new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Transparent);
        var borderTypes = new[] {
            TableBorderType.Top, TableBorderType.Bottom,
            TableBorderType.Left, TableBorderType.Right,
            TableBorderType.InsideH, TableBorderType.InsideV
};
        foreach (var type in borderTypes) table.SetBorder(type, border);



        for (int i = 0; i < cvDTO.Works.Length; i++)
        {
            table.Rows[i].Cells[0].Paragraphs[0].Append($"{cvDTO.Works[i].Start.ToString("MMMM yyyy", new CultureInfo("ru-RU"))} - {cvDTO.Works[i].End?.ToString("MMMM yyyy", new CultureInfo("ru-RU"))}", workExpCellFormat);
            table.Rows[i].Cells[0].Paragraphs[0].Alignment = Alignment.right;
            table.Rows[i].Cells[1].Paragraphs[0].Append($"{cvDTO.Works[i].Title}\n{cvDTO.Works[i].Position}\n", companyTitleFormat);
            table.Rows[i].Cells[1].Paragraphs[0].Append($"{cvDTO.Works[i].Description.Value}\n");


            for (int j = 0; j < SectionTitles.Count; j++)
            {
                if (ITFBWorkExperienceDescription(cvDTO.Works[i], j) != string.Empty)
                {
                    table.Rows[i].Cells[1].Paragraphs[0].Append($"\n{SectionTitles[j]}\n", descSubTitleFormat);
                    List bulletedList = doc.AddList(null, 0, ListItemType.Bulleted, null, false, false, workExpCellFormat);
                    foreach (var item in SplitStringByNewLine($"{ITFBWorkExperienceDescription(cvDTO.Works[i], j)}"))
                    {
                        if (item.StartsWith("-"))
                        {
                            doc.AddListItem(bulletedList, "•  " + item.Substring(1)); 
                        }

                        else doc.AddListItem(bulletedList, item);
                    }

                    string listAsString = string.Join("\n", bulletedList.Items.Select(item => item.Text))
                        .Replace("Что сделал:\n", string.Empty)
                        .Replace("Что сделал: \n", string.Empty)
                        .Replace("Команда:", string.Empty)
                        .Replace("Инструменты/Технологии:", string.Empty)
                        .Replace("Достижения:\n", string.Empty);

                    table.Rows[i].Cells[1].Paragraphs[0].Append($"{listAsString}\n");
                }
            }
        }
        doc.ReplaceTextWithObject("{WorkTable}", table);

    }

    private static string ITFBWorkExperienceDescription(WorkDTO experience, int descSubTitle)
    {
        if (experience == null) return string.Empty;


        switch (descSubTitle)
        {
            case 0:
                if (experience.Tasks.Value.Length == 0) return string.Empty;
                else return $"{experience.Tasks?.Value}";


            case 1:
                if (experience.Results.Value.Length == 0) return string.Empty;
                else return $"{experience.Results?.Value}";
            case 2:
                if (experience.Team.Value.Length == 0) return string.Empty;
                else return $"{experience.Team?.Value}";

            case 3:
                if (experience.Tools.Value.Length == 0) return string.Empty;
                else return $"{experience.Tools?.Value}";

        }
        throw new NotImplementedException();
    }

    public static string[] SplitStringByNewLine(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return Array.Empty<string>();
        }

        // Разбиваем строку по \n и удаляем пустые записи
        return input.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
    }

}
