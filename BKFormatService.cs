using CvParser.Document.DTO;
using CvParser.Document.Extensions;
using CvParser.Document.Interfaces.ClientCvFormats;
using System.Drawing;
using Xceed.Document.NET;
using Xceed.Words.NET;
using CvParser.Extensions;
using Font = Xceed.Document.NET.Font;
using System.Globalization;
using System.Reflection;

namespace CvParser.Document.Implementation.ClientCvFormats;

public sealed class BKFormatService : TemplateDocumentFormatService
{
    private const string EducationsPlaceName = "Educations";
    private const string AdditionalEducationsPlaceName = "AdditionalEducations";
    private const string EducationsBlockName = "Образование";
    private const string AdditionalEducationsBlockName = "Курсы";

    private static readonly Dictionary<int, string> SectionTitles = new Dictionary<int, string>
    {
        { 0, "Период работы на проекте" },
        { 1, "Длительность работы на проекте" },
        { 2, "Роль на проекте" },
        { 3, "Описание проекта" },
        { 4, "Команда проекта" },
        { 5, "Используемые технологии" },
        { 6, "Обязанности на проекте" },
        { 7, "Личные результаты" }
    };

    private static readonly Border _noneBorderStyle = new()
    {
        Tcbs = BorderStyle.Tcbs_none
    };

    public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
    {
        var ms = new MemoryStream();

        using var doc = DocX.Load($"Templates/VKFormatTemplate.docx");
        try
        {
            DocumentFormat(cvDocumentDTO, doc);

            InsertSomethingEducations(doc, cvDocumentDTO.Educations
                .Where(education => !education.IsAdditional)
                .ToArray(), EducationsPlaceName, EducationsBlockName);

            InsertSomethingEducations(doc, cvDocumentDTO.Educations
                .Where(education => education.IsAdditional)
                .ToArray(), AdditionalEducationsPlaceName, AdditionalEducationsBlockName);

            ReplaceTablesPlaceholder(CreateTables(cvDocumentDTO, doc), doc, cvDocumentDTO);

            string contacts = string.Join(", ", cvDocumentDTO.Contacts.Where(x => x != null).Select(x => x!.Value))
                              + (string.IsNullOrEmpty(cvDocumentDTO.Phone) ? "" : $", {cvDocumentDTO.Phone}");

            if (contacts == "")
            {
                doc.ReplaceText("<<Contacts>>", "");
            }
            else doc.ReplaceText("<<Contacts>>", $"E-mail / номер телефона: {contacts}");

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
        var cvProps = ConvertToDictionary(value);
        ReplacePlaceholdersInDocX(doc, cvProps);
    }

    private static Dictionary<string, object> PropDictionary = new Dictionary<string, object>();

    public static Dictionary<string, object> ConvertToDictionary<T>(T dto)
    {
        if (dto == null)
        {
            return new Dictionary<string, object>(); // Возвращаем новый словарь вместо статического
        }

        // Очищаем статический словарь перед началом
        PropDictionary.Clear();

        // Запускаем рекурсивное заполнение
        FillDictionary(dto, "");

        return new Dictionary<string, object>(PropDictionary); // Возвращаем копию, чтобы избежать модификаций извне
    }

    private static void FillDictionary<T>(T obj, string prefix, HashSet<object> visitedObjects = null)
    {
        if (obj == null) return;

        // Инициализируем HashSet при первом вызове
        if (visitedObjects == null)
            visitedObjects = new HashSet<object>();

        // Если объект уже был обработан, выходим
        if (visitedObjects.Contains(obj))
            return;

        // Добавляем текущий объект в посещённые
        visitedObjects.Add(obj);

        Type type = obj.GetType();
        PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (PropertyInfo property in properties)
        {
            if (!property.CanRead) continue;

            // Пропускаем индексированные свойства
            if (property.GetIndexParameters().Length > 0)
                continue;

            string fullKey = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";
            object value;

            try
            {
                value = property.GetValue(obj);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting value of {fullKey}: {ex.Message}");
                continue;
            }

            if (value != null && IsComplexType(property.PropertyType))
            {
                // Рекурсивно обрабатываем вложенные свойства, передавая HashSet
                FillDictionary(value, fullKey, visitedObjects);
            }
            else
            {
                PropDictionary[fullKey] = value;
            }
        }
    }

    private static bool IsComplexType(Type type)
    {
        // Проверяем, является ли тип "простым" (не нужно рекурсивно разбирать)
        return !type.IsPrimitive &&
               type != typeof(string) &&
               type != typeof(decimal) &&
               type != typeof(DateTime) &&
               !type.IsEnum;
    }

    public static void ReplacePlaceholdersInDocX(DocX doc, Dictionary<string, object> properties)
    {
        // Проходим по всем параграфам в документе
        foreach (var paragraph in doc.Paragraphs)
        {
            string text = paragraph.Text;

            // Заменяем плейсхолдеры в тексте
            foreach (var prop in properties)
            {
                string placeholder = $"<<{prop.Key}>>";

                // Проверяем, является ли значение коллекцией
                string value;
                if (prop.Value is System.Collections.IEnumerable enumerable && !(prop.Value is string))
                {
                    // Преобразуем все элементы коллекции в строки и объединяем через запятую
                    var items = new List<string>();
                    foreach (var item in enumerable)
                    {
                        items.Add(item?.ToString() ?? string.Empty);
                    }

                    value = string.Join(", ", items);
                }
                else
                {
                    value = prop.Value?.ToString() ?? string.Empty;
                }

                if (text.Contains(placeholder))
                {
                    paragraph.ReplaceText(placeholder, value);
                }
            }
        }
    }

    private static void InsertSomethingEducations(DocX doc, EducationDTO[] educations, string placeName,
        string blockName)
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
                    pare.ReplaceText($"<<{placeName}.Title>>", educations[icv].Title);
                    pare.ReplaceText($"<<{placeName}.Description>>", educations[icv].Description.Value);
                }
            }

            for (var i = educationParagraphs.Count - 1; i >= 0; i--)
            {
                var pareFierst = educationParagraphs[i];
                pareFierst.ReplaceText($"<<{placeName}.Title>>", educations[0].Title);
                pareFierst.ReplaceText($"<<{placeName}.Description>>", educations[0].Description.Value);
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

    private static List<Table> CreateTables(CvDTO cvDTO, DocX doc)
    {
        List<Table> tables = new List<Table>();

        for (int i = 0; i < cvDTO.Works.Length; i++)
        {
            Table table = doc.AddTable(1, 2);

            // Настраиваем стиль таблицы            
            SetTableDesign(table);

            // Создание таблицы
            for (int j = 0; j < SectionTitles.Count; j++)
            {
                if (WorkExperienceDescription(cvDTO.Works[i], j) != string.Empty)
                {
                    table.InsertRow();
                    table.Rows[j].Cells[0].Paragraphs[0]
                        .Append(SectionTitles[j], new Formatting() { Bold = true, Size = 12 });
                    table.Rows[j].Cells[1].Paragraphs[0].Append(WorkExperienceDescription(cvDTO.Works[i], j),
                        new Formatting() { Size = 12 });
                }
            }

            table.Rows[table.Rows.Count - 1].Remove();

            tables.Add(table);
        }

        return tables;
    } //Special

    private static void ReplaceTablesPlaceholder(List<Table> tables, DocX doc, CvDTO cvDTO)
    {
        Paragraph placeholder = doc.Paragraphs.FirstOrDefault(p => p.Text.Contains("<<Tables>>"));

        if (placeholder != null)
        {
            // Удаляем плейсхолдер
            placeholder.RemoveText(0);

            // Вставляем таблицы на место плейсхолдера
            int i = 0;
            foreach (var table in tables)
            {
                placeholder.InsertTableAfterSelf(table);
                // Добавляем пустую строку между таблицами (опционально)
                placeholder.InsertParagraphAfterSelf($"\n {cvDTO.Works[i].Title} \n", false,
                        new Formatting() { Size = 14, Bold = true })
                    .Alignment = Alignment.center;
                i++;
            }
        }
    }

    private static void SetTableDesign(Table table)
    {
        table.Design = TableDesign.TableNormal;
        table.Alignment = Alignment.center;
        var border = new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.FromArgb(189, 215, 237));
        var borderTypes = new[]
        {
            TableBorderType.Top, TableBorderType.Bottom, TableBorderType.Left, TableBorderType.Right,
            TableBorderType.InsideH, TableBorderType.InsideV
        };
        foreach (var borderType in borderTypes)
            table.SetBorder(borderType, border);
        table.SetColumnWidth(0, 150);
        table.SetColumnWidth(1, 324);
    }

    private static string WorkExperienceDescription(WorkDTO experience, int descSubTitle)
    {
        if (experience == null) return string.Empty;


        switch (descSubTitle)
        {
            case 0:
                return experience.End is null
                    ? string.Empty
                    : $"{experience.Start.ToString("MMMM yyyy", new CultureInfo("ru-RU"))} - {experience.End?.ToString("MMMM yyyy", new CultureInfo("ru-RU"))}";

            case 1:
                return experience.Duration ?? string.Empty;
            case 2:
                return $"{experience.Position}" ?? string.Empty;

            case 3:
                return experience.Description.Value ?? string.Empty;

            case 4:
                return experience.Team.Value ?? string.Empty;

            case 5:
                return experience.Tools.Value ?? string.Empty;

            case 6:
                return experience.Tasks.Value ?? string.Empty;

            case 7:
                return experience.Results.Value ?? string.Empty;

        }

        throw new NotImplementedException();
    } //Special
}
