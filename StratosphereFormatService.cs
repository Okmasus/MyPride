using CvParser.Document.DTO;
using System.Globalization;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace CvParser.Document.Implementation.ClientCvFormats
{
    public class StratosphereFormatService : TemplateDocumentFormatService
    {

        public override CreateDocumentResult CreateDocument(CvDTO cvDocumentDTO)
        {
            using var ms = new MemoryStream();

            using var doc = DocX.Load($"Templates/StratosphereTemplate.docx");
            try
            {
                doc.ReplaceText("{Staff.Grade}", cvDocumentDTO.Staff.Grade);

                foreach (var property in cvDocumentDTO.GetType().GetProperties())
                {
                    var value = property.GetValue(cvDocumentDTO)?.ToString() ?? string.Empty;
                    WriteData(doc, property.Name, value);
                }

                var keySkills = cvDocumentDTO.SkillsDto;
                string skills = string.Empty;

                

                doc.ReplaceText($"{cvDocumentDTO.LastName}", $"{cvDocumentDTO.LastName[0]}");

                WriteCollectionBlocks(doc, cvDocumentDTO, "Educations", EducationsWriter);
                CreateWorksInfoTable(doc, cvDocumentDTO);
                return ms;
            }
            finally
            {
                doc.SaveAs(ms);
            }
        }
        private static void EducationsWriter(Paragraph[] educationsParagraphs, Container container, Paragraph previous, CvDTO dto)
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

        public void CreateWorksInfoTable(DocX document, CvDTO cvDTO)
        {
            if (cvDTO.Works.Length == 0) return;
            // 1. Создаем таблицу: 2 строки, 1 колонка
            Table table = document.AddTable(cvDTO.Works.Length * 2, 1);

            // 2. Настраиваем таблицу
            table.Alignment = Alignment.right; // Выравнивание по центру
            table.SetWidths(new float[] { 700f }); // Ширина колонки (в пунктах)

            // 3. Настраиваем границы (опционально)
            var border = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, System.Drawing.Color.Black);
            table.SetBorder(TableBorderType.Top, border);
            table.SetBorder(TableBorderType.Bottom, border);
            table.SetBorder(TableBorderType.Left, border);
            table.SetBorder(TableBorderType.Right, border);
            table.SetBorder(TableBorderType.InsideH, border);

            for (int i = 0, rowIndex = 0; i < cvDTO.Works.Length; i++, rowIndex += 2)
            {
                // Создаем строку для названия проекта (нечетная строка)
                if (rowIndex >= table.Rows.Count)
                {
                    // Добавляем новые строки, если нужно
                    table.InsertRow();
                    table.InsertRow();
                }

                // Заполняем строку с названием проекта (нечетная строка)
                table.Rows[rowIndex].Cells[0].Paragraphs[0].Append($"{cvDTO.Works[i].Title}")
                    .Font("Montserrat")
                    .FontSize(11)
                    .Bold(true);
                table.Rows[rowIndex].Cells[0].FillColor = System.Drawing.Color.FromArgb(241, 194, 50);
                table.Rows[rowIndex].MinHeight = 30;

                // Создаем список деталей
                var SectionTitles = new List<KeyValuePair<string, string>>
    {
        new ("Период", $"{cvDTO.Works[i].Start.ToString("MMMM yyyy", new CultureInfo("ru-RU"))} - {cvDTO.Works[i].End?.ToString("MMMM yyyy", new CultureInfo("ru-RU"))}" ),
        new ("Роль", $"{cvDTO.Works[i].Position}"),
        new("Описание проекта", $"{cvDTO.Works[i].Description.Value}"),
        new("Обязанности", $"{cvDTO.Works[i].Tasks.Value}"),
        new("Стек", $"{cvDTO.Works[i].Tools.Value}")
    };

                // Заполняем строку с деталями (четная строка)
                var detailsParagraph = table.Rows[rowIndex + 1].Cells[0].Paragraphs[0];

                detailsParagraph.LineSpacing = 14.4f;


                for (int j = 0; j < SectionTitles.Count; j++)
                {
                    if (SectionTitles[j].Value.Length > 1)
                    {
                        detailsParagraph.Append($"{SectionTitles[j].Key} ")
                        .Font("Montserrat")
                        .FontSize(11)
                        .Bold(true)
                        .Append($"{SectionTitles[j].Value}")
                        .Font("Montserrat");

                        if (j < SectionTitles.Count - 1)
                        {
                            detailsParagraph.AppendLine(); // Добавляем перенос между секциями
                        }
                    }
                }
            }
            document.ReplaceTextWithObject("{WorkTable}", table);
        }

    }
}
