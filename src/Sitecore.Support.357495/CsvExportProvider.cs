using Sitecore.Diagnostics;
using Sitecore.ExperienceForms.Data;
using Sitecore.ExperienceForms.Data.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Sitecore.Configuration;

namespace Sitecore.Support.ExperienceForms.Client.Data
{
    public class CsvExportProvider : IExportDataProvider
    {
        private static readonly string delimiter = Settings.GetSetting("Sitecore.ExperienceForms.ExportDataDelimiter", ",");

        private readonly IFormDataProvider _formDataProvider;

        public CsvExportProvider(IFormDataProvider formDataProvider)
        {
            Assert.ArgumentNotNull(formDataProvider, "formDataProvider");
            _formDataProvider = formDataProvider;
        }

        public ExportDataResult Export(Guid formId, DateTime? startDate, DateTime? endDate)
        {
            IEnumerable<FormEntry> enumerable = _formDataProvider.GetEntries(formId, startDate, endDate).AsEnumerable();
            return new ExportDataResult
            {
                Content = ((enumerable == null) ? string.Empty : GenerateFileContent(enumerable)),
                FileName = GenerateFileName(formId, startDate, endDate)
            };
        }

        protected virtual string GenerateFileName(Guid formId, DateTime? startDate, DateTime? endDate)
        {
            string text = string.Empty;
            if (startDate.HasValue)
            {
                text = text + "_from_" + DateUtil.ToIsoDate(startDate.Value);
            }
            if (endDate.HasValue)
            {
                text = text + "_until_" + DateUtil.ToIsoDate(endDate.Value);
            }
            if (string.IsNullOrEmpty(text))
            {
                text = "-" + DateUtil.IsoNow;
            }
            return FormattableString.Invariant($"Form-Data{text}.csv");
        }

        protected virtual string GenerateFileContent(IEnumerable<FormEntry> formEntries)
        {
            Assert.ArgumentNotNull(formEntries, "formEntries");
            IOrderedEnumerable<FormEntry> orderedEnumerable = from item in formEntries
                                                              orderby item.Created descending
                                                              select item;
            List<FieldData> fieldColumnsList = new List<FieldData>();
            StringBuilder stringBuilder = new StringBuilder();
            foreach (FormEntry item in orderedEnumerable)
            {
                fieldColumnsList.AddRange(from x in item.Fields
                                          where fieldColumnsList.All((FieldData c) => c.FieldItemId != x.FieldItemId)
                                          select x);
            }
            if (fieldColumnsList.Count == 0)
            {
                return string.Empty;
            }
            stringBuilder.AppendFormat(CultureInfo.InvariantCulture, "Created{0}", delimiter);
            stringBuilder.AppendLine(string.Join(delimiter, (from f in fieldColumnsList
                                                       select f.FieldName).ToArray()));
            int count = fieldColumnsList.Count;
            foreach (FormEntry item2 in orderedEnumerable)
            {
                stringBuilder.AppendFormat(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd HH:mm}{1}", item2.Created, delimiter);
                string[] array = new string[count];
                for (int i = 0; i < count; i++)
                {
                    FieldData fieldItem = fieldColumnsList[i];
                    array[i] = EscapeCsvDelimiters(item2.Fields.FirstOrDefault((FieldData f) => f.FieldItemId == fieldItem.FieldItemId)?.Value);
                }
                stringBuilder.AppendLine(string.Join(delimiter, array));
            }
            return stringBuilder.ToString();
        }

        private static string EscapeCsvDelimiters(string fieldValue)
        {
            if (!string.IsNullOrEmpty(fieldValue))
            {
                fieldValue = fieldValue.Replace("\"", "\"\"");
                if (fieldValue.IndexOf(Environment.NewLine, 0, StringComparison.OrdinalIgnoreCase) >= 0 || fieldValue.IndexOf(delimiter, StringComparison.OrdinalIgnoreCase) >= 0 || fieldValue.IndexOf("\"", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    fieldValue = FormattableString.Invariant($"\"{fieldValue}\"");
                }
                fieldValue = fieldValue.Replace(Environment.NewLine, " ");
            }
            return fieldValue;
        }
    }
}