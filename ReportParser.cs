using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;


/*
 Парсер для сохранения/загрузки отчета
 Автор: Евгений Якайтис 
*/

namespace RO2D
{
    class ReportParser
    {
        
        private static int CurrentAttribute = 0;
        private static int RowsNumber = 0;

        /* загрузка и анализ ранее сохраненного отчета */
        public static void DownloadReport(string SavedReport)
        {
            Regex regexp = new Regex("\\d{1,2}");
      
            Match m; 
          
            ExaminationResults.Tables.Clear(); // очистка списка с таблицами перед новой загрузкой
            MicrosoftWordReport.WordTableDictionary.Clear();
            FormMain.MethodSelector.Invoke(null);

            XDocument xdoc = XDocument.Load(SavedReport);
            var query = from headers in xdoc.Descendants("Report") select headers.Attributes().ToArray();

            /* чтение основных заголовков отчета в программу */
            foreach (XAttribute[] headers in query)
            {
                FormProtocol.DeviceId                        = headers[0].Value;
                FormProtocol.CategoryInspection              = headers[1].Value;
                FormProtocol.KindOfInspection                = headers[2].Value;
                FormProtocol.SubkindOfInspection             = headers[3].Value;
                FormProtocol.SpokesManFlag = Convert.ToBoolean(headers[4].Value);
            }
            /* Поиск в сохраненном файле ранее пройденных таблиц и добавление их в список */
            IEnumerable<XElement> q = from elems in xdoc.Descendants() select elems;
            foreach (XElement xelem in q)
            {
                if (xelem.Name.ToString().Contains("Method")) //  если обнаружен маркер таблицы, то добавить ее в список и прочесть все её атрибуты
                {
                    m = regexp.Match(xelem.Name.ToString());
                    if (m.Success)
                    {
                        ExaminationResults.AddRemoveTable(Convert.ToInt32(m.Value)-1, true);
                        MicrosoftWordReport.FillLists(Convert.ToInt32(m.Value) - 1, true);
                    }
                    /* Чтение непосредственных результатов из пройденных таблиц */
                    var NestedTags = from checkpoint in xdoc.Descendants(xelem.Name.ToString()) let XmlSubTags = checkpoint.Descendants() select XmlSubTags.Attributes();
                    foreach (XAttribute attr in  NestedTags.First())
                    {                   
                        FuelFillTheTable(xelem.Name.ToString(), attr.Value); // наполнение результатами таблиц 
                    }
                }
            }       
        }


        /* Заполнение таблиц результатти */
        private static void FuelFillTheTable(string TabName, string AttributeValue)
        {
            int rowsQuantity = 0;         
            foreach(DataTable dt in ExaminationResults.Tables)
            {
                if(TabName == dt.TableName) // если есть такая таблица в списке, то заполнить ее данными из файла 
                {
                    rowsQuantity = dt.Rows.Count;
                    dt.Rows[RowsNumber].BeginEdit();
                    dt.Rows[RowsNumber][CurrentAttribute] = AttributeValue;
                    dt.Rows[RowsNumber].EndEdit();
                }
            }
            CurrentAttribute++;
          
            if (CurrentAttribute == 3)
            {
                RowsNumber++;
                CurrentAttribute = 0;
            }
            if (RowsNumber == rowsQuantity) RowsNumber = 0;
        }

        /* сохранение образа отчета в промежуточный формат (.parsed)*/
        public static void SaveReport(string File)
        {

            XDocument doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));
            doc.Add(new XElement("Report", new XAttribute("Name", FormProtocol.DeviceId), new XAttribute("InspectionCategory", FormProtocol.CategoryInspection),
                                           new XAttribute("KindOfCategory", FormProtocol.KindOfInspection), new XAttribute("Sub", FormProtocol.SubkindOfInspection),
                                           new XAttribute("Representative", FormProtocol.SpokesManFlag)));

            foreach(DataTable dt in ExaminationResults.Tables)
            {
                doc.Descendants("Report").First().Add(new XElement(dt.TableName));
                for(int c = 0; c < dt.Rows.Count; c++)
                {
                    doc.Descendants(dt.TableName).First().Add(new XElement("Row" + c, new XAttribute("PS1", dt.Rows[c][0]), new XAttribute("PS2", dt.Rows[c][1]), new XAttribute("Result", dt.Rows[c][2])));
                }
               
            }
            doc.Save(File);    
        }
    }
}
