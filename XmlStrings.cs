using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;

/*  
Этот класс используется для работы с xml-файлом и извлечения из него текстовых строк
для вывода на консоль прогресса сервисной программы РО-2Д

Автор: Евгений Якайтис, НИО-5
*/
namespace RO2D.XmlText
{
    static class XmlStrings
    {
        public static string[] Column3; // столбец "Величина по ТУ" в талице проверки
        public static string[] Column4;// столбец "Допуск" в талице проверки

        /* чтение строк из XML-файла */
        public static string GetStringFromXml(string MethodTag, string TagName)
        {           
            XDocument xdoc = XDocument.Load(Application.StartupPath + @"\XmlText\MethodicsInfo.met");         
            var query = from checkpoint in xdoc.Descendants(MethodTag) let xmltext = checkpoint.Descendants(TagName) select xmltext.First();            
            return FilterString(query.First().ToString());  
        }

        /* поиск и извлечение допусков и значений ТУ для таблицы отчета из файла .xml */
        public static void GetAttributeValFromXml(int MethNum, out string[] Column3, out string[] Column4)
        {
            MethNum++;
            List<string> col3 = new List<string>();
            List<string> col4 = new List<string>();
            byte SmallCounter = 0;
            string[] ArrayOfStrings = new string[15];
            XDocument xdoc = XDocument.Load(Application.StartupPath + @"\XmlText\Thresholds.met");
            var query = from checkpoint in xdoc.Descendants("M" + MethNum.ToString()) let XmlSubTags = checkpoint.Descendants() select XmlSubTags.OfType<XElement>().ToArray();            
            Column4 = null;
            foreach (XElement[] el in query)
            {
                foreach(XElement elem in el)
                {                   
                    col3.Add(elem.Attribute("TehnicalConditionValue").Value);
                    SmallCounter++;
                }                                           
            }
            Column3 = col3.ToArray<string>();

            foreach (XElement[] el in query)
            {
                foreach (XElement elem in el)
                {
                    col4.Add(elem.Attribute("Threshold").Value);
                    SmallCounter++;
                }
            }
            Column4 = col4.ToArray<string>();

        }

        /* Считывание IP адреса генератора при загрузке приложения */
        public static string LoadIpAddrValue()
        {        
            /* создание объекта XML-ридера */
            XmlReader xdoc = new XmlTextReader(Application.StartupPath + @"\XmlText\MethodicsInfo.met");
            /* Считывание из XML-файла */
            do
            {
                xdoc.Read();
            }
            while (xdoc.Name != "M1") ;
            string IP = xdoc.GetAttribute("IP");
            xdoc.Close();          
            return IP;
        }

        /* фильтрация строки прочитанной из xml файла (удаление лишних символов)  */
        static string FilterString(string XmlStringFromFile)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlStringFromFile);
            sb.Remove(0, 5);
            sb.Remove(sb.Length-6, 6);
            return sb.ToString();
        }

        /* извлечение параметров (сам файл встроен в приложение, как ресурс) */
        public static string[] GetParamsFromXML(string param)
        {
            string[] arrayOfChecks;
            XmlTextReader xm = new XmlTextReader(RO_2D.Properties.Resources.SelectionParams, XmlNodeType.Element, null);
            do
            {
                xm.Read();
            }
            while (xm.Name != param);
            string s = xm.ReadElementContentAsString();
            arrayOfChecks = s.Split(',');
            for (byte c=0; c<arrayOfChecks.Length; c++) arrayOfChecks[c] = arrayOfChecks[c].Remove(0, 6);

            return arrayOfChecks;

        }

    }
}
