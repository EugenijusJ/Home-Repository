using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using RO2D.XmlText;
using System.Data;
using System.Threading.Tasks;

namespace RO2D
{ 
/* 
    Класс по созданию отчета в формате PDF c использованием движка MS Word
    Автор: Евгений Якайтис
*/
    partial class MicrosoftWordReport : IDisposable
    {
        /* переменные для работы с текстом */
        private Type t;
        private dynamic WordApplication;
        private dynamic WordDocument;
        private dynamic SelectionObject;
        /* переменные для работы с таблицой */
        private dynamic WordTable;
        private dynamic WordTableBehaivior;
        private dynamic WordTableAutoFitBehaivior;
        private dynamic WordTableInscription;
        private dynamic WordTableCell;
        private dynamic AnotherWordTableCell;
        private dynamic WordTableCellRange;
        // cуммарное число срок динамической таблицы 
        public static int RowsQuantity = 0;    
        /* словарь для заполнения первого (№ пункта проверки) и второго (названия методик) столбца таблицы */
        public static Dictionary<int, string[]> WordTableDictionary = new Dictionary<int, string[]>();
        /* далее идут названия всех подпунктов проверки по каждой методике для таблицы отчета */
        private static string[] ArrayOfMethod1 =
        {
            "Параметры передатчика", "- диапазон рабочих частот передатчика, МГц", "- шаг сетки частот, МГц"         
        };
        private static string[] ArrayOfMethod2 =
        {
            "Максимальная и минимальная частота сдвига, шаг частоты сдвига", "- Шаг частоты сдвига Fд, кГц",  "- Минимальная и максимальная частота сдвига"
        };
        private static string[] ArrayOfMethod3 = {"Кратковременная нестабильность частоты сигнала Foi+Fд за 20 мс"};
        private static string[] ArrayOfMethod4 = {"Время готовности модуля, сек."};
        private static string[] ArrayOfMethod5 = {"Время переключения литерных частот в пределах выбранного диапазона Foi+Fд, мс"};
        private static string[] ArrayOfMethod6 = {"Уровень импульсной мощности выходного сигнала Foi+Fд, Вт" };
        private static string[] ArrayOfMethod7 = {"Уровень сигнала Мпрд, В" };
        private static string[] ArrayOfMethod8 = {"Длительность переднего и заднего фронтов излучаемого импульса, нс" };
        private static string[] ArrayOfMethod9  = {"Длительность излучаемого импульса, нс" };
        private static string[] ArrayOfMethod10 = {"Уровень фазовых и амплитудных шумов выходного сигнала Foi+Fд, дБ/Гц, при отстройке от несущей на: 2,5 кГц  5 кГц  100 кГц" };
        private static string[] ArrayOfMethod11 = {"Уровень дискретных составляющих выходного сигнала в диапазоне отстроек от несущей, дБ/н" };
        private static string[] ArrayOfMethod12 = {"Уровень побочных спектральных составляющих в спектре выходного сигнала передатчика, дБ"," - в диапазоне рабочих частот, дБ", " - вне диапазона рабочих частот до 2•Fo" };
        private static string[] ArrayOfMethod13 = {"Глубина запирания сигнала Foi+Fд по команде амплитудной модуляции, дБ" };       
        private static string[] ArrayOfMethod14 = {"Количество частотных литер, ед." };
        private static string[] ArrayOfMethod15 = {"Рабочий диапазон приемника" };
        private static string[] ArrayOfMethod16 = {"Глубина регулирования АРУ, уровень сигнала АРУтм в установившемся режиме кольца", " - глубина регулирования АРУ, дБ ",
                                                    " - Уровень сигнала АРУтм в установившемся режиме кольца" };
        private static string[] ArrayOfMethod17 = {"Глубина регулирования ступеней ВАРУ, дБ" };
        private static string[] ArrayOfMethod18 = {"Избирательность приемника, дБ", " - Избирательность по побочным каналам", " - Избирательность по зеркальным каналам" };
        private static string[] ArrayOfMethod19 = {"Полоса пропускания, КГц" };
        private static string[] ArrayOfMethod20 = {"Чувствительность приемника, измеренная по выходам Димп, Дпчк, Дз.имп; дБ" };
        private static string[] ArrayOfMethod21 = {"Уровень видеоимпульсов выходного сигнала, В" };
        private static string[] ArrayOfMethod22 = {"Величина задержки видеосигнала, мкс" };
        private static string[] ArrayOfMethod23 = {"Допустимые разбросы чувствительности приемника, дБ", " - допустимый разброс чувствительности приемника, измеренный по выходам Димп, Дзимп; дБ",
                                                    " - допустимый разброс чувствительности приемника, измеренный по выходам Дпчк, дБ"};
        private static string[] ArrayOfMethod24 = {"Уровень сигнала Апчк в установившемся режиме кольца АРУ, В" };
        private static string[] ArrayOfMethod25 = {"Нестабильность частоты выходного сигнала Foi+Fд" };
        private static string[] ArrayOfMethod26 = { "Ток потребления, A" };
        private static string[] ArrayOfMethod27 = { "Временные параметры выходного сигнала при частоте повторения 80 - 250 кГц и скважности 2.5 - 5",
                                                    " - длительность фронтов импульса, нс не более",
                                                    " - длительность СВЧ импульсов (минимальная), нс не менее",
                                                    " - длительность СВЧ импульсов (максимальная), нс не более"};


        private MicrosoftWordReport()
        {
            t = Type.GetTypeFromProgID("Word.Application");
            WordApplication = Activator.CreateInstance(t);
            WordDocument = WordApplication.Documents.Add();      
            // задание параметров шрифта
            SelectionObject = WordApplication.Selection;
            SelectionObject.Font.Name = "Times New Roman";
            SelectionObject.Font.Size = 10;
            SelectionObject.Font.Color = Word.WdColor.wdColorBlack;
            // межстрочный интервал в минимальный и убрать интервал после абзаца
            SelectionObject.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            SelectionObject.ParagraphFormat.SpaceAfter = 0.0f;
            // установка полей документа в минимум
            WordApplication.ActiveDocument.PageSetup.TopMargin = WordApplication.CentimetersToPoints(1.27);
            WordApplication.ActiveDocument.PageSetup.LeftMargin = WordApplication.CentimetersToPoints(1.27);
            WordApplication.ActiveDocument.PageSetup.RightMargin = WordApplication.CentimetersToPoints(1.27);
            WordApplication.ActiveDocument.PageSetup.BottomMargin = WordApplication.CentimetersToPoints(1.27);
            // выравнивание заголовка отчета по центру
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            SelectionObject.TypeText("Приложение к протоколу проверки электрических параметров\nмодуля ППМ РО-2Д зав. № ");  
        }

        public MicrosoftWordReport(string DeviceID, string CategoryInspection, string KindOfInspection, string SubKindOfInsp) : this()
        {
            if (string.IsNullOrEmpty(DeviceID)) DeviceID                     = "";
            if (string.IsNullOrEmpty(CategoryInspection)) CategoryInspection = "";
            if (string.IsNullOrEmpty(KindOfInspection)) KindOfInspection     = "";
            SelectionObject.TypeText(DeviceID);
            // Новый параграф и выравнивание по левому краю
            SelectionObject.TypeParagraph();         
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            SelectionObject.TypeText("\n\tКатегория испытаний: ");
            SelectionObject.TypeText(CategoryInspection);
            SelectionObject.TypeText("\n\tВид испытания: ");
            SelectionObject.TypeText(KindOfInspection);
            SelectionObject.TypeText("\n\tПодвид испытания: ");
            SelectionObject.TypeText(SubKindOfInsp);
            // Новый параграф и установка даты проверки
            SelectionObject.TypeParagraph();
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            SelectionObject.TypeText("\nДата: "+ DateTime.Now.ToString());
            SelectionObject.TypeText("\n");          
        }
        /* вставка динамической таблицы */
        public void InsertTable() => InsertTableHeader(RowsQuantity); // вставить шапку таблицы
 
        /* функция вставки окончания отчета и сохранения в формат pdf */
        public void InsertEnding(bool SpokesMan)
        {
            // AwaitingForm af = new AwaitingForm();
            // Выход из таблицы, перемещение вниз по тексту, новый параграф и выравнивание по левому краю
         
            SelectionObject.EndKey(Word.WdUnits.wdLine);
            SelectionObject.MoveDown(Word.WdUnits.wdScreen, 1);
            SelectionObject.TypeParagraph();
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            SelectionObject.TypeText("\nПредставитель ОТК\t\t\t\tПредставитель изготовителя ");
            if (SpokesMan) SelectionObject.TypeText("\n\n\nПредставитель ВП МО РФ");
            if (FormProtocol.DeviceId != string.Empty)  // Если был задан номер изделия, то сохранить в каталоге под этим именем, в противном случае сохранить в папке по умолчанию
            {
                if (!Directory.Exists(Application.StartupPath + @"\Отчёты проверок")) Directory.CreateDirectory(Application.StartupPath + @"\Отчёты проверок");
                try
                {
                    WordDocument.SaveAs(Application.StartupPath + @"\Отчёты проверок\№ Изд." + FormProtocol.DeviceId + ".pdf", Word.WdSaveFormat.wdFormatPDF);
                    WordDocument.SaveAs(Application.StartupPath + @"\Отчёты проверок\Report Preview\Preview", Word.WdSaveFormat.wdFormatHTML);
                }
                catch
                {
                 //   MessageBox.Show("Ошибка сохранения отчёта: закройте уже открытый отчёт", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Marshal.ReleaseComObject(WordApplication);
                    Dispose();
                }
            }
            else
            {
                if (!Directory.Exists(Application.StartupPath + @"\Отчёты проверок\DefaultReport")) Directory.CreateDirectory(Application.StartupPath + @"\Отчёты проверок\DefaultReport");
                try
                {
                    WordDocument.SaveAs(Application.StartupPath + @"\Отчёты проверок\DefaultReport\Результаты проверки РО-2Д на требования ТУ", Word.WdSaveFormat.wdFormatPDF);
                    WordDocument.SaveAs(Application.StartupPath + @"\Отчёты проверок\Report Preview\Preview", Word.WdSaveFormat.wdFormatHTML);
                }
                catch
                {
                  //  MessageBox.Show("Ошибка сохранения отчёта: закройте уже открытый отчёт", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    WordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                   // Marshal.ReleaseComObject(WordApplication);
                    Dispose();
                }
            }
           if(WordApplication!=null) WordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }


        /* шапка таблицы отчета по методикам ТУ */
        private void InsertTableHeader(int RowsQuantity)
        {
            // Новый параграф и выравнивание по центру
            SelectionObject.TypeParagraph();
            SelectionObject.TypeText("\n");
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // настройка станадартного поведения таблицы
            WordTableBehaivior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            WordTableAutoFitBehaivior = Word.WdAutoFitBehavior.wdAutoFitContent;
            // позиционирование и добавление таблицы на страницу
            SelectionObject.EndKey(Word.WdUnits.wdLine);
            WordTable = WordDocument.Tables.Add(SelectionObject.Range, 2 + RowsQuantity, 7, WordTableBehaivior, WordTableAutoFitBehaivior);
            /* Настройка предпочитаемой ширины столбцов таблицы */
            WordApplication.ActiveDocument.Tables[1].Columns[1].PreferredWidth = 8;     // ширина столбца №
            //WordApplication.ActiveDocument.Tables[1].Columns[2].PreferredWidth = 50;   // ширина столбца названипя методик ТУ
            //WordApplication.ActiveDocument.Tables[1].Columns[3].PreferredWidth = 50;
            //WordApplication.ActiveDocument.Tables[1].Columns[4].PreferredWidth = 50;
            //WordApplication.ActiveDocument.Tables[1].Columns[5].PreferredWidth = 50;
            //WordApplication.ActiveDocument.Tables[1].Columns[6].PreferredWidth = 20;
            // оформление и заполнение заголовочной строки таблицы
            /* Объединение ячеек таблицы */
            WordTableCell = WordTable.Cell(1, 5).Range.Start;
            AnotherWordTableCell = WordTable.Cell(1, 6).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();

            WordTableCell = WordTable.Cell(1, 1).Range.Start;
            AnotherWordTableCell = WordTable.Cell(2, 1).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();

            WordTableCell = WordTable.Cell(1, 2).Range.Start;
            AnotherWordTableCell = WordTable.Cell(2, 2).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();

            WordTableCell = WordTable.Cell(1, 3).Range.Start;
            AnotherWordTableCell = WordTable.Cell(2, 3).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();

            WordTableCell = WordTable.Cell(1, 4).Range.Start;
            AnotherWordTableCell = WordTable.Cell(2, 4).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();

            WordTableCell = WordTable.Cell(1, 6).Range.Start;
            AnotherWordTableCell = WordTable.Cell(2, 7).Range.End;
            WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
            WordTableCellRange.Select();
            WordApplication.Selection.Cells.Merge();
            // наполнение ячеек надписями с выравниванием по центру (заголовки)
            WordTableInscription = WordDocument.Tables[1].Cell(1, 1).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "№ методики";

            WordTableInscription = WordDocument.Tables[1].Cell(1, 2).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "Параметр";

            WordTableInscription = WordDocument.Tables[1].Cell(1, 3).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "Величина\nпо ТУ";

            WordTableInscription = WordDocument.Tables[1].Cell(1, 4).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "Допуск";

            WordTableInscription = WordDocument.Tables[1].Cell(1, 5).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "Величина измеренная";

            WordTableInscription = WordDocument.Tables[1].Cell(1, 6).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "Соответствие / Несоответствие";

            WordTableInscription = WordDocument.Tables[1].Cell(2, 5).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "ПС1";

            WordTableInscription = WordDocument.Tables[1].Cell(2, 6).Range;
            WordTableInscription.Font.Size = 9;
            WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordTableInscription.Text = "ПС2";
            /* вертикальное выравнивание ячеек таблицы так же по центру */
            WordTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

        }

        /* наполнение таблицы результатов данными */
        public void FillWordTable(List<DataTable> Tables)
        {
            int TUPoint = 0;
            int rows = 0;
            if (RowsQuantity == 0) return;
            byte SmallColumn1Counter = 0; // приращение по ячейкам таблицы
            /* наполнение записями 1 и второго столбца таблицы отчета (№ п.ТУ и наименование методики соотвестственно) */
            foreach (KeyValuePair<int, string[]> KeyVal in WordTableDictionary)
            {
                WordTableInscription = WordDocument.Tables[1].Cell(3 + SmallColumn1Counter, 1).Range;
                WordTableInscription.Font.Size = 9;
                TUPoint = 1 + KeyVal.Key;
                WordTableInscription.Text = TUPoint;  // вывод номера методики ТУ в первый столбец таблицы отчета 
                                                                   //    
                XmlStrings.GetAttributeValFromXml(KeyVal.Key, out XmlStrings.Column3, out XmlStrings.Column4); // получение допусков и значений согласно ТУ
                /* вывод в таблицу данных для второго столбца */
                for (byte x = 0; x < KeyVal.Value.Length; x++) // проход по строкам с названиями проверок
                {
                    WordTableInscription = WordDocument.Tables[1].Cell(3 + SmallColumn1Counter + x, 2).Range;
                    WordTableInscription.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    WordTableInscription.Font.Bold = (x == 0) ? 1 : 0;  // выделение названия проверки ТУ полужирным шрифтом
                    WordTableInscription.Font.Size = 9;
                    WordTableInscription.Text = KeyVal.Value.GetValue(x);

                    // Заполнение 3 столца таблицы отчета (Величина по ТУ)
                    WordTableInscription = WordDocument.Tables[1].Cell(3 + SmallColumn1Counter + x, 3).Range;
                    WordTableInscription.Font.Size = 9;
                    WordTableInscription.Text = (XmlStrings.Column3[x] == "null") ? " " : XmlStrings.Column3[x].Replace(' ','\n');
                    // Заполнение 4 столца таблицы отчета (Допуск)
                    WordTableInscription = WordDocument.Tables[1].Cell(3 + SmallColumn1Counter + x, 4).Range;
                    WordTableInscription.Font.Size = 9;
                    WordTableInscription.Text = (XmlStrings.Column4[x] == "null") ? " " : XmlStrings.Column4[x].Replace(' ', '\n');
                    /* обработка частного случая, возведения числа в степень */
                    if (XmlStrings.Column3[x].Length == 8 && XmlStrings.Column3[x].Substring(5) == "E-6")
                    {
                        WordDocument.Tables[1].Cell(3 + SmallColumn1Counter + x, 3).Range.Select();
                        SelectionObject.TypeText("±5•10");
                        SelectionObject.Font.Superscript = Word.WdConstants.wdToggle;
                        SelectionObject.TypeText("-6");
                    }
                    if (XmlStrings.Column3[x].Length == 8 && XmlStrings.Column3[x].Substring(5) == "E-8")
                    {
                        WordDocument.Tables[1].Cell(3 + SmallColumn1Counter + x, 3).Range.Select();
                        SelectionObject.TypeText("±1•10");
                        SelectionObject.Font.Superscript = Word.WdConstants.wdToggle;
                        SelectionObject.TypeText("-8");
                        
                    }

                }

                // объединение ячеек со строками из первого столбца таблицы(№ пунктов) когда это необходимо (если проверка состоит из нескольких подпунктов)
                if (KeyVal.Value.Length > 1) 
                {
                    WordTableCell = WordTable.Cell(4 + SmallColumn1Counter, 1).Range.Start;
                    SmallColumn1Counter += (byte)KeyVal.Value.Length;
                    AnotherWordTableCell = WordTable.Cell(4 + SmallColumn1Counter - 2, 1).Range.End;
                    WordTableCellRange = WordDocument.Range(WordTableCell, AnotherWordTableCell);
                    WordTableCellRange.Select();
                    WordApplication.Selection.Cells.Merge();
                }
                else SmallColumn1Counter += (byte)KeyVal.Value.Length;

            }
            // Вывод результатов проверки по ТУ (5, 6, 7 столбцы в таблице)
            foreach(DataTable Table in Tables)
            {
                for(int r = 0; r < Table.Rows.Count; r++)
                {                       
                    for (int c = 0; c < Table.Columns.Count-2; c++)
                    {                     
                        WordTableInscription.Font.Size = 9;
                        WordTableInscription = WordDocument.Tables[1].Cell(3 + rows, 5 + c).Range;
                        if (Table.Rows[r][c].ToString().Contains("E"))
                        {
                           
                            string[] digits = Table.Rows[r][c].ToString().Split('E');
                            WordDocument.Tables[1].Cell(3 + rows, 5 + c).Range.Select();
                            SelectionObject.TypeText(digits[0]+ "•10");
                            SelectionObject.Font.Superscript = Word.WdConstants.wdToggle;
                            string re = digits[1].Replace('0',' ').Remove(1,2);
                            SelectionObject.TypeText(re);
                        }
                        else WordTableInscription.Text = Table.Rows[r][c];
                    }
                    rows++;
                }
               
            }
            //XmlStrings.Column3 = null;
            //XmlStrings.Column4 = null;
        }

        /* заполнение в словарь данных методики ТУ (ключ - порядковый номер методики, значение - строковый массив с пунктами проверок*/
        public static void FillLists(int MethodNumber, bool AddRemove)
        {
            RowsQuantity = 0;
            foreach (DataTable table in ExaminationResults.Tables) RowsQuantity += table.Rows.Count;
            if (AddRemove)
            {
                if (WordTableDictionary.ContainsKey(MethodNumber)) return;
                switch (MethodNumber )
                {
                    case 0: WordTableDictionary.Add(MethodNumber, ArrayOfMethod1); break;
                    case 1: WordTableDictionary.Add(MethodNumber, ArrayOfMethod2); break;
                    case 2: WordTableDictionary.Add(MethodNumber, ArrayOfMethod3); break;
                    case 3: WordTableDictionary.Add(MethodNumber, ArrayOfMethod4); break;
                    case 4: WordTableDictionary.Add(MethodNumber, ArrayOfMethod5); break;
                    case 5: WordTableDictionary.Add(MethodNumber, ArrayOfMethod6); break;
                    case 6: WordTableDictionary.Add(MethodNumber, ArrayOfMethod7); break;
                    case 7: WordTableDictionary.Add(MethodNumber, ArrayOfMethod8); break;
                    case 8: WordTableDictionary.Add(MethodNumber, ArrayOfMethod9); break;
                    case 9: WordTableDictionary.Add(MethodNumber, ArrayOfMethod10); break;
                    case 10: WordTableDictionary.Add(MethodNumber, ArrayOfMethod11); break;
                    case 11: WordTableDictionary.Add(MethodNumber, ArrayOfMethod12); break;
                    case 12: WordTableDictionary.Add(MethodNumber, ArrayOfMethod13); break;
                    case 13: WordTableDictionary.Add(MethodNumber, ArrayOfMethod14); break;
                    case 14: WordTableDictionary.Add(MethodNumber, ArrayOfMethod15); break;
                    case 15: WordTableDictionary.Add(MethodNumber, ArrayOfMethod16); break;
                    case 16: WordTableDictionary.Add(MethodNumber, ArrayOfMethod17); break;
                    case 17: WordTableDictionary.Add(MethodNumber, ArrayOfMethod18); break;
                    case 18: WordTableDictionary.Add(MethodNumber, ArrayOfMethod19); break;
                    case 19: WordTableDictionary.Add(MethodNumber, ArrayOfMethod20); break;
                    case 20: WordTableDictionary.Add(MethodNumber, ArrayOfMethod21); break;
                    case 21: WordTableDictionary.Add(MethodNumber, ArrayOfMethod22); break;
                    case 22: WordTableDictionary.Add(MethodNumber, ArrayOfMethod23); break;
                    case 23: WordTableDictionary.Add(MethodNumber, ArrayOfMethod24); break;
                    case 24: WordTableDictionary.Add(MethodNumber, ArrayOfMethod25); break;
                    case 25: WordTableDictionary.Add(MethodNumber, ArrayOfMethod26); break;
                    case 26: WordTableDictionary.Add(MethodNumber, ArrayOfMethod27); break;
                }
            }
            else WordTableDictionary.Remove(MethodNumber);       
               
            }
                    
        /* открытие диалогового окна сохранения отчета */
        //private void SaveDialog()
        //{
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Title = "Сохранение нового отчёта";
        //    sfd.Filter = "Файл PDF ( *.pdf) | *pdf";
        //    sfd.InitialDirectory = Application.StartupPath + @"\Отчёты проверок";
        //    sfd.FileName = "№ Изд." + FormProtocol.DeviceId;
        //    if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        AwaitingForm af = new AwaitingForm();
        //        af.Show();
        //        WordDocument.SaveAs(sfd.FileName + ".pdf", Word.WdSaveFormat.wdFormatPDF);
        //        af.CompleteSaving();
        //        WordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        //     //   System.Diagnostics.Process.Start(sfd.FileName + ".pdf");
        //    }
        //}

        public void SaveReportAs(int Format, bool SpokesMan, string SaveDirectory)
        {

            AwaitingForm af = new AwaitingForm();
            // Выход из таблицы, перемещение вниз по тексту, новый параграф и выравнивание по левому краю
            SelectionObject.EndKey(Word.WdUnits.wdLine);
            SelectionObject.MoveDown(Word.WdUnits.wdScreen, 1);
            SelectionObject.TypeParagraph();
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            SelectionObject.TypeText("\nПредставитель ОТК\t\t\t\tПредставитель изготовителя ");
            if (SpokesMan) SelectionObject.TypeText("\n\n\nПредставитель ВП МО РФ"); 
            af.Show();
            try
              {
                switch (Format)
                {
                    case 1: WordDocument.SaveAs(SaveDirectory + ".pdf", Word.WdSaveFormat.wdFormatPDF);
                        break;
                    case 2: WordDocument.SaveAs(SaveDirectory + ".docx", Word.WdSaveFormat.wdFormatDocumentDefault);
                        break;
                    case 3: WordDocument.SaveAs(SaveDirectory + ".xps", Word.WdSaveFormat.wdFormatXPS);
                        break;
                    case 4: WordDocument.SaveAs(SaveDirectory + ".rtf", Word.WdSaveFormat.wdFormatRTF);
                        break;
                    case 5: //WordDocument.SaveAs(SaveDirectory + ".odt", Word.WdSaveFormat.wdFormatOpenDocumentText);
                        break;
                    case 6: WordDocument.SaveAs(SaveDirectory + ".html", Word.WdSaveFormat.wdFormatHTML);
                        break;
                }
                af.CompleteSaving();
               }
               catch
               {
                 MessageBox.Show("Ошибка сохранения отчёта: закройте уже открытый отчёт", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 Marshal.ReleaseComObject(WordApplication);
                 Dispose();
                 af.Close();
                 return;
                }
            WordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
                     
        /* автосохренние отчета на диск каждый раз при прохождении какого-либо пождпункта/пункта проверки */
        public static void AutoSaveReportAsync()
        {  
            Task.Factory.StartNew(()=> 
            {
                using (MicrosoftWordReport mswr = new MicrosoftWordReport(FormProtocol.DeviceId, FormProtocol.CategoryInspection, FormProtocol.KindOfInspection, FormProtocol.SubkindOfInspection))
                {
                    mswr.InsertTable();
                    mswr.FillWordTable(ExaminationResults.Tables);
                    mswr.InsertEnding(FormProtocol.SpokesManFlag);
                }
            });
        }

        /* построение preview-файла при открытии ранее сохраненного отчета */
        public void BuildPreviewFile(bool SpokesMan)
        {
            SelectionObject.EndKey(Word.WdUnits.wdLine);
            SelectionObject.MoveDown(Word.WdUnits.wdScreen, 1);
            SelectionObject.TypeParagraph();
            SelectionObject.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            SelectionObject.TypeText("\nПредставитель ОТК\t\t\t\tПредставитель изготовителя ");
            if (SpokesMan) SelectionObject.TypeText("\n\n\nПредставитель ВП МО РФ");
            if (!Directory.Exists(Application.StartupPath + @"\Отчёты проверок\Report Preview")) Directory.CreateDirectory(Application.StartupPath + @"\Отчёты проверок\Report Preview");
            WordDocument.SaveAs(Application.StartupPath + @"\Отчёты проверок\Report Preview\Preview", Word.WdSaveFormat.wdFormatHTML);
            WordApplication.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
          //  Marshal.ReleaseComObject(WordApplication);
            Dispose();
       
        }

        /* освобождение ресурсов */
        public void Dispose()
        {
           // Marshal.ReleaseComObject(WordApplication);
            WordApplication = null;
            WordDocument = null;
        }
    }
}
