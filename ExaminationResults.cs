using System;
using System.Collections.Generic;
using System.Data;
/* 
Класс, собирающий информацию о пройденных методиках
Автор: Евгений Якайтис
отдел: НИО-5 
*/
namespace RO2D
{
    partial class ExaminationResults : Method
    {
        static DataTable ResultsTable = new DataTable("Results");
       
        static ExaminationResults()
        {       
         
           for(int c=0; c < 27; c++) InitTable(c); // инициализация пустых таблиц
        }

        public static bool successRD1 = false;
        public static bool successRD2 = false;
 
        /* добавление или удаление таблиц с результатами */
        public static void AddRemoveTable(int SelectedMethod, bool AddOrRemove)
        {
            SelectedMethod++;    
            if (AddOrRemove)
            {
                if (Tables.Exists(match => match.TableName == "Method" + SelectedMethod)) return;             
                switch (SelectedMethod)
                {
                    case 1: Tables.Add(Table1);
                         break;
                    case 2: Tables.Add(Table2);
                        break;
                    case 3: Tables.Add(Table3);
                        break;
                    case 4: Tables.Add(Table4);
                        break;
                    case 5: Tables.Add(Table5);
                        break;
                    case 6: Tables.Add(Table6);
                        break;
                    case 7: Tables.Add(Table7);
                        break;
                    case 8: Tables.Add(Table8);
                        break;
                    case 9: Tables.Add(Table9);
                        break;
                    case 10: Tables.Add(Table10);
                        break;
                    case 11: Tables.Add(Table11);
                        break;
                    case 12: Tables.Add(Table12);
                        break;
                    case 13: Tables.Add(Table13);
                        break;
                    case 14: Tables.Add(Table14);
                        break;
                    case 15: Tables.Add(Table15);
                        break;
                    case 16: Tables.Add(Table16);
                        break;
                    case 17: Tables.Add(Table17);
                        break;
                    case 18: Tables.Add(Table18);
                        break;
                    case 19: Tables.Add(Table19);
                        break;
                    case 20: Tables.Add(Table20);
                        break;
                    case 21: Tables.Add(Table21);
                        break;
                    case 22: Tables.Add(Table22);
                        break;
                    case 23: Tables.Add(Table23);
                        break;
                    case 24: Tables.Add(Table24);
                        break;
                    case 25: Tables.Add(Table25);
                        break;
                    case 26: Tables.Add(Table26);
                        break;
                    case 27:
                        Tables.Add(Table27);
                        break;
                }
            }
            else Tables.RemoveAll((table) => table.TableName.Contains(SelectedMethod.ToString()));
       }   
    }  
 }


