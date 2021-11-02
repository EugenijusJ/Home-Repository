using System;
using RO2D.XmlText; // открытие видимости для листинга "XmlStrings.cs в папке "XmlText"
using System.Drawing;
using System.Windows.Forms;
using Core.Devices;  // пространство имен для работы с генератором SMB100A
using RO_2D;
/*
*   Автор: Евгений Якайтис
*   отдел: НИО-5
*/
namespace RO2D
{
    partial class Method
    {
        /* Глубина регулирования ступеней ВАРУ */
        static void MethodTU17(ref sbyte CurrentMethidicStage)
        {
            CURRENT = CURRENT_METHODIC.M17;
            QuestionForm qf;
            switch (CurrentMethidicStage)
            {
                case 0:
                    FormMain.Print(MethodicArrayList[16], HorizontalAlignment.Center, Color.Black, FontStyle.Bold);
                    FormMain.Print(RO_2D.Properties.Resources.Preface, HorizontalAlignment.Center, Color.Black, FontStyle.Bold);
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P1"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P2"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P3"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 1:

                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P4"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 2:
                    // далее установить следующие значения параметров блока БКУ
                    BKU.gs_BKU.Tx.Rd = 1;    // номер диапазона
                    BKU.gs_BKU.Tx.Nlit = 2;    // номер литеры
                    BKU.gs_BKU.Tx.Fg = 0;    // смещение Доплера
                    BKU.gs_BKU.Tx.ImpSTR = 1;    // строб ИСТР
                    BKU.gs_BKU.Tx.ImpBLN = 0;    // строб ИБЛН
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P6"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 3:
                    if (BAND == BAND_NUMBER.RD1) FormMain.Print("Проверка по диапазону ПС1", HorizontalAlignment.Center, Color.Black, FontStyle.Bold | FontStyle.Underline, false);
                    else FormMain.Print("Проверка по диапазону ПС2", HorizontalAlignment.Center, Color.Black, FontStyle.Bold | FontStyle.Underline, false);
                    if (bn != null && bn != BAND) BandWasChanged = true;   // если был выбран другой диапазон ПС, то выставить флаг в true
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P7"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    CurrentMethidicStage++;
                    return;
                case 4:
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P8"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    BPRO.RGU.gs_RGU.Tx.Режим = 1;   // непрерывный режим работы БПРО
                    BPRO.RGU.gs_RGU.Tx.Цикл = 0;    // обязательная манипуляция
                    SMB100A.PulseModulationOn(); // включение модуляции               
                    CurrentMethidicStage++;
                    return;
                case 5:
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P9"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.PlusButtonDisable();
                    FormMain.MakeDecision();     // принятие решения по п.9 методики ТУ
                    CurrentMethidicStage++;
                    return;
                case 6:
                    if (CHECKPOINT_DECISION == DECISION.YES)
                    {
                        FormMain.Print("\n       Значение литерной частоты в выбранном диапазоне соотвествует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        CurrentMethidicStage++;
                        MethodTU17(ref CurrentMethidicStage);
                        return;
                    }
                    else if (CHECKPOINT_DECISION == DECISION.NO)
                    {
                        FormMain.Print("\n       Значение литерной частоты в выбранном диапазоне не соотвествует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        CurrentMethidicStage++;
                        MethodTU17(ref CurrentMethidicStage);
                        return;
                    }
                    else if (CHECKPOINT_DECISION == DECISION.UNKNOWN)
                    {
                        FormMain.MakeDecision();
                        return;
                    }
                    return;
                case 7:
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P10"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    CurrentMethidicStage++;
                    return;
                case 8:
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P11"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    BKU.gs_BKU.Tx.Vzm = 0;  // ВЗМ (взаимодействие выключить)
                    CurrentMethidicStage++;
                    return;
                //CurrentMethidicStage++;
                //if (CycleQuantity%2 == 0)
                //{                   
                //    return;
                //}
                //else
                //{
                //    MethodTU11(ref CurrentMethidicStage);
                //    return;
                //}
                case 9:
                    ExaminationResults.Table17.Rows[0].BeginEdit();
                    qf = new QuestionForm("Запиши значения подавления сигнала при изменении кода ВАРУ в Дб", "", ThreeTxtFileds: false);
                    qf.ShowDialog();
                    if (qf.UserDecicsionOk)
                    {
                        FormMain.Print("     Ввод оператора:\t Δ0 →  " + qf.UserEnteredValue + " дБ;\t Δ1 →  " + qf.UserEnteredValue1 + " дБ;\t Δ2 →  " + qf.UserEnteredValue2 + " дБ;\t Δ3 →  "
                                       + qf.UserEnteredValue3 + " дБ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        if (qf.UserEnteredValue >= 1 && qf.UserEnteredValue <= 3 && qf.UserEnteredValue1 >= 3 && qf.UserEnteredValue1 <= 5 &&
                            qf.UserEnteredValue2 >= 7 && qf.UserEnteredValue2 <= 9 && qf.UserEnteredValue3 >= 15 && qf.UserEnteredValue3 <= 18)
                           {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table17.Rows[0][3] = true;
                            else ExaminationResults.Table17.Rows[0][4] = true;
                            FormMain.Print("\n       Глубина регулирования ВАРУ в выбранном диапазоне ПС на Fл соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);

                        }
                        else
                        {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table17.Rows[0][3] = false;
                            else ExaminationResults.Table17.Rows[0][4] = false;
                            FormMain.Print("\n       Глубина регулирования ВАРУ в выбранном диапазоне ПС на Fл не соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);

                        }
                        if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table17.Rows[0][0] = $"{qf.UserEnteredValue}\n{qf.UserEnteredValue1}\n{qf.UserEnteredValue2}\n{qf.UserEnteredValue3}";
                        else ExaminationResults.Table17.Rows[0][1] = $"{qf.UserEnteredValue}\n{qf.UserEnteredValue1}\n{qf.UserEnteredValue2}\n{qf.UserEnteredValue3}";
                        if ((bool)ExaminationResults.Table17.Rows[0][3]  || (bool)ExaminationResults.Table17.Rows[0][4])
                        {
                            ExaminationResults.Table17.Rows[0][2] = "Соответствует";  
                        }
                        else ExaminationResults.Table17.Rows[0][2] = "Не соответствует";
                        ExaminationResults.Table17.Rows[0].EndEdit();
                        MicrosoftWordReport.AutoSaveReportAsync();
                        CurrentMethidicStage++;
                        MethodTU17(ref CurrentMethidicStage);
                        return;
                    }
                    else
                    {
                        FormMain.Print("\nМетодика завершена, т.к. оператор отменил действие.", HorizontalAlignment.Center, Color.Red, FontStyle.Regular);
                        CurrentMethidicStage = -1;
                        MethodTU11(ref CurrentMethidicStage);
                        return;
                    }
                case 10:
                    FormMain.Print(XmlStrings.GetStringFromXml("M17", "P12"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.MakeDecision();
                    CurrentMethidicStage++;
                    return;
                case 11:
 
                    if (GoAhead)
                    {
                        CycleQuantity++;
                        GoAhead = false;
                        bn = BAND;                   // запоминаем текущий выбранный диапазон    
                        CurrentMethidicStage = 2;    // если оператор нажал на плюсик, то перейти на п.7 методики ТУ и выбор другого диапазона ПС      
                        MethodTU17(ref CurrentMethidicStage);
                        return;
                    }
                    else
                    {
                        if (CHECKPOINT_DECISION == DECISION.ESC)
                        {
                            bn = null;
                            if((bool)ExaminationResults.Table17.Rows[0][3] || (bool)ExaminationResults.Table17.Rows[0][4])
                            {
                                ExaminationResults.successRD1 = true;
                                ExaminationResults.successRD2 = true;
                            }
                            if (BandWasChanged)  mc(null, new MethodCompleteEventArgs(17, true, BAND));  // если уже проверены оба диапазона ПС, то вызвать событие завершения 
                            else mc(null, new MethodCompleteEventArgs(17, false, BAND)); // в противном случае инициировать событие только для одного диапазона ПС
                            FormMain.Print("\nМетодика завершена\n", HorizontalAlignment.Center, Color.Black, FontStyle.Bold);
                            CurrentMethodicStage = -1;
                            GoAhead = false;
                            ResetCounters();
                            return;
                        }
                        else if (CHECKPOINT_DECISION == DECISION.UNKNOWN)
                        {
                            FormMain.MakeDecision(); // если клавиша ошибочная, то повторно принять решение по п.25 методики ТУ
                            return;
                        }
                    }
                    return;
            }
        }

        /* Избирательность приемника */
        static void MethodTU18(ref sbyte CurrentMethidicStage)
        {
            CURRENT = CURRENT_METHODIC.M18;
            QuestionForm qf;
            switch (CurrentMethidicStage)
            {
                case 0:
                    FormMain.Print(MethodicArrayList[17], HorizontalAlignment.Center, Color.Black, FontStyle.Bold);
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P1"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P2"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P3"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 1:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P4"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 2:
                    // далее установить следующие значения параметров блока БКУ
                    BKU.gs_BKU.Tx.Rd = 1;        // номер диапазона
                    BKU.gs_BKU.Tx.Nlit = 2;      // номер литеры
                    BKU.gs_BKU.Tx.Fg = 0;        // смещение Доплера
                    BKU.gs_BKU.Tx.ImpSTR = 1;    // строб ИСТР
                    BKU.gs_BKU.Tx.ImpBLN = 0;    // строб ИБЛН
                    BKU.gs_BKU.Tx.ServiceMode = 5;//0x101
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P6"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular, true);
                    CurrentMethidicStage++;
                    return;
                case 3:
                    if (BAND == BAND_NUMBER.RD1) FormMain.Print("Проверка по диапазону ПС1", HorizontalAlignment.Center, Color.Black, FontStyle.Bold | FontStyle.Underline, false);
                    else FormMain.Print("Проверка по диапазону ПС2", HorizontalAlignment.Center, Color.Black, FontStyle.Bold | FontStyle.Underline, false);
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P7"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    if (bn != null && bn != BAND) BandWasChanged = true;   // если был выбран другой диапазон ПС, то выставить флаг в true
                    CurrentMethidicStage++;
                    return;
                case 4:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P8"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    BPRO.RGU.gs_RGU.Tx.Режим = 1;   // непрерывный режим работы БПРО
                    BPRO.RGU.gs_RGU.Tx.Цикл = 0;    // обязательная манипуляция
                    SMB100A.PulseModulationOn(); // включение импульсной модуляции 
                    CurrentMethidicStage++;
                    return;
                case 5:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P9"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    CurrentMethidicStage++;
                    return;
                case 6:
                    ExaminationResults.Table18.Rows[1].BeginEdit();
                    qf = new QuestionForm("Запиши в поле уровень подавления сигнала на побочном канале приема в Дб", "", ThreeTxtFileds: false);
                    qf.ShowDialog();
                    if (qf.UserDecicsionOk)
                    {
                        FormMain.Print("     Ввод оператора: " + qf.UserEnteredValue + " дБ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        if (Math.Abs(qf.UserEnteredValue) >= 40)
                        {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[1][3] = true;
                            else ExaminationResults.Table18.Rows[1][4] = true;
                            FormMain.Print("\n       Избирательность приемника по побочным каналам в выбранном диапазоне для выбранной частоты соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);

                        }
                        else
                        {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[1][3] = false;
                            else ExaminationResults.Table18.Rows[1][4] = false;
                            FormMain.Print("\n       Избирательность приемника по побочным каналам в выбранном диапазоне для выбранной частоты не соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);

                        }
                        if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[1][0] = Convert.ToString(Math.Abs(qf.UserEnteredValue));
                        else ExaminationResults.Table18.Rows[1][1] = Convert.ToString(Math.Abs(qf.UserEnteredValue));
                        if (BandWasChanged)
                        {
                            if ((bool)ExaminationResults.Table18.Rows[1][3] && (bool)ExaminationResults.Table18.Rows[1][4]) ExaminationResults.Table18.Rows[1][2] = "Соответствует";
                            else ExaminationResults.Table18.Rows[1][2] = "Не соответствует";
                        }
                        ExaminationResults.Table18.Rows[1].EndEdit();
                        MicrosoftWordReport.AutoSaveReportAsync();
                        CurrentMethidicStage = (SatisfactionCounter == 7) ? (sbyte)8 : (SatisfactionCounter == 8) ? (sbyte)9 : (sbyte)7;
                        MethodTU18(ref CurrentMethidicStage);
                        return;
                    }
                    else
                    {
                        FormMain.Print("\nМетодика завершена, т.к. оператор отменил действие.", HorizontalAlignment.Center, Color.Red, FontStyle.Regular);
                        CurrentMethidicStage = -1;
                        MethodTU18(ref CurrentMethidicStage);
                        return;
                    }
                case 7:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P10"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    SatisfactionCounter = 7;
                    CurrentMethidicStage = 4;
                    return;
                case 8:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P12"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    SatisfactionCounter = 8;
                    CurrentMethidicStage = 4;
                    return;
                case 9:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P14"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    CurrentMethidicStage++;
                    return;
                case 10:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P15"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.MakeDecision(); // принятие решения по п.11 методики ТУ
                    CurrentMethidicStage++;
                    return;
                case 11:
                    ExaminationResults.Table18.Rows[2].BeginEdit();
                    qf = new QuestionForm("Запиши в поле уровень подавления сигнала на анализаторе\nспектра по маркеру на зеркальном канале приема", "", ThreeTxtFileds: false);
                    qf.ShowDialog();
                    if (qf.UserDecicsionOk)
                    {
                        FormMain.Print("     Ввод оператора: " + qf.UserEnteredValue + " дБ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        if (Math.Abs(qf.UserEnteredValue) >= 30)
                        {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[2][3] = true;
                            else ExaminationResults.Table18.Rows[2][4] = true;
                            FormMain.Print("\n       Избирательность приемника по зеркальным каналам в выбранном диапазоне соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);
                        }
                        else
                        {
                            if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[2][3] = true;
                            else ExaminationResults.Table18.Rows[2][4] = true;
                            FormMain.Print("\n       Избирательность приемника по зеркальным каналам в выбранном диапазоне не соответствует ТУ", HorizontalAlignment.Left, Color.Black, FontStyle.Bold);

                        }
                        if (BAND == BAND_NUMBER.RD1) ExaminationResults.Table18.Rows[2][0] = Convert.ToString(Math.Abs(qf.UserEnteredValue));
                        else ExaminationResults.Table18.Rows[2][1] = Convert.ToString(Math.Abs(qf.UserEnteredValue));
                        if (BandWasChanged)
                        {
                            if ((bool)ExaminationResults.Table18.Rows[2][3] && (bool)ExaminationResults.Table18.Rows[2][4]) ExaminationResults.Table18.Rows[2][2] = "Соответствует";
                            else ExaminationResults.Table18.Rows[2][2] = "Не соответствует";
                        }
                        ExaminationResults.Table18.Rows[2].EndEdit();
                        MicrosoftWordReport.AutoSaveReportAsync();
                        CurrentMethidicStage++;
                        MethodTU18(ref CurrentMethidicStage);
                        return;
                    }
                    else
                    {
                        FormMain.Print("\nМетодика завершена, т.к. оператор отменил действие.", HorizontalAlignment.Center, Color.Red, FontStyle.Regular);
                        CurrentMethidicStage = -1;
                        MethodTU18(ref CurrentMethidicStage);
                        return;
                    }
                case 12:
                    FormMain.Print(XmlStrings.GetStringFromXml("M18", "P16"), HorizontalAlignment.Left, Color.Black, FontStyle.Regular);
                    FormMain.MakeDecision(); // принятие решения по п.12 методики ТУ
                    CurrentMethidicStage++;
                    return;
                case 13:
                    for(int c=1; c< ExaminationResults.Table18.Rows.Count; c++)
                    {
                        if ((bool)ExaminationResults.Table18.Rows[c][3]) ExaminationResults.successRD1 = true;
                        else
                        {
                            ExaminationResults.successRD1 = false;
                            break;
                        }                   
                        if ((bool)ExaminationResults.Table18.Rows[c][4]) ExaminationResults.successRD2 = true;
                        else
                        {
                             ExaminationResults.successRD2 = false;
                             break;
                        }
                    }
                    if (GoAhead)
                    {

                        SatisfactionCounter = 0;
                        GoAhead = false;
                        bn = BAND;                   // запоминаем текущий выбранный диапазон    
                        CurrentMethidicStage = 2;    // если оператор нажал на плюсик, то перейти на п.7 методики ТУ и выбор другого диапазона ПС      
                        MethodTU18(ref CurrentMethidicStage);
                        return;
                    }
                    else
                    {
                        if (CHECKPOINT_DECISION == DECISION.ESC)
                        {
                            bn = null;
                            for (int c = 1; c < ExaminationResults.Table18.Rows.Count; c++)
                            {
                                if (ExaminationResults.Table18.Rows[c][2].ToString() == "Соответствует")
                                {
                                    ExaminationResults.Table18.Rows[0].BeginEdit();
                                    ExaminationResults.Table18.Rows[0][2] = "Соответствует";
                                    ExaminationResults.Table18.Rows[0].EndEdit();
                                }
                                else
                                {
                                    ExaminationResults.Table18.Rows[0].BeginEdit();
                                    ExaminationResults.Table18.Rows[0][2] = "Не соответствует";
                                    ExaminationResults.Table18.Rows[0].EndEdit();
                                    break;
                                }

                            }
                            if (BandWasChanged)  mc(null, new MethodCompleteEventArgs(18, true, BAND));  // если уже проверены оба диапазона ПС, то вызвать событие завершения 
                            else mc(null, new MethodCompleteEventArgs(18, false, BAND)); // в противном случае инициировать событие только для одного диапазона ПС
                            FormMain.Print("\nМетодика завершена\n", HorizontalAlignment.Center, Color.Black, FontStyle.Bold);
                            CurrentMethodicStage = -1;
                            GoAhead = false;
                            ResetCounters();
                            return;
                        }
                        else if (CHECKPOINT_DECISION == DECISION.UNKNOWN)
                        {
                            FormMain.MakeDecision(); // если клавиша ошибочная, то повторно принять решение по п.25 методики ТУ
                            return;
                        }
                    }
                    return;
            }
        }
    }
}
