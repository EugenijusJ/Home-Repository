using System;
/* 
Класс для сбора инфоромации о пользовательском событии окончания методики ТУ
Автор: Евгений Якайтис
отдел: НИО-5
*/
namespace RO2D 
{
    public class MethodCompleteEventArgs : EventArgs
    {
        public byte MethodNumber { get; } // текущая методика
        public bool AnotherBand { get; } // была ли проверка по обоим диапазонам
        internal Method.BAND_NUMBER CurrBand {get;} // текущий номер диапазона ПС

        internal MethodCompleteEventArgs(byte MethodNum, bool BandChanged, Method.BAND_NUMBER CurrentBand)
        {
            MethodNumber = MethodNum;
            AnotherBand = BandChanged;
            CurrBand = CurrentBand;
        }
    }
}
