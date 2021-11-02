using System;
using Core.ExchangeWithDevices;
using RO2D.XmlText;
using System.Threading.Tasks;

namespace Core.Devices
{
    /* Управление генератором Rohde&Schwarz */
    static class SMB100A
    {
        private static EWDEthernet _ewdEthernet;
  
        static SMB100A()
        {
            string ipAddress = XmlStrings.LoadIpAddrValue();    // загрузка ip адреса из xml-файла
            int ipPort = 5025;                                  // порт генератора 
            _ewdEthernet = new EWDEthernet(ipAddress, ipPort);
        }

      
        #region Управление генератором сигналов

        #region Идентификация / Включение / Выключение / Сброс / Тест

        public static string Identification()
        {
            return Task.Factory.StartNew<string>(() => _ewdEthernet.SendRequestDataString("*IDN?;")).Result; //_ewdEthernet.SendRequestDataString("*IDN?;");
        }

        /// <summary>
        ///     Сбрасывает параметры устройства 
        /// </summary>
        public static void Reset() => _ewdEthernet.SendRequestDataString("RST;");
   
        /// <summary>
        ///     Тестирует устройство
        /// </summary>
        public static string Test() => _ewdEthernet.SendRequestDataString("*TST?;");
 
        /// <summary>
        ///     Включает модуляцию
        /// </summary>
        public static void ModulationStateOn() => _ewdEthernet.SendRequestDataString("SOUR:MOD:ALL:STAT ON;");
 
        /// <summary>
        ///     Выключает модуляцию
        /// </summary>
        public static void ModulationStateOff()=> _ewdEthernet.SendRequestDataString("SOUR:MOD:ALL:STAT OFF;");
 
        /// <summary>
        ///     Включает RF-выход
        /// </summary>
        public static void RfOutputStateOn() => _ewdEthernet.SendRequestDataString(":OUTPut:STAT ON;");
 
        /// <summary>
        ///     Выключает RF-выход
        /// </summary>
        public static void RfOutputStateOff() => _ewdEthernet.SendRequestDataString(":OUTPut:STAT OFF;");
 
        /// <summary>
        ///     Возвращает текущее состояние модуляции
        /// </summary>
        /// <returns></returns>
        public static bool GetModulationState()
        {
            if (_ewdEthernet.SendRequestDataInt(":MOD:STAT?;") == 1) return true;
            else return false;
        }

        /// <summary>
        ///     Запрашивает текущее состояние RF-выхода
        /// </summary>
        public static bool GetRfOutputState()
        {
            if (_ewdEthernet.SendRequestDataInt(":OUTPut:STAT?;") == 1) return true;
            else return false;
        }

        #endregion

        #region Установка параметров излучения

        #region General

        /// <summary>
        ///     Устанавливает значение частоты сигнала на выходе в MHZ
        /// </summary>
        /// <param name="mHz">Значение частоты</param>
        public static void SetFreqKhz(int mHz) => _ewdEthernet.SendRequestDataString(":FREQ " + Convert.ToString(mHz) + " KHZ;");
 
        /// <summary>
        ///     Устанавливает значение частоты сигнала на выходе в заданном формате:
        ///     0 - Hz, 1 - kHz, 2 - mHz, 3 - gHz
        /// </summary>
        /// <param name="freq">Значение частоты</param>
        /// <param name="type">Значение размерности</param>
        public static void SetFreq(float freq, int type)
        {
            switch (type)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString(":FREQ " + Convert.ToString(freq) + " HZ;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString(":FREQ " + Convert.ToString(freq) + " kHZ;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString(":FREQ " + Convert.ToString(freq) + " MHZ;");
                        break;
                    }
                case 3:
                    {
                        _ewdEthernet.SendRequestDataString(":FREQ " + Convert.ToString(freq) + " GHZ;");
                        break;
                    }
            }
        }

        /// <summary>
        ///     Устанавливает значение выходной мощности в DBM
        /// </summary>
        /// <param name="dBm">Значение мощности</param>
        public static void SetPowdBm(double dBm) => _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(dBm) + " DBM;");
 
        /// <summary>
        ///     Устанавливает значение выходной мощности сигнала на выходе в заданном формате:
        ///     0 - dBm, 1 - DbuV, 2 - nv, 3 - uv, 4 - mv
        /// </summary>
        /// <param name="level">Значение выходной мощности</param>
        /// <param name="type"></param>
        public static void SetPow(double level, int type)
        {
            switch (type)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " DBM;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " DBuV;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " nv;");
                        break;
                    }
                case 3:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " uv;");
                        break;
                    }
                case 4:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " mv;");
                        break;
                    }
                case 5:
                    {
                        _ewdEthernet.SendRequestDataString(":POW " + Convert.ToString(level) + " V;");
                        break;
                    }
            }
        }

        #endregion

        #region Pulse Modulation | Pulse Generator

        /// <summary>
        ///     Включает импульсную модуляцию
        /// </summary>
        public static void PulseModulationOn() => _ewdEthernet.SendRequestDataString(":PULM:STAT ON;");
 
        /// <summary>
        ///     Выключает импульсную модуляцию
        /// </summary>
        public static void PulseModulationOff() => _ewdEthernet.SendRequestDataString(":PULM:STAT OFF;");
 
        /// <summary>
        ///     Возвращает состояние импульсной модуляции 
        /// </summary>
        /// <returns>True, если импульсная модуляция включена, иначе false</returns>
        public static bool GetPulseModulationState()
        {
            if (_ewdEthernet.SendRequestDataInt(":PULM:STAT?;") == 1) return true;
            else return false;
        }

        /// <summary>
        ///     Выбирает внутренний источник для импульсной модуляции
        /// </summary>
        public static void PulseModulationSelectInternalSource() =>  _ewdEthernet.SendRequestDataDouble(":PULM:SOUR INT;");
 
        /// <summary>
        ///     Выбирает внешний источник для импульсной модуляции
        /// </summary>
        public static void PulseModulationSelectExternalSource() => _ewdEthernet.SendRequestDataString(":PULM:SOUR EXT;");
 
        /// <summary>
        ///     Выбирает полярность инверсия
        /// </summary>
        public static void PulseModulationInverseExternalSource() =>  _ewdEthernet.SendRequestDataString(":PULM:POL INV;");
 
        public static void G10K() => _ewdEthernet.SendRequestDataString("SOUR:PULM:TRIG:EXT:TMP:G10K;");
 
        /// <summary>
        ///     Устанавливает режим генератора импульсов
        /// </summary>
        public static void PulseModulationSetSingleMode() => _ewdEthernet.SendRequestDataDouble("PULM:MODE SING;");

        /// <summary>
        ///     Устанавливает значение периода генерации импульса в миллисекундах (1 kHz = 1 ms)
        /// </summary>
        /// <param name="period">значение в кГц</param>
        public static void PulseInternalPeriodMs(double period) => _ewdEthernet.SendRequestDataDouble("PULM:PER " + 1 / period + "ms;");
 
        /// <summary>
        ///     Устанавливет длительность импульса в мкс
        /// </summary>
        /// <param name="width">Значение длительности, в мкс</param>
        public static void PulseInternalPWidthUs(double width) => _ewdEthernet.SendRequestDataDouble("PULM:WIDT " + width + "us;");
 
        /// <summary>
        ///     Устанавливет длительность импульса, согласно заданным параметрам:
        ///     0 - s, 1 - ns, 2 - ms, 3 - us
        /// </summary>
        /// <param name="width">Значение длительности импульса</param>
        /// <param name="timeType">Тип размерности</param>
        public static void PulsePWidth(double width, int timeType)
        {
            switch (timeType)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString(":PULM:WIDT " + width + " s;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString(":PULM:WIDT " + width + " ns;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString(":PULM:WIDT " + width + " ms;");
                        break;
                    }
                case 3:
                    {
                        _ewdEthernet.SendRequestDataString(":PULM:WIDT " + width + " us;");
                        break;
                    }
            }
        }

        /// <summary>
        ///     Устанавливает задержку генерации импульса
        /// </summary>
        /// <param name="delay">Значение задержки импульса</param>
        public static void PulseInternalDelayUs(double delay) => _ewdEthernet.SendRequestDataDouble("PULM:DEL " + delay + "us;");
  
        /// <summary>
        ///     Устанавливет задержку между импульсами, согласно заданным параметрам:
        ///     0 - s, 1 - ns, 2 - ms, 3 - us
        /// </summary>
        /// <param name="delay">Значение задержки между импульсами</param>
        /// <param name="timeType">Тип размерности</param>
        public static void SetPulseDelay(double delay, int timeType)
        {
            switch (timeType)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:DEL " + delay + " s;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:DEL " + delay + " ns;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:DEL " + delay + " ms;");
                        break;
                    }
                case 3:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:DEL " + delay + " us;");
                        break;
                    }
            }
        }

        /// <summary>
        ///     Устанавливает частоту повторения сигнала
        /// </summary>
        /// <param name="num">Значение частоты повторяемого сигнала</param>
        private static void FreqRepeat(int num) => _ewdEthernet.SendRequestDataDouble(":PULM:TRA:REP " + Convert.ToString(num) + ";");
 
        /// <summary>
        ///     Устанавливет частоту повторения сигнала, согласно заданным параметрам:
        ///     0 - s, 1 - ns, 2 - ms, 3 - us
        /// </summary>
        /// <param name="period">Значение частоты повторения сигнала</param>
        /// <param name="timeType">Тип размерности</param>
        public static void SetPulsePeriod(double period, int timeType)
        {            //1 Килогерц равне одной мс
            switch (timeType)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:PER " + period + " s;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:PER " + period + " ns;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:PER " + period + " ms;");
                        break;
                    }
                case 3:
                    {
                        _ewdEthernet.SendRequestDataString("PULM:PER " + period + " us;");
                        break;
                    }
            }
        }

        #endregion

        #region Frequency Modulation | LFOutput

        /// <summary>
        ///     Включает частотную модуляцию
        /// </summary>
        public static void FrequencyModulationStateOn() => _ewdEthernet.SendRequestDataString(":FM:STAT ON;");
 
        /// <summary>
        ///     Выключает частотную модуляцию
        /// </summary>
        public static void FrequencyModulationStateOff() => _ewdEthernet.SendRequestDataString(":FM:STAT OFF;");
 
        /// <summary>
        ///     Возвращает текущий статус частотной модуляции 
        /// </summary>
        /// <returns>True, если импульсная модуляция включена, иначе false</returns>
        public static bool GetFrequencyModulationState()
        {
            if (_ewdEthernet.SendRequestDataInt(":FM:STAT?;") == 1) return true;
            else return false;
        }

        /// <summary>
        ///     Выбирает в качестве источника частотной модуляции LF-генератор/LFOutput
        /// </summary>
        public static void FmSelectInternalSource() =>  _ewdEthernet.SendRequestDataDouble(":FM:SOUR INT;");
 
        /// <summary>
        ///     Выбирает Normal mode в частотной модуляции
        /// </summary>
        public static void FmSelectNormalMode()
        {
            _ewdEthernet.SendRequestDataDouble(":FM:MODE NORM;");
        }

        /// <summary>
        ///     Выбирает High Deviation в частотной модуляции
        /// </summary>
        public static void FmSelectHighDeviationMode()
        {
            _ewdEthernet.SendRequestDataDouble(":FM:MODE HIDEV");
        }

        /// <summary>
        ///    Устанавливает девиацию сигнала в мГц
        /// </summary>
        /// <param name="dev">Величина девиации, задаётся в формате 5 -> 5 mhz</param>
        public static void FmDeviationMhz(int dev)
        {
            _ewdEthernet.SendRequestDataDouble(":FM:DEV " + dev + " mHz;");
        }

        /// <summary>
        ///     Устанавливет девиацию сигнала, согласно заданным параметрам:
        ///     0 - hz, kHz - mHz
        /// </summary>
        /// <param name="deviation">Значение длительности импульса</param>
        /// <param name="typeHz">Тип размерности</param>
        public static void SetFmDeviation(double deviation, int typeHz)
        {
            switch (typeHz)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataString(":FM:DEV " + deviation + " Hz;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataString(":FM:DEV " + deviation + " kHz;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataString(":FM:DEV " + deviation + " mHz;");
                        break;
                    }
            }
        }

        /// <summary>
        ///     Устанавливет девиацию сигнала, согласно заданным параметрам:
        ///     0 - hz, kHz - mHz
        /// </summary>
        /// <param name="mode">Значение длительности импульса</param>
        public static void SetFmSelectMode(int mode)
        {
            switch (mode)
            {
                case 0:
                    {
                        _ewdEthernet.SendRequestDataDouble("FM:MODE NORM;");
                        break;
                    }
                case 1:
                    {
                        _ewdEthernet.SendRequestDataDouble("FM:MODE LNO;");
                        break;
                    }
                case 2:
                    {
                        _ewdEthernet.SendRequestDataDouble("FM:MODE HDEV;");
                        break;
                    }
            }
        }

        /// <summary>
        ///     Включает LF-выход
        /// </summary>
        public static void FmlfoOn()
        {
            _ewdEthernet.SendRequestDataString(":LFO ON;");
        }

        /// <summary>
        ///     Выключает LF-выход
        /// </summary>
        public static void FmlfoOff()
        {
            _ewdEthernet.SendRequestDataString(":LFO OFF;");
        }

        /// <summary>
        ///     Устанавливает выходное значение вольтажа на LF-выходе
        /// </summary>
        /// <param name="volt">Значение вольтажа</param>
        public static void FmlfOutputVoltage(int volt)
        {
            _ewdEthernet.SendRequestDataDouble(":LFO:VOLT " + volt + " V;");
        }

        /// <summary>
        ///     Устанавливает частоту генератора для LFO1
        /// </summary>
        /// <param name="hz">Целочисленное значение Гц</param>
        public static void FmlfOutGenFreq(int hz)
        {
            _ewdEthernet.SendRequestDataDouble(":LFO1:FREQ " + hz + " hz;");
        }

        /// <summary>
        ///     Возвращает состояние LFO
        /// </summary>
        /// <returns></returns>
        public static bool FmLFOutState()
        {
            if (_ewdEthernet.SendRequestDataInt(":LFO:STAT?;") == 1)
                return true;
            return false;
        }

        /// <summary>
        ///     Включает LFO
        /// </summary>
        public static void FmlfoStateOn()
        {
            _ewdEthernet.SendRequestDataString(":LFO ON;");
        }

        /// <summary>
        ///     Выключает LFO 
        /// </summary>
        public static void SourceSignalFmlfoStateOff()
        {
            _ewdEthernet.SendRequestDataString(":LFO OFF;");
        }
        #endregion

        #region Phase Modulation

        /// <summary>
        ///     Возвращает состояние фазовой модуляции
        /// </summary>
        /// <returns></returns>
        public static bool GetPhaseModulationState()
        {
            if (_ewdEthernet.SendRequestDataInt(":PM:STAT?;") == 1)
                return true;
            return false;
        }

        /// <summary>
        ///     Выключает фазовую модуляцию
        /// </summary>
        public static void PhaseModulationStateOff()
        {
            _ewdEthernet.SendRequestDataString(":PM:STAT OFF;");
        }

        /// <summary>
        ///     Включает фазовую модуляцию
        /// </summary>
        public static void PhaseModulationStateOn()
        {
            _ewdEthernet.SendRequestDataString(":PM:STAT ON;");
        }

        #endregion

        #region Amplitude Modulation

        /// <summary>
        ///     Возвращает состояние амплитудной модуляции
        /// </summary>
        /// <returns></returns>
        public static bool GetAmplitudeModulationState()
        {
            if (_ewdEthernet.SendRequestDataInt(":AM:STAT?;") == 1)
                return true;
            return false;
        }

        /// <summary>
        ///     Включает амплитудную модуляцию
        /// </summary>
        public static void AmplitudeModulationStateOn()
        {
            _ewdEthernet.SendRequestDataString(":AM:STAT ON;");
        }

        /// <summary>
        ///     Выключает амплитудную модуляцию
        /// </summary>
        public static void AmplitudeModulationStateOff()
        {
            _ewdEthernet.SendRequestDataString(":AM:STAT OFF;");
        }

        #endregion

        #endregion

        #region Запрос и возврат параметров излучения генератора сигналов

        /// <summary>
        ///     Возвращает значение мощности (Pии)
        /// </summary>
        /// <returns></returns>
        public static double GetPowdBm() => _ewdEthernet.SendRequestDataDouble(":POW?;");

        /// <summary>
        ///     Возвращает значение частоты
        /// </summary>
        /// <returns></returns>
        public static double GetFreqMhz() => _ewdEthernet.SendRequestDataDouble(":FREQ?;");
    
        /// <summary>
        ///     Возвращает значение ширины импульса
        /// </summary>
        /// <returns></returns>
        public static double GetPulseWidth() => _ewdEthernet.SendRequestDataDouble(":PULM:WIDT?;");
 
        /// <summary>
        ///     Возвращает значение период генерации импульса
        /// </summary>
        /// <returns></returns>
        public static double GetPulsePeriod() => _ewdEthernet.SendRequestDataDouble(":PULM:PER?;");
 
        /// <summary>
        ///     Возвращает значение девиации
        /// </summary>
        /// <returns></returns>
        public static double GetFreqDeviation() =>_ewdEthernet.SendRequestDataDouble("FM:DEV?;");
 
        /// <summary>
        ///     Возвращает значение задержки между импульсами
        /// </summary>
        /// <returns></returns>
        public static double GetPulseDelay() => _ewdEthernet.SendRequestDataDouble("PULM:DEL?;");
 
        #endregion

        #endregion
 
    }
}
