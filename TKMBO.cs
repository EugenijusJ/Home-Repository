using System;
using _Kernel32;


namespace _KMBO
{
    class TKMBO : TKernel32//-----------------------------------------------------------------------------------------------//TKMBO
    {
        //-------------------------------------------------------------------------------------------------------------------------//private var
        const UInt32 GENERIC_READ = 0x80000000;
        const UInt32 GENERIC_WRITE = 0x80000000;
        const UInt16 OPEN_EXISTING = 3;
        const UInt16 CREATE_NEW = 1;
        const UInt16 NULL = 0;
        const UInt16 FILE_SHARE_READ = 0;
        const UInt16 FILE_SHARE_WRITE = 0;
        const UInt16 FILE_ATTRIBUTE_NORMAL = 0;
        const UInt16 FIRST_IOCTL_INDEX = 0x0800;
        const UInt16 FILE_ANY_ACCESS = 0x0000;
        const UInt32 FILE_DEVICE_KMBO = 0x00008000;
        const UInt16 METHOD_BUFFERED = 0;
        const UInt16 METHOD_IN_DIRECT = 1;
        const UInt16 METHOD_OUT_DIRECT = 2;
        const UInt16 METHOD_NEITHER = 3;
        const UInt16 REGIM_REST = 0;                                         //Режим "Покой"
        const UInt16 REGIM_CONTROLLER = 0x1;                                 //Режим "Контроллер"
        struct TSetRegim
        {
            public UInt32 Regim;
        };                                                  //Установка режима работы
        TSetRegim SetRegim;
        //-------------------------------------------------------------------------------------------------------------------------//public var
        public IntPtr Handle;
        //-------------------------------------------------------------------------------------------------------------------------//public function                                                
        unsafe public struct TWrite
        {
            public UInt16 Address;	//Адрес ОУ
            public UInt16 Count;    //Количество записываемых слов 1..64
            public UInt16 Count1;    //Количество записываемых слов 1..64
            public UInt16 Synchro;	//Если 0, то сигнал синхронизации на осциллограф не выводится
            public fixed UInt16 lpBuffer[64];	//Массив записываемых слов
        };                                       //Структура данных для записи
        unsafe public struct TRead
        {
            public UInt16 Address;	//Адрес ОУ
            public UInt16 Count;		//Количество записываемых слов 1..64
            public UInt16 Count1;		//Количество записываемых слов 1..64
            public UInt16 Synchro;	//Если 0, то сигнал синхронизации на осциллограф не выводится
            public fixed UInt16 lpBuffer[64];	//Массив записываемых слов
        };                                        //Структура данных для чтения;
        unsafe public bool Open(string FileName)
        {
            UInt32 err = 0;
            uint ret = 0;
            UInt32 lu_SetRejime = (UInt32)((FILE_DEVICE_KMBO << 16) | (FILE_ANY_ACCESS << 14) | ((FIRST_IOCTL_INDEX + 5) << 2) | (METHOD_BUFFERED));
            Handle = CreateFile(FileName, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
            if (Handle != IntPtr.Zero)
            {
                SetRegim.Regim = 1;//KMBO_SET_REGIM_CONTROLLER;  
                fixed (TSetRegim* Arr = &SetRegim)
                {
                    if (!DeviceIoControl(Handle, (uint)lu_SetRejime, (UInt16*)Arr, (uint)sizeof(TSetRegim), (UInt16*)NULL, 0, &ret, NULL))
                    {
                        err = GetLastError();
                        return false;
                    }
                    else return true;
                }
            }
            else return false;
        }                          //Открыть устройство
        unsafe public uint Write(TWrite FrameW) //Кадр Записи
        {
            UInt32 lu_SetRejime = (FILE_DEVICE_KMBO << 16) | (FILE_ANY_ACCESS << 14) | ((FIRST_IOCTL_INDEX + 1) << 2) | (METHOD_BUFFERED);
            UInt32 ret = 0;
            uint err = 0;
            if (!DeviceIoControl(Handle, (UInt32)lu_SetRejime, (UInt16*)&FrameW, (UInt32)sizeof(TWrite), (ushort*)NULL, 0, &ret, NULL))
            {
                err = GetLastError();
            }
            return err;
        }                          
        unsafe public uint Read(TRead FrameR, ref UInt16[] arr) //Кадр Чтения
        {
            UInt32 lu_SetRejime = (FILE_DEVICE_KMBO << 16) | (FILE_ANY_ACCESS << 14) | ((FIRST_IOCTL_INDEX) << 2) | (METHOD_BUFFERED);
            uint ret = 0;
            uint err = 0;

            TRead FrameR1;

            if (!DeviceIoControl(Handle, (UInt32)lu_SetRejime, (UInt16*)&FrameR, (UInt32)sizeof(TRead), (UInt16*)&FrameR1, (UInt32)sizeof(TRead), &ret, NULL))
            {
                err = GetLastError();
            }

            ushort* Arr = FrameR1.lpBuffer;
            {
                for (int i = 0; i < 32; i++) arr[i] = Arr[i];
            }
            return err;
        }          
    }
}




