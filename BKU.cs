using System;
using System.Collections;

namespace RO2D
{
    public class BKU
    {
        public struct _Кадр_Записи
        {
            public UInt16 Nlit;
            public UInt16 Nn;
            public UInt16 KPK;

            public UInt16 Rd;
            public UInt16 Rk;
            public UInt16 Nzxv;
            public UInt16 Kzxv;
            public UInt16 Nr;
            public UInt16 SmNn;
            public UInt16 SmNlit;

            public UInt16 Dnr;
            public UInt16 Dvzm;
            public UInt16 RzROtz;
            public UInt16 RzRObu;

            public UInt16 TayR;
            public UInt16 Fg;

            public UInt16 AVK;
            public UInt16 VK;
            public UInt16 Vzm;
            public UInt16 Pr;
            public UInt16 tzaxv;
            public UInt16 Ptr;
            public UInt16 Vro;
            public UInt16 ZkRO;

            public UInt16 tci;
            public UInt16 tpr;

            public UInt16 Tc;
            public UInt16 t1j;
            public UInt16 Cheka;

            public UInt16 tzPr;
            public UInt16 ORO;
            public UInt16 ZaxvCl;
            public UInt16 BlM;
            public UInt16 PK;
            public UInt16 ServiceMode;
            public UInt16 ImpSTR;
            public UInt16 ImpBLN;
         //   public UInt16 BlockImpBLN;
        }
        public struct _Кадр_Чтения
        {
            public float MPrd;
            public UInt16 DPrd;

            public float Apchk;
            public UInt16 Dpchk;

            public UInt16 ZkRO;
            public UInt16 Kzxv;
            public UInt16 Sinxr;
            public UInt16 PK;
            public UInt16 GRO;
            public UInt16 GRO_LRK;
            public UInt16 Sinxr2;
            public UInt16 KeyOk;
            public UInt16 Sinxr1;
            public UInt16 OpnPrm;
            public UInt16 OpnPrd;
            public UInt16 ZaxvR;
            public UInt16 Nzxv;
            public UInt16 SmNlit;
            public UInt16 Rk;

            public UInt16 EndZaxvR;
            public UInt16 NottzPr;
            public UInt16 SmNn;
            public UInt16 Shod;
            public UInt16 RzROtz;
            public UInt16 RzRObu0;
            public UInt16 RCC;
            public UInt16 SoprR;
            public UInt16 KdVARU;
            public UInt16 Kl2_t6m20;
            public UInt16 Kl_t6m20;
            public UInt16 Kl_t1j;
            public UInt16 Kl_tci;

            public UInt16 NL_TM;
            public UInt16 KC_TM;

            public UInt16 GPPM_TM;
            public UInt16 ClZaxv;
            public UInt16 Version;

            public float ARU_TM;
        }
        public struct KMBO
        {
            public _Кадр_Записи Tx;
            public _Кадр_Чтения Rx;
        }
        public static KMBO gs_BKU = new KMBO();
        UInt16 old_Liter;
        //-------------------------------------------------------------------------------------------------------------//public function
        public ArrayList Pack_BKU()
        {
            ArrayList al = new ArrayList();
            UInt16 lu_temp;
            if (old_Liter != gs_BKU.Tx.Nlit)
            {
                gs_BKU.Tx.KPK = 1; /*gs_BKU.Tx.Rk = 1;*/ gs_BKU.Tx.SmNlit = 1;
                old_Liter = gs_BKU.Tx.Nlit;
            }
            else
            {
                gs_BKU.Tx.KPK = 0; /*gs_BKU.Tx.Rk = 1;*/ gs_BKU.Tx.SmNlit = 0;
            }

            lu_temp  = (UInt16)((gs_BKU.Tx.Nlit & 0x3ff) << 5);
            lu_temp |= (UInt16)((gs_BKU.Tx.Nn & 0xf) << 1);
            lu_temp |= (UInt16)(gs_BKU.Tx.KPK & 0x1);
            al.Add((UInt16)lu_temp);

            lu_temp = (gs_BKU.Tx.Rd == 1) ? (ushort)((2 & 0x3) << 14): (ushort)((1 & 0x3) << 14); 
            lu_temp |= (UInt16)((gs_BKU.Tx.Rk & 0x3) << 12);
            lu_temp |= (UInt16)((gs_BKU.Tx.Nzxv & 0x1) << 11);
            lu_temp |= (UInt16)((gs_BKU.Tx.Kzxv & 0x1) << 10);
            lu_temp |= (UInt16)((gs_BKU.Tx.Nr & 0x1f) << 5);
            lu_temp |= (UInt16)((gs_BKU.Tx.SmNn & 0x1) << 4);
            lu_temp |= (UInt16)((gs_BKU.Tx.SmNlit & 0x1) << 3);
            lu_temp |= (UInt16)(gs_BKU.Tx.ServiceMode & 0x7);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.Dnr & 0x7f) << 9);
            lu_temp |= (UInt16)((gs_BKU.Tx.Dvzm & 0x7f) << 2);
            lu_temp |= (UInt16)((gs_BKU.Tx.RzROtz & 0x1) << 1);
            lu_temp |= (UInt16)(gs_BKU.Tx.RzRObu & 0x1);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.TayR & 0xfff) << 4);
            lu_temp |= (UInt16)(gs_BKU.Tx.Fg & 0xf);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.AVK & 0x1) << 15);
            lu_temp |= (UInt16)((gs_BKU.Tx.VK & 0x1) << 14);
            lu_temp |= (UInt16)((gs_BKU.Tx.Vzm & 0x1) << 13);
            lu_temp |= (UInt16)((gs_BKU.Tx.Pr & 0x3) << 11);
            lu_temp |= (UInt16)((gs_BKU.Tx.tzaxv & 0x3f) << 5);
            lu_temp |= (UInt16)((gs_BKU.Tx.Ptr & 0x3) << 3);
            lu_temp |= (UInt16)((gs_BKU.Tx.Vro & 0x1) << 2);
            lu_temp |= (UInt16)((gs_BKU.Tx.ZkRO & 0x1) << 1);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.tci & 0xff) << 8);
            lu_temp |= (UInt16)(gs_BKU.Tx.tpr & 0xff);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.Tc & 0xff) << 8);
            lu_temp |= (UInt16)((gs_BKU.Tx.t1j & 0x3f) << 1);
            lu_temp |= (UInt16)(gs_BKU.Tx.Cheka & 0x1);
            al.Add((UInt16)lu_temp);

            lu_temp  = (UInt16)((gs_BKU.Tx.tzPr & 0xff) << 8);
          //  lu_temp |= (UInt16)((gs_BKU.Tx.BlockImpBLN & 0x1) << 6);
            lu_temp |= (UInt16)((gs_BKU.Tx.ImpSTR & 0x1) << 5);
            lu_temp |= (UInt16)((gs_BKU.Tx.ImpBLN & 0x1) << 4);
            lu_temp |= (UInt16)((gs_BKU.Tx.ORO & 0x1) << 3);
            lu_temp |= (UInt16)((gs_BKU.Tx.ZaxvCl & 0x1) << 2);
            lu_temp |= (UInt16)((gs_BKU.Tx.BlM & 0x1) << 1);
            lu_temp |= (UInt16)(gs_BKU.Tx.KPK & 0x1);
            al.Add((UInt16)lu_temp);

            return al;
        }
        public static bool gb_BKUChange = false;
        public void UnPackBKU(ArrayList al)
        {
            gs_BKU.Rx.MPrd    = (float)((((UInt16)al[0] & 0x3ffe) >> 1) * 5.0f / 512.0f);
            gs_BKU.Rx.DPrd    = (UInt16)((UInt16)(al[0]) & 0x1);

            gs_BKU.Rx.Apchk   = (float)((((UInt16)al[1] & 0x3ffe) >> 1) * 5.0f / 512.0f);
            gs_BKU.Rx.Dpchk   = (UInt16)((UInt16)(al[1]) & 0x1);

            gs_BKU.Rx.ZkRO    = (UInt16)(((UInt16)al[2] & 0x8000) >> 15);
            gs_BKU.Rx.Kzxv    = (UInt16)(((UInt16)al[2] & 0x4000) >> 14);
            gs_BKU.Rx.Sinxr   = (UInt16)(((UInt16)al[2] & 0x2000) >> 13);
            gs_BKU.Rx.PK      = (UInt16)(((UInt16)al[2] & 0x1000) >> 12);
            gs_BKU.Rx.GRO     = (UInt16)(((UInt16)al[2] & 0x0800) >> 11);
            gs_BKU.Rx.GRO_LRK = (UInt16)(((UInt16)al[2] & 0x0400) >> 10);
            gs_BKU.Rx.Sinxr2  = (UInt16)(((UInt16)al[2] & 0x0200) >> 9);
            gs_BKU.Rx.KeyOk   = (UInt16)(((UInt16)al[2] & 0x0100) >> 8);
            gs_BKU.Rx.Sinxr1  = (UInt16)(((UInt16)al[2] & 0x0080) >> 7);
            gs_BKU.Rx.OpnPrm  = (UInt16)(((UInt16)al[2] & 0x0040) >> 6);
            gs_BKU.Rx.OpnPrd  = (UInt16)(((UInt16)al[2] & 0x0020) >> 5);
            gs_BKU.Rx.ZaxvR   = (UInt16)(((UInt16)al[2] & 0x0010) >> 4);
            gs_BKU.Rx.Nzxv    = (UInt16)(((UInt16)al[2] & 0x0008) >> 3);
            gs_BKU.Rx.SmNlit  = (UInt16)(((UInt16)al[2] & 0x0004) >> 2);
            gs_BKU.Rx.Rk      = (UInt16)((UInt16)al[2] & 0x0003);

            gs_BKU.Rx.EndZaxvR = (UInt16)(((UInt16)al[3] & 0x8000) >> 15);
            gs_BKU.Rx.NottzPr  = (UInt16)(((UInt16)al[3] & 0x4000) >> 14);
            gs_BKU.Rx.SmNn     = (UInt16)(((UInt16)al[3] & 0x2000) >> 13);
            gs_BKU.Rx.Shod     = (UInt16)(((UInt16)al[3] & 0x1000) >> 12);
            gs_BKU.Rx.RzROtz   = (UInt16)(((UInt16)al[3] & 0x0800) >> 11);
            gs_BKU.Rx.RzRObu0  = (UInt16)(((UInt16)al[3] & 0x0400) >> 10);
            gs_BKU.Rx.RCC      = (UInt16)(((UInt16)al[3] & 0x0200) >> 9);
            gs_BKU.Rx.SoprR    = (UInt16)(((UInt16)al[3] & 0x0100) >> 8);
            gs_BKU.Rx.KdVARU   = (UInt16)(((UInt16)al[3] & 0x00f0) >> 4);

            gs_BKU.Rx.Kl2_t6m20 = (UInt16)(((UInt16)al[3] & 0x0008) >> 3);
            gs_BKU.Rx.Kl_t6m20  = (UInt16)(((UInt16)al[3] & 0x0004) >> 2);
            gs_BKU.Rx.Kl_t1j    = (UInt16)(((UInt16)al[3] & 0x0002) >> 1);
            gs_BKU.Rx.Kl_tci    = (UInt16)((UInt16)al[3] & 0x0001);

            gs_BKU.Rx.NL_TM     = (UInt16)((UInt16)al[4] & 0x0fff);
            gs_BKU.Rx.KC_TM     = (UInt16)(((UInt16)al[4] & 0xf000) >> 12);

            gs_BKU.Rx.GPPM_TM   = (UInt16)((UInt16)al[5] & 0x0001);
            gs_BKU.Rx.ClZaxv    = (UInt16)(((UInt16)al[5] & 0x0002) >> 1);
            gs_BKU.Rx.Version   = (UInt16)(((UInt16)al[5] & 0x03c0) >> 2);

            gs_BKU.Rx.ARU_TM    = (float)((((UInt16)al[6] & 0x3ffe) >> 1) * 5.0f / 512.0f);

            gb_BKUChange = true;

            if (FormMain.gb_Varu == true)
            {

                if ((gs_BKU.Tx.ImpSTR == 1) & (gs_BKU.Tx.ImpBLN == 0)) gs_BKU.Tx.ServiceMode = 0;
                FormMain.gb_Varu = false;
            }  

        }
    }
}