using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StructualTextEditer
{
    public class EncodeInfomation
    {
        //本クラスは以下のアドレスにおけるサンプルコードを一部書き換えたものです。
        //http://www.geocities.jp/gakaibon/tips/csharp2008/charset-check-samplecode3.html
        //サンプルコードを公開して頂いた雅階凡様には感謝いたします。
        //このソースコードのライセンスはそちらを参照してください。
        //また、このソースコードに関するお問い合わせは雅階凡様でなく作者の私にお願いいたします。

        public enum Unicode
        {
            LE,BE,None
        }

        public class EncodeBomInfo
        {
            public Encoding Encoding;
            public bool BOM;

            public static implicit operator EncodeBomInfo(Encoding enc){return new EncodeBomInfo(){Encoding=enc,BOM=false};}
            public static implicit operator Encoding(EncodeBomInfo enc){return enc.Encoding;}

        }


        public static EncodeBomInfo GetEncode(byte[] bytes)
        {
            Unicode utf16stat=IsUTF16(bytes);
            if (utf16stat!=Unicode.None)
            {
                if (utf16stat==Unicode.LE)
                    return Encoding.Unicode;

                else if (utf16stat == Unicode.BE)
                    return Encoding.BigEndianUnicode;
            }

            else if (IsJis(bytes) == true)
                return Encoding.GetEncoding(50220);

            else if (IsAscii(bytes) == true)
                return Encoding.ASCII;

            else
            {
                UTF8Info bUTF8 = IsUTF8(bytes);
                EncodeInfo bShitJis = IsShiftJis(bytes);
                EncodeInfo bEUC = IsEUC(bytes);

                if (bUTF8.Status == true || bShitJis.Status == true || bEUC.Status == true)
                {
                    if (bEUC.Fitness > bShitJis.Fitness && bEUC.Fitness > bUTF8.Fitness)
                        return Encoding.GetEncoding(51932);
                    else if (bShitJis.Fitness > bEUC.Fitness && bShitJis.Fitness > bUTF8.Fitness)
                        return Encoding.GetEncoding(932);
                    else if (bUTF8.Fitness > bEUC.Fitness && bUTF8.Fitness > bShitJis.Fitness){
                        return new EncodeBomInfo()
                        {
                            Encoding = Encoding.UTF8,
                            BOM = bUTF8.Bom
                        };
                    }
                }
                else
                {
                    return null;
                }
            }

            return null;
        }

        static public Unicode IsUTF16(byte[] bytes)      // Check for UTF-16
        {
            int len = bytes.Length;
            byte b1, b2;

            if (len >= 2)
            {
                b1 = bytes[0];
                b2 = bytes[1];

                if (b1 == 0xFF && b2 == 0xFE)
                {
                    return Unicode.LE;
                }
                else if (b1 == 0xFE && b2 == 0xFF)
                {
                    return Unicode.BE;
                }
                else
                {
                    return Unicode.None;
                }
            }
            else
            {
                return Unicode.None;
            }
        }

        static public bool IsJis(byte[] bytes)        // Check for JIS (ISO-2022-JP)
        {
            int len = bytes.Length;
            byte b1, b2, b3, b4, b5, b6;

            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];

                if (b1 > 0x7F)
                {
                    return false;   // Not ISO-2022-JP (0x00～0x7F)
                }
                else if (i < len - 2)
                {
                    b2 = bytes[i + 1]; b3 = bytes[i + 2];

                    if (b1 == 0x1B && b2 == 0x28 && b3 == 0x42)
                        return true;    // ESC ( B  : JIS ASCII

                    else if (b1 == 0x1B && b2 == 0x28 && b3 == 0x4A)
                        return true;    // ESC ( J  : JIS X 0201-1976 Roman Set

                    else if (b1 == 0x1B && b2 == 0x28 && b3 == 0x49)
                        return true;    // ESC ( I  : JIS X 0201-1976 kana

                    else if (b1 == 0x1B && b2 == 0x24 && b3 == 0x40)
                        return true;    // ESC $ @  : JIS X 0208-1978(old_JIS)

                    else if (b1 == 0x1B && b2 == 0x24 && b3 == 0x42)
                        return true;    // ESC $ B  : JIS X 0208-1983(new_JIS)
                }
                else if (i < len - 3)
                {
                    b2 = bytes[i + 1]; b3 = bytes[i + 2]; b4 = bytes[i + 3];

                    if (b1 == 0x1B && b2 == 0x24 && b3 == 0x28 && b4 == 0x44)
                        return true;    // ESC $ ( D  : JIS X 0212-1990（JIS_hojo_kanji）
                }
                else if (i < len - 5)
                {
                    b2 = bytes[i + 1]; b3 = bytes[i + 2];
                    b4 = bytes[i + 3]; b5 = bytes[i + 4]; b6 = bytes[i + 5];

                    if (b1 == 0x1B && b2 == 0x26 && b3 == 0x40 &&
                         b4 == 0x1B && b5 == 0x24 && b6 == 0x42)
                    {
                        return true;    // ESC & @ ESC $ B  : JIS X 0208-1990
                    }
                }
            }

            return false;
        }

        static public bool IsAscii(byte[] bytes)      // Check for Ascii
        {
            int len = bytes.Length;

            for (int i = 0; i < len; i++)
            {
                if (bytes[i] <= 0x7F)
                {
                    // ASCII : 0x00～0x7F
                    ;
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        public struct EncodeInfo{
            public int Fitness;
            public bool Status;
        }

        static public EncodeInfo IsShiftJis(byte[] bytes)   // Check for Shift-JIS
        {
            int len = bytes.Length;
            byte b1, b2;
            int sjis = 0;

            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];

                if ((b1 <= 0x7F) || (b1 >= 0xA1 && b1 <= 0xDF))
                {
                    // ASCII : 0x00～0x7F
                    // kana  : 0xA1～0xDF
                    ;
                }
                else if (i < len - 1)
                {
                    b2 = bytes[i + 1];

                    if (
                        ((b1 >= 0x81 && b1 <= 0x9F) || (b1 >= 0xE0 && b1 <= 0xFC)) &&
                        ((b2 >= 0x40 && b2 <= 0x7E) || (b2 >= 0x80 && b2 <= 0xFC))
                        )
                    {
                        // kanji first byte  : 0x81～0x9F、0xE0～0xFC
                        //       second byte : 0x40～0x7E、0x80～0xFC
                        i++;
                        sjis += 2;
                    }
                    else
                    {
                        return new EncodeInfo()
                        {
                            Fitness = sjis,
                            Status = false
                        };
                    }
                }
            }

            return new EncodeInfo()
            {
                Fitness = sjis,
                Status = true
            };
        }

        static public EncodeInfo IsEUC(byte[] bytes)        // Check for euc-jp 
        {
            int len = bytes.Length;
            byte b1, b2, b3;
            int euc = 0;

            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];

                if (b1 <= 0x7F)
                {   //  ASCII : 0x00～0x7F
                    ;
                }
                else if (i < len - 1)
                {
                    b2 = bytes[i + 1];

                    if ((b1 >= 0xA1 && b1 <= 0xFE) && (b2 >= 0xA1 && b2 <= 0xFE))
                    { // kanji - first & second byte : 0xA1～0xFE
                        i++;
                        euc += 2;
                    }
                    else if ((b1 == 0x8E) && (b2 >= 0xA1 && b2 <= 0xDF))
                    { // kana - first byte : 0x8E, second byte : 0xA1～0xDF
                        i++;
                        euc += 2;
                    }
                    else if (i < len - 2)
                    {
                        b3 = bytes[i + 2];

                        if ((b1 == 0x8F) &&
                            (b2 >= 0xA1 && b2 <= 0xFE) && (b3 >= 0xA1 && b3 <= 0xFE))
                        { // hojo kanji - first byte : 0x8F, second & third byte : 0xA1～0xFE
                            i += 2;
                            euc += 3;
                        }
                        else
                        {
                            return new EncodeInfo()
                            {
                                Fitness = euc,
                                Status = false
                            };
                        }
                    }
                    else
                    {
                        return new EncodeInfo()
                        {
                            Fitness = euc,
                            Status = false
                        };
                    }
                }
            }

            return new EncodeInfo()
            {
                Fitness = euc,
                Status = true
            };
        }

        public struct UTF8Info
        {
            public bool Bom;
            public int Fitness;
            public bool Status;
        }

        static public UTF8Info IsUTF8(byte[] bytes)       // Check for UTF-8
        {
            int len = bytes.Length;
            byte b1, b2, b3, b4;
            int utf8 = 0;
            bool bBOMStat = false;

            for (int i = 0; i < len; i++)
            {
                b1 = bytes[i];

                if (b1 <= 0x7F)
                { //  ASCII : 0x00～0x7F
                    ;
                }
                else if (i < len - 1)
                {
                    b2 = bytes[i + 1];

                    if ((b1 >= 0xC0 && b1 <= 0xDF) &&
                        (b2 >= 0x80 && b2 <= 0xBF))
                    { // 2 byte char
                        i += 1;
                        utf8 += 2;
                    }
                    else if (i < len - 2)
                    {
                        b3 = bytes[i + 2];

                        if (b1 == 0xEF && b2 == 0xBB && b3 == 0xBF)
                        { // BOM : 0xEF 0xBB 0xBF
                            bBOMStat = true;
                            i += 2;
                            utf8 += 3;
                        }
                        else if ((b1 >= 0xE0 && b1 <= 0xEF) &&
                            (b2 >= 0x80 && b2 <= 0xBF) &&
                            (b3 >= 0x80 && b3 <= 0xBF))
                        { // 3 byte char
                            i += 2;
                            utf8 += 3;
                        }

                        else if (i < len - 3)
                        {
                            b4 = bytes[i + 3];

                            if ((b1 >= 0xF0 && b1 <= 0xF7) &&
                                (b2 >= 0x80 && b2 <= 0xBF) &&
                                (b3 >= 0x80 && b3 <= 0xBF) &&
                                (b4 >= 0x80 && b4 <= 0xBF))
                            { // 4 byte char
                                i += 3;
                                utf8 += 4;
                            }
                        }
                        else
                        {
                            return new UTF8Info()
                            {
                                Fitness = utf8,
                                Bom = bBOMStat,
                                Status=false
                            };
                        }
                    }
                }
            }
            return new UTF8Info()
            {
                Fitness = utf8,
                Bom = bBOMStat,
                Status = true
            };
        }
    }
}
