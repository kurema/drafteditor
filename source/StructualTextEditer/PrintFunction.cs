using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace StructualTextEditer
{
    public class PrintFunction
    {
        public class GetLettersInLineReturn
        {
            public int Position;
            public BreakType Break;
            public enum BreakType
            {
                EnterCode,EOF,Overflow
            }
        }

        public static GetLettersInLineReturn GetLettersInLine(string line,int width,Font f,Graphics g)
        {
            //入力テキストのエンターコードは\n固定
            int pos=0;
            while (true)
            {
                if(line.Length<=pos){
                    return new GetLettersInLineReturn()
                    {
                        Position = pos,
                        Break = GetLettersInLineReturn.BreakType.EOF
                    };
                }
                else if (line[pos] == '\n')
                {
                    return new GetLettersInLineReturn()
                    {
                        Position = pos,
                        Break = GetLettersInLineReturn.BreakType.EnterCode
                    };
                }
                else if (g.MeasureString(line.Substring(0, pos), f).Width > width)
                {
                    string notLast = "([（｛〔〈《「『【〘〖〝‘“｟«";
                    string notFirst = ",)）]｝、〕〉》」』】〙〗〟’”｠"
                        + "»ヽヾーァィゥェォッャュョヮヵヶぁぃぅぇぉっゃゅょゎゕゖㇰㇱㇲㇳㇴㇵㇶㇷㇸㇹㇺㇻㇼㇽㇾㇿ々〻"
                        + "‐゠–〜?!‼⁇⁈⁉・:;。.";
                    string Alphabet="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
                    if (notLast.Contains(line[Math.Max(pos - 2, 0)]))
                    {
                        return new GetLettersInLineReturn()
                        {
                            Position = Math.Max(pos - 2, 0),
                            Break = GetLettersInLineReturn.BreakType.Overflow
                        };
                    }
                    else if (notFirst.Contains(line[Math.Max(pos - 1, 0)]))
                    {
                        return new GetLettersInLineReturn()
                        {
                            Position = Math.Max(pos - 0, 0),
                            Break = GetLettersInLineReturn.BreakType.Overflow
                        };
                    }
                    else if (Alphabet.Contains(line[Math.Max(pos - 1, 0)]))
                    {
                        int Dif=0;
                        while (Alphabet.Contains(line[Math.Max(pos - 1 - Dif, 0)]))
                        {
                            if (Dif > 10)
                            {
                                return new GetLettersInLineReturn()
                                {
                                    Position = Math.Max(pos - 1 - Dif, 0),
                                    Break = GetLettersInLineReturn.BreakType.Overflow
                                };
                            }
                            Dif++;
                        }
                        return new GetLettersInLineReturn()
                        {
                            Position = Math.Max(pos - Dif, 0),
                            Break = GetLettersInLineReturn.BreakType.Overflow
                        };
                    }
                    else
                    {
                        return new GetLettersInLineReturn()
                        {
                            Position = Math.Max(pos - 1, 0),
                            Break = GetLettersInLineReturn.BreakType.Overflow
                        };
                    }
                }
                pos++;
            }
        }

        public class SimpleTextWriter
        {
            //装飾無しのテキスト出力。
            public LayoutInformation Layout;
            public string Text;
//            private int Position;
            public bool Continue=true;

            public SimpleTextWriter()
            {
                Layout = new LayoutInformation();
                Text = "";
            }
            public SimpleTextWriter(string Text) : this() { this.Text = Text.Replace("\r\n","\n").Replace("\r","\n"); }
            public SimpleTextWriter(LayoutInformation Layout) : this() { this.Layout = Layout; }
            public SimpleTextWriter(string Text, LayoutInformation Layout) : this(Text) { this.Layout = Layout; }

            public void WritePage(System.Drawing.Printing.PrintPageEventArgs e)
            {
                int x = e.MarginBounds.Left;
                int y = e.MarginBounds.Top;
                e.HasMorePages = true;

                while (e.MarginBounds.Bottom > y + Layout.Font.Height)
                {
                    GetLettersInLineReturn lil =
                        GetLettersInLine(Text, e.MarginBounds.Width-Layout.Margin.Left-Layout.Margin.Right, Layout.Font, e.Graphics);
                    if (lil.Break ==GetLettersInLineReturn.BreakType.EnterCode)
                    {

                        e.Graphics.DrawString(Text.Substring(0, lil.Position), Layout.Font, Layout.Brush,
                            (Layout.Align == Align.left ? x + Layout.Margin.Left :
                            (Layout.Align == Align.center ?
                            e.MarginBounds.Left + ((e.MarginBounds.Width - e.Graphics.MeasureString(Text.Substring(0, lil.Position), Layout.Font).Width) / 2) + Layout.Margin.Left
                            : e.MarginBounds.Left + e.MarginBounds.Width - e.Graphics.MeasureString(Text.Substring(0, lil.Position), Layout.Font).Width)) + Layout.Margin.Left
                            , y);

                        Text = Text.Substring(lil.Position + 1);
                        y += Layout.Font.Height + Layout.Margin.Bottom+Layout.Margin.Top;
                    }
                    else if (lil.Break == GetLettersInLineReturn.BreakType.Overflow)
                    {
                        //                        e.Graphics.DrawString(Text.Substring(0, lil.Position), Layout.Font, Layout.Brush, x + Layout.Margin.Left, y);
                        PrintLine(Text.Substring(0, lil.Position), Layout.Font, Brushes.Black, x + Layout.Margin.Left, y,
                            e.MarginBounds.Width, e);
                        Text = Text.Substring(lil.Position);

                        y += Layout.Font.Height + Layout.LineSpace;
                    }
                    else if (lil.Break == GetLettersInLineReturn.BreakType.EOF)
                    {
                        e.Graphics.DrawString(Text, Layout.Font, Layout.Brush,
                            (Layout.Align == Align.left ? x + Layout.Margin.Left : 
                            (Layout.Align == Align.center?
                            e.MarginBounds.Left + ((e.MarginBounds.Width - e.Graphics.MeasureString(Text, Layout.Font).Width) / 2) + Layout.Margin.Left
                            : e.MarginBounds.Left + e.MarginBounds.Width - e.Graphics.MeasureString(Text, Layout.Font).Width)) + Layout.Margin.Left
                            , y);
                        
                        e.HasMorePages = false;
                        Continue = false;
                        break;
                    }
                }
            }
        }

        public class LayoutedText
        {
            public struct LayoutedLine
            {
                public string Text;
                public LayoutInformation LayoutInformation;
            }
            public Dictionary<string, DisplayInfo> DisplayInfos = new Dictionary<string, DisplayInfo>();
            public List<LayoutedLine> Texts = new List<LayoutedLine>();
            public List<int> Count = new List<int>();
            public LayoutInformation DefaultLayout = new LayoutInformation();

            public void Add(string t)
            {
                Texts.Add(new LayoutedLine(){Text=t,LayoutInformation=DefaultLayout});
            }
            public void Add(string t, LayoutInformation l)
            {
                Texts.Add(new LayoutedLine() { Text = t, LayoutInformation = l });
            }
            public void Add(string t,DisplayInfo d,int ChapterLevel=-1)
            {
                IncrementCount(ChapterLevel);
                Texts.Add(new LayoutedLine() { Text = d.AdoptDisplayType(t, Count.ToArray()), LayoutInformation = d.Layout });
            }
            public void Add(string t, string keyword, int ChapterLevel = -1)
            {
                IncrementCount(ChapterLevel);
                DisplayInfo d=new DisplayInfo() { Layout=DefaultLayout};
                if (DisplayInfos.ContainsKey(keyword))
                {
                    d = DisplayInfos[keyword];
                }

                Texts.Add(new LayoutedLine() { Text = d.AdoptDisplayType(t, Count.ToArray()), LayoutInformation = d.Layout });
            }

            public void IncrementCount(int ChapterLevel)
            {
                if (ChapterLevel >= 1)
                {
                    int i = 0;
                    while (ChapterLevel > i)
                    {
                        if (Count.Count <= i) { Count.Add(0); }
                        i++;
                    }
                    Count[i - 1]++;
                }
            }
        }

        public class ExpandedTextWriter
        {
            //装飾アリのテキスト出力。
            public LayoutedText LayoutedText;
            public int Line=0;
            public int PositionInLine = 0;
            public bool Continue = true;
            private LayoutedText.LayoutedLine lLine;

            public ExpandedTextWriter()
            {
                this.LayoutedText = new LayoutedText();
                if (LayoutedText.Texts.Count() > 0)
                {
                    lLine = LayoutedText.Texts[Line];
                }
                else
                {
                    lLine = new LayoutedText.LayoutedLine() { Text = "", LayoutInformation = new LayoutInformation() };
                }
            }
            public ExpandedTextWriter(LayoutedText Text)
            {
                this.LayoutedText = Text;
                if (LayoutedText.Texts.Count() > 0)
                {
                    lLine = LayoutedText.Texts[Line];
                }
                else
                {
                    lLine = new LayoutedText.LayoutedLine() { Text="",LayoutInformation=new LayoutInformation()};
                }
            }

            public void WritePage(System.Drawing.Printing.PrintPageEventArgs e)
            {
                int x = e.MarginBounds.Left;
                int y = e.MarginBounds.Top;
                e.HasMorePages = true;

                while (e.MarginBounds.Bottom > y + lLine.LayoutInformation.Font.Height)
                {
                    GetLettersInLineReturn lil =
                        GetLettersInLine(lLine.Text, e.MarginBounds.Width - lLine.LayoutInformation.Margin.Left - lLine.LayoutInformation.Margin.Right, lLine.LayoutInformation.Font, e.Graphics);
                    if (lil.Break == GetLettersInLineReturn.BreakType.EnterCode)
                    {
                        e.Graphics.DrawString(lLine.Text.Substring(0, lil.Position), lLine.LayoutInformation.Font, lLine.LayoutInformation.Brush,
                            (lLine.LayoutInformation.Align == Align.left ? x + lLine.LayoutInformation.Margin.Left :
                            (lLine.LayoutInformation.Align == Align.center ?
                            e.MarginBounds.Left + ((e.MarginBounds.Width - e.Graphics.MeasureString(lLine.Text.Substring(0, lil.Position), lLine.LayoutInformation.Font).Width) / 2) + lLine.LayoutInformation.Margin.Left
                            : e.MarginBounds.Left + e.MarginBounds.Width - e.Graphics.MeasureString(lLine.Text.Substring(0, lil.Position), lLine.LayoutInformation.Font).Width)) + lLine.LayoutInformation.Margin.Left
                            , y);

                        lLine.Text = lLine.Text.Substring(lil.Position + 1);
                        y += lLine.LayoutInformation.Font.Height + lLine.LayoutInformation.Margin.Bottom + lLine.LayoutInformation.Margin.Top;
                    }
                    else if (lil.Break == GetLettersInLineReturn.BreakType.Overflow)
                    {
                        PrintLine(lLine.Text.Substring(0, lil.Position), lLine.LayoutInformation.Font, Brushes.Black, x, y + lLine.LayoutInformation.Margin.Left,
                            e.MarginBounds.Width, e);
                        lLine.Text = lLine.Text.Substring(lil.Position);

                        y += lLine.LayoutInformation.Font.Height + lLine.LayoutInformation.LineSpace;
                    }
                    else if (lil.Break == GetLettersInLineReturn.BreakType.EOF)
                    {
                        e.Graphics.DrawString(lLine.Text, lLine.LayoutInformation.Font, lLine.LayoutInformation.Brush,
                            (lLine.LayoutInformation.Align == Align.left ? x + lLine.LayoutInformation.Margin.Left :
                            (lLine.LayoutInformation.Align == Align.center ?
                            e.MarginBounds.Left + ((e.MarginBounds.Width - e.Graphics.MeasureString(lLine.Text, lLine.LayoutInformation.Font).Width) / 2) + lLine.LayoutInformation.Margin.Left
                            : e.MarginBounds.Left + e.MarginBounds.Width - e.Graphics.MeasureString(lLine.Text, lLine.LayoutInformation.Font).Width)) + lLine.LayoutInformation.Margin.Left
                            , y);

                        y += lLine.LayoutInformation.Font.Height + lLine.LayoutInformation.Margin.Bottom;
                        Line++;
                        if (LayoutedText.Texts.Count > Line)
                        {
                            lLine = LayoutedText.Texts[Line];
                        }
                        y += lLine.LayoutInformation.Margin.Top;
                    }
                    if (LayoutedText.Texts.Count <= Line)
                    {
                        e.HasMorePages = false;
                        Continue = false;
                        break;
                    }

                }
            }
        }


        public static void PrintLine(string text, Font f, Brush b, float x, float y, float width, System.Drawing.Printing.PrintPageEventArgs e)
        {
            System.Text.RegularExpressions.Regex r =new System.Text.RegularExpressions.Regex(@"^.+?[\s、。\t　]",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            System.Text.RegularExpressions.Match m = r.Match(text);
            if (!m.Success)
            {
                e.Graphics.DrawString(text, f, b, x, y);
            }
            else
            {
                string ts=text;
                float p=0;
                int num = 0;
                while (m.Success)
                {
                    p += (e.Graphics.MeasureString(m.Value, f).Width);
                    ts = ts.Substring(m.Length);
                    m = r.Match(ts);
                    num++;
                }
                p += (e.Graphics.MeasureString(ts, f).Width);
                m = r.Match(text);

                float space = (width - p) /num;
                while (m.Success)
                {
                    e.Graphics.DrawString(m.Value, f, b, x, y);
                    x += (e.Graphics.MeasureString(m.Value,f).Width);
                    x += space;
                    text = text.Substring(m.Length);
                    m = r.Match(text);
                }
                e.Graphics.DrawString(text, f, b, x, y);
            }
        }

        public class DisplayInfo
        {
            public LayoutInformation Layout = new LayoutInformation();
            public string DisplayType = TEXT;//%tはテキストをそのまま印刷する事を指す。
            public static string TEXT = "%TEXT%";
            public static string NUMBER = "%NUMBER%";
            public static string CHAP = "%CHAP:"+NUMBER+"%";
            public static string TIME = "%TIME%";
            public static string DATE = "%DATE%";

            public static string Chapter(int i)
            {
                return CHAP.Replace(NUMBER, i.ToString());
            }

            public string AdoptDisplayType(string text,params int[] count)
            {
                string s = DisplayType;
                System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(CHAP.Replace(NUMBER, @"(\d+)"));
                System.Text.RegularExpressions.MatchCollection mc = r.Matches(s);
                foreach (System.Text.RegularExpressions.Match m in mc)
                {
                    if (count.Length > int.Parse(m.Groups[1].Value))
                    {
                        s=s.Replace(m.Groups[0].Value, count[int.Parse(m.Groups[1].Value)].ToString());
                    }
                    else
                    {
                        s=s.Replace(m.Groups[0].Value, "0");
                    }
                }
                s=s.Replace(TEXT, text);

                s=s.Replace(TIME, DateTime.Now.ToShortTimeString());
                s=s.Replace(DATE, DateTime.Now.ToShortDateString());

                return s;

            }
        }
        public enum Align
        {
            center,left,right
        }
        public class LayoutInformation
        {
            public Margin Margin = new Margin();
            public int LineSpace = 0;
            public Font Font=new Font("ＭＳ Ｐゴシック", 10);
            public Brush Brush = Brushes.Black;
            public Align Align = Align.left;

            public LayoutInformation()
            {
                Margin=new Margin();
                LineSpace=3;
            }
            public LayoutInformation(int LineSpace):this() { this.LineSpace = LineSpace; }
            public LayoutInformation(Margin Margin) : this() { this.Margin = Margin; }
            public LayoutInformation(Font f) : this() { this.Font = f; }
            public LayoutInformation(int Top, int Bottom, int Left, int Right) : this() { this.Margin = new Margin(Top, Bottom, Left, Right); }
            public LayoutInformation(Margin Margin, int LineSpace) : this(Margin) { this.LineSpace = LineSpace; }
            public LayoutInformation(Margin Margin, int LineSpace,Font f) : this(Margin,LineSpace) { this.Font=f; }
            public LayoutInformation(Margin Margin, Font f) : this(Margin) { this.Font = f; }
            public LayoutInformation(int Top, int Bottom, int Left, int Right,int LineSpace)
                : this(Top, Bottom, Left, Right)
            { this.LineSpace=LineSpace; }
            public LayoutInformation(int Top, int Bottom, int Left, int Right, Font f): this(Top, Bottom, Left, Right)
            { this.Font=f; }
            public LayoutInformation(int Top, int Bottom, int Left, int Right, int LineSpace,Font f)
                : this(Top, Bottom, Left, Right,LineSpace)
            { this.Font = f; }

        }
        public class Margin
        {
            public int Top;
            public int Bottom;
            public int Left;
            public int Right;

            public Margin(){Top = 5; Bottom = 5; Left = 0; Right = 0;}
            public Margin(int Top, int Bottom, int Left, int Right)
            {this.Top = Top; this.Bottom = Bottom; this.Left = Left; this.Right = Right;}
        }
    }
}
