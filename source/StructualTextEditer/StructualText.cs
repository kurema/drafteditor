using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace StructualTextEditer
{
    class StructualText
    {
        #region Member//メンバー変数
        public List<string> Text = new List<string>();
        public Structure Top = new Structure();
        #endregion

        #region Constructer//コンストラクター
        public StructualText()
        {
        }
        #endregion

        #region Reader//リーダークラス
        public class StructualTextReader
        {
            //ToDo.イメージ、著者、タイトル情報対応。
            #region Member//メンバー変数
            private StringReader sr;
            private string textBuffer = "";
            private StructualTextReader cache;
            public LineType Type;//行のタイプを示します。
            public int Depth = 0;//現在の構造の深さを示します。
            public string Title = "";//現在の構造のタイトル(章の名前とか)を示します。
            public string Text = "";//現在の行のテキストを示します。Mathの場合も当てはまります。
            #endregion

            #region Constucter//コンストラクター
            public StructualTextReader(string s)
            {
                sr = new StringReader(s);
            }
            #endregion

            #region Method//メソッド
            public enum LineType
            {
                Text,Math,Structure,EOF,Comment
                    ,Image
            }



            private static LineType GetLineType(string line){
                System.Text.RegularExpressions.Regex r =
                    new System.Text.RegularExpressions.Regex(@"^([\.\:\/：・]+)(.+)$");
                System.Text.RegularExpressions.Match m = r.Match(line);

                if (line.StartsWith("#Math:")) { return LineType.Math; }
                else if (line.StartsWith("#Image:")) { return LineType.Image; }
                else if (line.StartsWith("#")) { return LineType.Comment; }
                else if (m.Success) { return LineType.Structure; }
                else { return LineType.Text; }
            }

            public bool ReadBlock()
            {
                if (ReadLine())
                {
                    if (Type == LineType.Text)
                    {
                        string s = Text;
                        StructualTextReader temp = this;
                        while (ReadLine() && Type == LineType.Text)
                        {
                            s += "\n";
                            s += Text;
                        }
                        cache = new StructualTextReader("")
                        {
                            Text=this.Text,Depth=this.Depth,Title=this.Title,Type=this.Type
                        };
                        Text = s;
                        Depth = temp.Depth;
                        Title = temp.Title;
                        Type = LineType.Text;
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }

            public bool ReadLine()
            {
                string Line;
                if (cache != null)
                {
                    Text = cache.Text;
                    Title = cache.Title;
                    Depth = cache.Depth;
                    Type = cache.Type;

                    if (cache.Type == LineType.EOF)
                    {
                        cache = null;
                        return false;
                    }
                    else
                    {
                        cache = null;
                        return true;
                    }

                }
                else if (textBuffer == "")
                {
                    if ((Line = sr.ReadLine()) == null)
                    {
                        Type = LineType.EOF;
                        return false;
                    }
                }
                else
                {
                    Line = textBuffer;
                    textBuffer = "";
                    Type = GetLineType(Line);
                }

                Type = GetLineType(Line);
                if (Type == LineType.Text) { Text = Line; }
                else if (Type == LineType.Comment) {
                    Text = Line.Substring(1);
                }
                else if (Type == LineType.Structure)
                {
                    System.Text.RegularExpressions.Regex r =
                        new System.Text.RegularExpressions.Regex(@"^([\.\:\/：・]+)(.+)$");
                    System.Text.RegularExpressions.Match m = r.Match(Line);
                    Depth = m.Groups[1].Length;
                    Text = m.Groups[2].Value;
                    Title = m.Groups[2].Value;
                }
                else if (Type == LineType.Math)
                {
                    System.Text.RegularExpressions.Regex r =
                        new System.Text.RegularExpressions.Regex(@"^\#Math:(.+)$");
                    System.Text.RegularExpressions.Match m = r.Match(Line);
                    Text = m.Groups[1].Value;
                }
                else if (Type == LineType.Image)
                {
                    System.Text.RegularExpressions.Regex r =
                        new System.Text.RegularExpressions.Regex("^\\#Image:\\\"?(.+?)\\\"?$");
                    System.Text.RegularExpressions.Match m = r.Match(Line);

                    try
                    {
                        Text = System.IO.Path.GetFullPath(m.Groups[1].Value);
                    }
                    catch
                    {
                        Text = m.Groups[1].Value;
                    }
                }
                return true;
            }
            #endregion

        }
        #endregion

        #region Method//メソッド
        public static StructualText Parse(string text)
        {
            StructualText ret = new StructualText();
            StringReader sr = new StringReader(text);
            string line="";
            int num=0;
            int Place = 0;
            while ((line = sr.ReadLine()) != null)
            {
                System.Text.RegularExpressions.Regex r =
                    new System.Text.RegularExpressions.Regex(@"^([\.\:\/：・]+)(.+)$");
                System.Text.RegularExpressions.Match m = r.Match(line);
                Structure lastSt = ret.Top;                
                if (m.Success) {
                    ret.Text.Add(m.Groups[2].Value);
                    Structure st = ret.Top;
                    for (int i = 1; i < m.Groups[1].Length; i++)
                    {
                        if (st.Tree.Count==0)
                        {
                            Structure tst = new Structure(num,Place,lastSt);
                            lastSt.Next = tst;
                            lastSt = tst;
                            st.Tree.Add(tst);
                            st = tst;
                        }
                        else
                        {
                            st = st.Tree[st.Tree.Count() - 1];
                        }
                    }
                    Structure t2st = new Structure(num,Place, lastSt);
                    st.Tree.Add(t2st);
                    lastSt.Next = t2st;
                    lastSt = t2st;
                }
                else { ret.Text.Add(line); }
                num++;
                Place += line.Length+1;
            }
            return ret;
        }


        public class Structure
        {
            public int Line = 0;
            public int Place = 0;
            public List<Structure> Tree = new List<Structure>();
            public Structure Next=null;
            public Structure Back=null;

            public Structure() { }
            public Structure(int Line) { this.Line = Line; }
            public Structure(int Line,int Place):this(Line) { this.Place = Place; }
            public Structure(int Line, int Place, Structure Back) : this(Line,Place) { this.Back = Back; }
            public Structure(int Line, int Place, Structure Back, Structure Next) : this(Line,Place, Back) { this.Next = Next; }
        }
        #endregion
    }
}
