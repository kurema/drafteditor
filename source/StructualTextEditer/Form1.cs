using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;


namespace StructualTextEditer
{
    public partial class Form1 : Form
    {
//        private PrintInfo pinfo;

        #region Constructer//コンストラクター

        public Form1()
        {
            InitializeComponent();

            AddCleanTab();
            this.Load += new EventHandler((s, e) => { SelectedRichTextBox().Focus(); });
        }
        #endregion

        #region Method//メソッド

        private void SaveTab(TabPage tp)
        {
            RichTextBox rt = (RichTextBox)tp.Controls[0];
            if (!((FileInfo)rt.Tag).isDefault())
            {
                saveFileDialogTXT.InitialDirectory = Path.GetDirectoryName(((FileInfo)rt.Tag).FileName);
                saveFileDialogTXT.FileName = Path.GetFileName(((FileInfo)rt.Tag).FileName);
            }
            if (saveFileDialogTXT.ShowDialog() == DialogResult.OK)
            {
                //using (System.IO.StreamWriter sw = new System.IO.StreamWriter(saveFileDialog1.FileName, false))
                //{
                //    sw.Write(rt.Text);
                //}
                rt.SaveFile(saveFileDialogTXT.FileName, RichTextBoxStreamType.PlainText);
                ((FileInfo)rt.Tag).FileName = saveFileDialogTXT.FileName;
                ((FileInfo)rt.Tag).IsSaved = true;
                ((FileInfo)rt.Tag).Title = Path.GetFileName(saveFileDialogTXT.FileName);
            }
        }

        private void SaveAll()
        {
            foreach (TabPage tp in tabControl1.TabPages)
            {
                if (!((FileInfo)((RichTextBox)tp.Controls[0]).Tag).IsSaved)
                {
                    OverWriteTab(tp);
                }
            }
        }

        private RichTextBox SelectedRichTextBox() { return ((RichTextBox)tabControl1.SelectedTab.Controls[0]); }

        private void AddIndent(string s)
        {
            SelectedRichTextBox().Focus();
            StringReader sr = new StringReader(SelectedRichTextBox().SelectedText);
            string temp = "";
            string line = "";
            while ((line = sr.ReadLine()) != null)
            {
                temp += s + line + System.Environment.NewLine;
            }
            SelectedRichTextBox().SelectedText = temp;
        }

        private void SubIndent(string s)
        {
            SelectedRichTextBox().Focus();
            StringReader sr = new StringReader(SelectedRichTextBox().SelectedText);
            string temp = "";
            string line = "";
            while ((line = sr.ReadLine()) != null)
            {
                if (line.StartsWith(s))
                {
                    temp += line.Substring(s.Length) + System.Environment.NewLine;
                }
                else
                {
                    temp += line + System.Environment.NewLine;
                }
            }
            SelectedRichTextBox().SelectedText = temp;
        }        
        
        private void LoadTree(StructualText st)
        {
            treeView1.Nodes.Clear();
            LoadTreeRep(ref st, st.Top, treeView1.Nodes);
            treeView1.ExpandAll();
        }

        private void LoadTreeRep(ref StructualText st, StructualText.Structure stt, TreeNodeCollection tn)
        {
            foreach (StructualText.Structure st2 in stt.Tree)
            {
                TreeNode tn2 = new TreeNode(st.Text[st2.Line]);
                LoadTreeRep(ref st, st2, tn2.Nodes);
                tn2.Tag = st2;
                tn.Add(tn2);
            }
        }
        private TabPage AddCleanTab()
        {
            TabPage t = new TabPage();
            RichTextBox rt = new RichTextBox();
            rt.TextChanged += richTextBox1_TextChanged;
            rt.Tag = new FileInfo();
            t.Controls.Add(rt);
            tabControl1.TabPages.Add(t);
            rt.AcceptsTab = true;
            rt.Width = t.Width;
            rt.Height = t.Height;
            rt.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            RenewTabTitle(t);
            return t;
        }

        #endregion

        #region FileInfo//クラス
        private class FileInfo
        {
            public string FileName = null;
            public string Title = "新規ファイル";
            public bool IsEdited = false;
            public bool IsSaved = true;

            public bool isDefault()
            {
                return FileName == null;
            }
        }

        #endregion

        #region PrintInfo//クラス
        private class PrintInfo
        {
            private StructualText.StructualTextReader sr;
            private string buffer = "";
            private int printPoint = 0;
            private Font currentFont=new Font("ＭＳ Ｐゴシック", 12);
            private List<Font> FontSetting = new List<Font>();//0:標準,1:数式モード,2:システム(ページ番号など),3～:レイアウト
            private bool isOver = false;
            private Font lastFont;

            public PrintInfo(string text) { sr = new StructualText.StructualTextReader(text); SetDefaultFont(); ReadLine(); }

            public char NextLetter()
            {
                if (buffer.Length <= printPoint)
                {
                    ReadLine();
                    if(isOver){return '\0';}else{return '\n';}
                }
                else
                {
                    printPoint++;
                    return buffer[printPoint - 1];
                }
            }

            private void ReadLine()
            {
                do
                {
                    sr.ReadLine();
                }
                while (sr.Type == StructualText.StructualTextReader.LineType.Comment);
                buffer = sr.Text;
                lastFont = currentFont;
                if (sr.Type == StructualText.StructualTextReader.LineType.Text) { currentFont = BasicFont(); }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Math) { currentFont = MathFont(); }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Structure) { currentFont = GetFontByStructure(sr.Depth); }
                else if (sr.Type == StructualText.StructualTextReader.LineType.EOF)
                {
                    isOver = true;
                }
                printPoint = 0;
            }

            public bool IsOver() { return isOver; }

            public Font CurrentFont() { return currentFont; }
            public Font LastFont() { return lastFont; }

            public Font GetFontByStructure(int num)
            {
                num = Math.Max(num, 1);
                return FontSetting[Math.Min(2+num,FontSetting.Count)];
            }
            public Font SystemFont() { return FontSetting[2]; }
            public Font MathFont() { return FontSetting[1]; }
            public Font BasicFont() { return FontSetting[0]; }

            public void SetSystemFont(Font f) { FontSetting[2] = f; }
            public void SetMathFont(Font f) { FontSetting[1] = f; }
            public void SetBasicFont(Font f) { FontSetting[0] = f; }
            public void SetBasicFont(params Font[] f) { FontSetting.RemoveRange(3, FontSetting.Count - 2); FontSetting.AddRange(f); }

            private void SetDefaultFont()
            {
                FontSetting.Clear();
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 10));//標準
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 10));//Math
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 10));//System
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 20));//タイトル
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 18));//章
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 16));//見出し1
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 14));//見出し2
                FontSetting.Add(new Font("ＭＳ Ｐゴシック", 12));//見出し3
            }
        }
        #endregion

        #region event //イベント処理
        //private void docx形式で保存DToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //}



        private void treeView1_AfterSelect_1(object sender, TreeViewEventArgs e)
        {
            RichTextBox rt=(RichTextBox) tabControl1.SelectedTab.Controls[0];
            rt.SelectionStart = ((StructualText.Structure)treeView1.SelectedNode.Tag).Place;
            rt.Focus();
            rt.ScrollToCaret();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (!((FileInfo)((RichTextBox)sender).Tag).IsEdited)
            {
                AddCleanTab();
                ((FileInfo)((RichTextBox)sender).Tag).IsEdited = true;
            }
            if (階層パネルToolStripMenuItem.Checked)
            {
                LoadTree(StructualText.Parse(((RichTextBox)sender).Text));
            }
            ((FileInfo)((RichTextBox)sender).Tag).IsSaved = false;
            RenewTabTitle(((TabPage)((RichTextBox)sender).Parent));

            if (statusStrip1.Visible)
            {
                toolStripStatusLabel1.Text = String.Format("{0:#,0}文字 ", ((RichTextBox)sender).Text.Length);
                toolStripStatusLabel2.Text = String.Format("{0:#,0}B ", Encoding.UTF8.GetByteCount(((RichTextBox)sender).Text));
            }

        }

        private void バージョン情報VToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tools.AboutBox1 ab = new Tools.AboutBox1();
            ab.Show();
        }


        private void 保存SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveTab(tabControl1.SelectedTab);
            RenewTabTitle(tabControl1.SelectedTab);
        }


        private void 上書き保存SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OverWriteTab(tabControl1.SelectedTab);
        }

        private void OverWriteTab(TabPage tab)
        {
            RichTextBox rt = (RichTextBox)tab.Controls[0];
            if (((FileInfo)rt.Tag).FileName == null)
            {
                SaveTab(tab);
            }
            else
            {
                //using (System.IO.StreamWriter sw = new System.IO.StreamWriter(((FileInfo)rt.Tag).FileName, false))
                //{
                //    sw.Write(rt.Text);
                //}
                rt.SaveFile(((FileInfo)rt.Tag).FileName, RichTextBoxStreamType.PlainText);
                ((FileInfo)rt.Tag).IsSaved = true;
            }
            RenewTabTitle(tab);
        }

        private void RenewTabTitle(TabPage tp)
        {
            RichTextBox rt = (RichTextBox)tp.Controls[0];
            if (((FileInfo)rt.Tag).IsSaved) { tp.Text = ((FileInfo)rt.Tag).Title; }
            else { tp.Text = ((FileInfo)rt.Tag).Title + "*"; }
        }

        private void 全て保存EToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAll();
        }

        private void 終了EToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Cut();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Copy();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            InsertText(Clipboard.GetText());
//            SelectedRichTextBox().Paste();
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Undo();
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Redo();
        }

        private void 開くOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialogTXT.ShowDialog() == DialogResult.OK)
            {
                TabPage tp = tabControl1.TabPages[tabControl1.TabPages.Count - 1];
                ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).FileName = openFileDialogTXT.FileName;
                //((RichTextBox)tp.Controls[0]).LoadFile(openFileDialog1.FileName,RichTextBoxStreamType.PlainText);//文字が化ける

                //using(StreamReader sr=new StreamReader(openFileDialog1.FileName,Encoding.UTF8,true)){
                //    while (!sr.EndOfStream)
                //    {
                //        ((RichTextBox)tp.Controls[0]).Text+=sr.ReadLine()+"\n";
                //    }
                //}//重い。実はこっちでも化ける。
                
                using(FileStream fs=new FileStream(openFileDialogTXT.FileName,FileMode.Open,FileAccess.Read)){
                    byte[] bs = new byte[fs.Length];
                    fs.Read(bs, 0, bs.Length);
                    EncodeInfomation.EncodeBomInfo enc = EncodeInfomation.GetEncode(bs);
                    if (enc.Encoding == Encoding.UTF8 && enc.BOM==true)
                    {
                        ((RichTextBox)tp.Controls[0]).Text = enc.Encoding.GetString(bs, 3, bs.Length - 3);
                    }
                    else if (enc.Encoding == Encoding.UTF8 && enc.BOM == false)
                    {
                        ((RichTextBox)tp.Controls[0]).Text = enc.Encoding.GetString(bs);
                    }
                    else if (enc.Encoding == Encoding.Unicode ||
                                 enc.Encoding == Encoding.BigEndianUnicode)
                    {
                        ((RichTextBox)tp.Controls[0]).Text = enc.Encoding.GetString(bs, 2, bs.Length - 2);
                    }
                    else if (enc == null)
                    {
                        ((RichTextBox)tp.Controls[0]).LoadFile(openFileDialogTXT.FileName,RichTextBoxStreamType.PlainText);
                    }
                    else
                    {
                        ((RichTextBox)tp.Controls[0]).Text = enc.Encoding.GetString(bs);
                    }
                }


                ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).Title = Path.GetFileName(openFileDialogTXT.FileName);
                ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).IsEdited = true;
                ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).IsSaved = true;
                RenewTabTitle(tp);
            }
        }

        private void Form1Closing(object sender, FormClosingEventArgs e)
        {
            bool t = true;
            foreach (TabPage tp in tabControl1.TabPages)
            {
                t = t && ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).IsSaved;
            }
            if (!t)
            {
                DialogResult dr = MessageBox.Show("内容が変更されています。全ての変更を保存しますか。", "確認"
                    , MessageBoxButtons.YesNoCancel);
                if (dr == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
                else if (dr == DialogResult.No)
                {
                    e.Cancel = false;
                }
                else
                {
                    SaveAll();
                    bool p = true;
                    foreach (TabPage tp in tabControl1.TabPages)
                    {
                        p = p && ((FileInfo)((RichTextBox)tp.Controls[0]).Tag).IsSaved;
                    }
                    if (p) { e.Cancel = false; }
                    else { e.Cancel = true; }
                }
            }
            else
            {
                e.Cancel = false;
            }

        }

        private void 検索SToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SearchDialog sc = new SearchDialog(SearchDialog.dialogMode.Find,SelectedRichTextBox());
            sc.Show();
            sc.SearchStatusChanged += new SearchDialog.SearchStatusChangedEventHandler(sc_SearchStatusChanged);
        }

        void sc_SearchStatusChanged(object sender, SearchDialog.SearchStatusChangedEventArgs e)
        {
            toolStripTextBox1.Text = e.SearchText();
        }

        private void 置換RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SearchDialog sc = new SearchDialog(SearchDialog.dialogMode.Replace, SelectedRichTextBox());
            sc.Show();
            sc.SearchStatusChanged += new SearchDialog.SearchStatusChangedEventHandler(sc_SearchStatusChanged);
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            SearchDialog.ExecSearch(SelectedRichTextBox(), toolStripTextBox1.Text);
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            SearchDialog.ExecSearch(SelectedRichTextBox(), toolStripTextBox1.Text,false);
        }

        private void InsertText(string text)
        {
            SelectedRichTextBox().Focus();
            SelectedRichTextBox().SelectedText = text;
        }

        private void InsertTextWithEnter(string text)
        {
            RichTextBox rt=SelectedRichTextBox();
            rt.Focus();
            char t=rt.TextLength>0?rt.Text[Math.Max(rt.SelectionStart - 1,0)]:'\n';
            if (t == '\n' || t == '\r' || rt.SelectionStart == 0)
            {
                SelectedRichTextBox().SelectedText = text;
            }
            else
            {
                SelectedRichTextBox().SelectedText = "\n" + text;
            }
        }

        private void 日付と時刻TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertText(DateTime.Now.ToString());
        }

        private void 署名SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertText(string.Format("{0} ({1})", Environment.UserName,Environment.MachineName));
        }

        private void 数式MToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertTextWithEnter("#Math:");
        }

        private void 階層パネルToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadTree(StructualText.Parse(SelectedRichTextBox().Text));
            splitContainer1.Panel1Collapsed = !階層パネルToolStripMenuItem.Checked;
        }

        private void 引用貼り付けQToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringReader sr = new StringReader(Clipboard.GetText());
            string temp = "";
            string line = "";
            while ((line = sr.ReadLine()) != null)
            {
                temp += ">" + line + System.Environment.NewLine;
            }
            InsertText(temp);

        }

        private void 小文字に変換LToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Focus();
            SelectedRichTextBox().SelectedText = SelectedRichTextBox().SelectedText.ToLower();
        }

        private void 大文字に変換UToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Focus();
            SelectedRichTextBox().SelectedText = SelectedRichTextBox().SelectedText.ToUpper();
        }

        private void タブインデントTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddIndent("\t");
        }

        private void タブインデントAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SubIndent("\t");
        }

        private void 空白インデントSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddIndent(" ");
        }

        private void 空白逆インデントPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SubIndent(" ");
        }

        private void 空白削除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Focus();
            SelectedRichTextBox().SelectedText=SelectedRichTextBox().SelectedText.Replace("\t","").Replace(" ","");
        }

        private void 改行削除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SelectedRichTextBox().Focus();
            SelectedRichTextBox().SelectedText=SelectedRichTextBox().SelectedText.Replace("\n","");
        }

        private void gUIDGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertText(Guid.NewGuid().ToString());
        }

        private void docx形式DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!((FileInfo)SelectedRichTextBox().Tag).isDefault())
            {
                saveFileDialogDocx.FileName = Path.GetFileNameWithoutExtension(((FileInfo)SelectedRichTextBox().Tag).FileName);
            }
            if (saveFileDialogDocx.ShowDialog() == DialogResult.OK)
            {
                //using (var package =
                //    DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create
                //    (saveFileDialog2.FileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                //{
                //    DocumentFormat.OpenXml.Packaging.MainDocumentPart maindoc = package.AddMainDocumentPart();
                //    maindoc.Document = Converter.OfficeOpenXMLDocument(SelectedRichTextBox().Text);
                //}
                Converter.SaveOfficeOpenXML(SelectedRichTextBox().Text, saveFileDialogDocx.FileName);
            }
        }

        private void コメントCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertTextWithEnter("#");
        }

        private void xHTMLHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!((FileInfo)SelectedRichTextBox().Tag).isDefault())
            {
               saveFileDialogXHTML.FileName = Path.GetFileNameWithoutExtension(((FileInfo)SelectedRichTextBox().Tag).FileName);
            }
            if (saveFileDialogXHTML.ShowDialog() == DialogResult.OK)
            {
                List<Converter.XhtmlInfo> ls= Converter.Xhtml(SelectedRichTextBox().Text, 1).Xhtml;
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(saveFileDialogXHTML.FileName, false,
                    System.Text.Encoding.GetEncoding("utf-8")))
                {
                    sw.Write(ls[0].Text);
                    sw.Close();
                }
            }
        }
        #endregion

        private void tex形式TToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!((FileInfo)SelectedRichTextBox().Tag).isDefault())
            {
                saveFileDialogTex.FileName = Path.GetFileNameWithoutExtension(((FileInfo)SelectedRichTextBox().Tag).FileName);
            }
            if (saveFileDialogTex.ShowDialog() == DialogResult.OK)
            {
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(saveFileDialogTex.FileName, false,
                    System.Text.Encoding.GetEncoding("shift_jis")))
                {
                    sw.Write(Converter.Tex(SelectedRichTextBox().Text));
                    sw.Close();
                }
            }

        }

        private void printDocument1_PrintPage_1(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //si.WritePage(e);
            ei.WritePage(e);
            ////印刷する初期位置を決定
            //int x = e.MarginBounds.Left;
            //int y = e.MarginBounds.Top;
            //string linebuf = "";
            //Font tempFont;

            ////1ページに収まらなくなるか、印刷する文字がなくなるかまでループ
            //while (e.MarginBounds.Bottom > y + pinfo.CurrentFont().Height+2 && !pinfo.IsOver())
            //{
            //    string line = linebuf;
            //    for (; ; )
            //    {
            //        char nextLetter = pinfo.NextLetter();
            //        //印刷する文字がなくなるか、
            //        //改行の時はループから抜けて印刷する
            //        if (pinfo.IsOver() ||nextLetter == '\n')
            //        {
            //            linebuf = "";
            //            tempFont = pinfo.LastFont();
            //            break;
            //        }
            //        //一文字追加し、印刷幅を超えるか調べる
            //        line += nextLetter;
            //        if (e.Graphics.MeasureString(line, pinfo.CurrentFont()).Width
            //            > e.MarginBounds.Width)
            //        {
            //            //印刷幅を超えたため、折り返す
            //            line = line.Substring(0, line.Length - 1);
            //            linebuf = nextLetter.ToString();
            //            tempFont = pinfo.CurrentFont();
            //            break;
            //        }
            //    }
            //    //一行書き出す
            //    e.Graphics.DrawString(line, tempFont, Brushes.Black, x, y);
            //    //次の行の印刷位置を計算
            //    y += (int)tempFont.GetHeight(e.Graphics)+2;
            //}

            ////次のページがあるか調べる
            //if (pinfo.IsOver())
            //{
            //    e.HasMorePages = false;
            //}
            //else
            //    e.HasMorePages = true;
        }

        private void epub形式DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!((FileInfo)SelectedRichTextBox().Tag).isDefault())
            {
                saveFileDialogEPUB.FileName = Path.GetFileNameWithoutExtension(((FileInfo)SelectedRichTextBox().Tag).FileName);
            }
            if (saveFileDialogEPUB.ShowDialog() == DialogResult.OK)
            {
                Converter.SaveEPub(SelectedRichTextBox().Text, saveFileDialogEPUB.FileName);
            }
        }

        private void ステータスバーSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip1.Visible = ステータスバーSToolStripMenuItem.Checked;
        }

        private void 画像PToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertTextWithEnter("#Image:");
        }

        private void 印刷PToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ei = PrintStructualText(SelectedRichTextBox().Text);
//            pinfo = new PrintInfo(SelectedRichTextBox().Text);
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.PrinterSettings = printDialog1.PrinterSettings;
                printDocument1.Print();
            }
        }

        private void 用紙設定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pageSetupDialog1.ShowDialog() == DialogResult.OK)
            {
                printDocument1.DefaultPageSettings = pageSetupDialog1.PageSettings;
            }
        }
//        private PrintFunction.SimpleTextWriter si;
        private PrintFunction.ExpandedTextWriter ei;

        private void 印刷プレビューToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintPreviewDialog pv = new PrintPreviewDialog();
//            pinfo = new PrintInfo(SelectedRichTextBox().Text);
//            si = new PrintFunction.SimpleTextWriter(SelectedRichTextBox().Text, new PrintFunction.LayoutInformation(3, 3, 1,1, 3));
            ei = PrintStructualText(SelectedRichTextBox().Text);
            pv.Document = printDocument1;
            pv.Show();
        }

        public static PrintFunction.ExpandedTextWriter PrintStructualText(string s)
        {
            PrintFunction.LayoutedText lt = new PrintFunction.LayoutedText()
            {
                DefaultLayout = new PrintFunction.LayoutInformation(3, 3, 1, 1, 3)
                {
                    Font=new Font("Meiryo",12)
                }
            };

            lt.DisplayInfos.Add("ChapLv1", new PrintFunction.DisplayInfo()
            {
                DisplayType = PrintFunction.DisplayInfo.TEXT+"\n"+PrintFunction.DisplayInfo.DATE,
                Layout = new PrintFunction.LayoutInformation(10, 5, 5, 10, 5)
                {
                    Font=new Font("Meiryo",20),
                    Align=PrintFunction.Align.center
                }
            }
            );
            lt.DisplayInfos.Add("ChapLv2", new PrintFunction.DisplayInfo()
            {
                DisplayType = "第" + PrintFunction.DisplayInfo.Chapter(1) + "部\n" + PrintFunction.DisplayInfo.TEXT,
                Layout = new PrintFunction.LayoutInformation(8, 3, 3, 8, 4)
                {
                    Font = new Font("Meiryo", 18),
                }
            }
            );
            lt.DisplayInfos.Add("ChapLv3", new PrintFunction.DisplayInfo()
            {
                DisplayType = PrintFunction.DisplayInfo.Chapter(2) + " " + PrintFunction.DisplayInfo.TEXT,
                Layout = new PrintFunction.LayoutInformation(6, 3, 2, 6, 3)
                {
                    Font = new Font("Meiryo", 16),
                }
            }
            );
            lt.DisplayInfos.Add("ChapLv4", new PrintFunction.DisplayInfo()
            {
                DisplayType = PrintFunction.DisplayInfo.Chapter(2) + "."+
                PrintFunction.DisplayInfo.Chapter(3) + " "+ PrintFunction.DisplayInfo.TEXT,
                Layout = new PrintFunction.LayoutInformation(6, 3, 2, 6, 3)
                {
                    Font = new Font("Meiryo", 16),
                }
            }
            );
            lt.DisplayInfos.Add("ChapLv5", new PrintFunction.DisplayInfo()
            {
                DisplayType = PrintFunction.DisplayInfo.Chapter(2) + "." +
                PrintFunction.DisplayInfo.Chapter(3) +"."+
                PrintFunction.DisplayInfo.Chapter(2) + "." +" " + PrintFunction.DisplayInfo.TEXT,
                Layout = new PrintFunction.LayoutInformation(6, 3, 2, 6, 3)
                {
                    Font = new Font("Meiryo", 16),
                }
            }
            );
            lt.DisplayInfos.Add("Math", new PrintFunction.DisplayInfo()
            {
                DisplayType = PrintFunction.DisplayInfo.TEXT,
                Layout = new PrintFunction.LayoutInformation(3, 3, 1, 1, 3)
                {
                    Font = new Font("Meiryo", 12),
                    Align=PrintFunction.Align.center
                }
            }
            );

            StructualText.StructualTextReader sr = new StructualText.StructualTextReader(s);
            while (sr.ReadBlock())
            {
                if (sr.Type == StructualText.StructualTextReader.LineType.Text)
                {
                    lt.Add(sr.Text);
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Math)
                {
                    lt.Add(sr.Text, "Math");
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Structure && sr.Depth <= 5)
                {
                    lt.Add(sr.Text, "ChapLv" + sr.Depth, sr.Depth);
                }
                else if(sr.Type == StructualText.StructualTextReader.LineType.Structure)
                {
                    lt.Add(sr.Text);
                }
            }
            return new PrintFunction.ExpandedTextWriter(lt);

        }

        private void ヘルプHToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("help.chm");
        }

        private void 閉じるXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void コメントアウトCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddIndent("#");
        }

        private void コメント解除HToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SubIndent("#");
        }

        private void 画像を開くOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialogPicture.ShowDialog() == DialogResult.OK)
            {
                InsertTextWithEnter("#Image:\"" + openFileDialogPicture.FileName + "\"");
            }
        }

        private void xML相互運用目的ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!((FileInfo)SelectedRichTextBox().Tag).isDefault())
            {
                saveFileDialogXML.FileName = Path.GetFileNameWithoutExtension(((FileInfo)SelectedRichTextBox().Tag).FileName);
            }
            if (saveFileDialogXML.ShowDialog() == DialogResult.OK)
            {
                System.Xml.XmlDocument xd = Converter.XML(SelectedRichTextBox().Text);
                xd.Save(saveFileDialogXML.FileName);
            }
        }

        private void xSLTを適用ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialogXslt.ShowDialog() == DialogResult.OK && saveFileDialogXslt.ShowDialog() == DialogResult.OK)
            {
                System.Xml.Xsl.XslCompiledTransform xslt = new System.Xml.Xsl.XslCompiledTransform();
                string temppath = System.IO.Path.GetTempFileName();

                xslt.Load(openFileDialogXslt.FileName);
                System.Xml.XmlDocument xd = Converter.XML(SelectedRichTextBox().Text);
                xd.Save(temppath);
                xslt.Transform(temppath, saveFileDialogXslt.FileName);
            }
        }
    }
}
