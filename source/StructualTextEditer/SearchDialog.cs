using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;

namespace StructualTextEditer
{
    public partial class SearchDialog : Form
    {
        private TextBoxBase _textBox;
        private dialogMode _mode;
        private string _searchString;


        public enum dialogMode
        {
            Find,Replace
        }

        #region Constructor //コンストラクタ


        public SearchDialog(dialogMode mode, TextBoxBase tBox)
        {
            InitializeComponent();

            this._mode = mode;
            this._textBox = tBox;

            ChangeDialogMode(_mode);
        }

        public SearchDialog(TextBoxBase tBox) : this(dialogMode.Find, tBox) { }
        public SearchDialog() : this(new TextBox()) { }
        #endregion

        private void DoSearch(bool fromTop)
        {
            //イベント呼び出し。
            OnSearchStatusChanged(new SearchStatusChangedEventArgs
                (textBox1.Text,_mode,checkBox1.Checked,checkBox2.Checked,checkBox3.Checked,radioButton2.Checked,textBox2.Text));
            if (checkBox3.Checked)
            {
                SearchRegularExpression(_textBox, textBox1.Text, radioButton2.Checked, fromTop, checkBox2.Checked);
            }
            else
            {
                ExecSearch(_textBox, textBox1.Text, radioButton2.Checked, fromTop, checkBox2.Checked, checkBox1.Checked);
            }
        }

        public static void ExecSearch(TextBoxBase tbox, string Original, bool ToDown=true, bool fromTop=false, bool Escape = false, bool Capital = false)
        {
            TextBoxBase _textBox = tbox;
            string targetString = _textBox.Text;
            string searchString =(Escape)?
                Original.Replace(@"\n","\n").Replace(@"\t","\t").Replace(@"\\","\\"):Original;
            
//            searchStartIndex=(searchStartIndex==0)?((fromTop)?0:_textBox.SelectionStart):searchStartIndex;
//            searchStartIndex=(fromTop)?0:((searchStartIndex==0)?_textBox.SelectionStart:searchStartIndex);
            int searchStartIndex = ((fromTop) ? 
                ((ToDown)?0:_textBox.Text.Length) :
                (ToDown)? _textBox.SelectionStart + _textBox.SelectionLength:Math.Max(0, _textBox.SelectionStart-1));
            int searchIndex = (ToDown)?( 
                targetString.IndexOf(searchString, searchStartIndex
                , (Capital) ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase)):
                (targetString.LastIndexOf(searchString, searchStartIndex
                , (Capital) ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase));
            if (searchIndex == -1 ||(!ToDown && _textBox.SelectionStart==0))
            {
                System.Media.SystemSounds.Beep.Play();
            }
            else
            {
                _textBox.Select(searchIndex, searchString.Length);
                _textBox.ScrollToCaret();
                searchStartIndex = searchIndex + searchString.Length;
                _textBox.Focus();
            }
        }
        public static void SearchRegularExpression
            (TextBoxBase tbox, string Original, bool ToDown = true, bool fromTop = false,bool Case=false)
        {
            TextBoxBase _textBox = tbox;
            string searchString = tbox.Text;
            try
            {
                System.Text.RegularExpressions.Regex r =
                    (Case) ?
                    new System.Text.RegularExpressions.Regex(Original) :
                    new System.Text.RegularExpressions.Regex(Original, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                int searchStartIndex = ((fromTop) ?
                    ((ToDown) ? 0 : _textBox.Text.Length) :
                    (ToDown) ? _textBox.SelectionStart + _textBox.SelectionLength : Math.Max(0, _textBox.SelectionStart - 1));
                System.Text.RegularExpressions.Match m = r.Match(searchString, searchStartIndex);
                int searchIndex = m.Index;
                if (!m.Success || (!ToDown && _textBox.SelectionStart == 0))
                {
                    System.Media.SystemSounds.Beep.Play();
                }
                else
                {
                    _textBox.Select(searchIndex, m.Length);
                    _textBox.ScrollToCaret();
                    searchStartIndex = searchIndex + m.Length;
                    _textBox.Focus();
                }
            }
            catch
            {
                MessageBox.Show("入力された文字列は正規表現として適切ではありません。\n詳しくはヘルプを参照してください。"
                    , "正規表現エラー", MessageBoxButtons.OK);
            }
        }

        private void ResetSearchPosition()
        {
            _textBox.SelectionStart = 0;
            _textBox.Select(0, 0);
            _textBox.Focus();
        }

        public TextBoxBase TextBoxBase() { return _textBox; }
        public void SetTextBoxBase(TextBoxBase tb) { _textBox=tb; }

        public dialogMode Mode() { return _mode; }
        public void SetMode(dialogMode mode) { _mode = mode; }

        public void Set(TextBoxBase rt, dialogMode md) { SetTextBoxBase(rt); SetMode(md); }

        private void button5_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void ChangeDialogMode()
        {
            if (_mode == dialogMode.Find) { ChangeDialogMode(dialogMode.Replace); }
            else{ChangeDialogMode(dialogMode.Find);}
        }

        private void ChangeDialogMode(dialogMode md)
        {
            if (md == dialogMode.Find)
            {
                panel1.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
            }
            else
            {
                panel1.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
            _mode = md;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ChangeDialogMode();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _searchString = textBox1.Text;
            DoSearch(fromTop: true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _textBox.Focus();
            _searchString = textBox1.Text;
            DoSearch(fromTop: false);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (_searchString != null && _textBox.SelectedText.ToLower() == _searchString.ToLower()) { _textBox.SelectedText = textBox2.Text; }
            _searchString = textBox1.Text;
            DoSearch(fromTop: false);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _textBox.Text = _textBox.Text.Replace(textBox1.Text, textBox2.Text);
        }
        #region Event//イベント発生
        public delegate void SearchStatusChangedEventHandler(object sender,SearchStatusChangedEventArgs e);
     
        public event SearchStatusChangedEventHandler SearchStatusChanged;
        protected virtual void OnSearchStatusChanged(SearchStatusChangedEventArgs e)
        {
            if (SearchStatusChanged != null) { SearchStatusChanged(this, e); }
        }

        public class SearchStatusChangedEventArgs : EventArgs
        {
            private readonly string text;
            private readonly string reptext;
            private readonly dialogMode mode;
            private readonly bool cap;
            private readonly bool esc;
            private readonly bool reg;
            private readonly bool down;

            public SearchStatusChangedEventArgs(string text, dialogMode mode, bool cap, bool esc, bool reg, bool down,string reptext)
            {
                this.text = text;
                this.reptext = reptext;
                this.mode = mode;
                this.cap = cap;
                this.esc = esc;
                this.reg = reg;
                this.down = down;
            }

            public string SearchText() { return text; }
            public string ReplaceText() { return reptext; }
            public dialogMode Mode() { return mode; }
            public bool IgnoreCase() { return cap; }
            public bool Escape() { return esc; }
            public bool RegularExpression() { return reg; }
            public bool SearchForward() { return down; }
            public bool SearchBackward() { return !down; }
        }
        #endregion
    }
}
