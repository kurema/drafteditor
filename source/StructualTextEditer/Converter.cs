using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace StructualTextEditer
{
    public partial class Converter
    {
        #region XHTML//XHTML出力部

        public struct XhtmlReturn
        {
            public List<XhtmlInfo> Xhtml;
            public List<FileInfo> Images;
        }

        public struct FileInfo
        {
            public string FileName;
            public string mimetype;
            public string ID;
        }

        public static XhtmlReturn Xhtml(string text,int basicDepth=1,string css=null,string imageDirectory=null,string textDirectory=null)
        {
            //imageへの対応はあまり美しくない。
            StructualText.StructualTextReader sr = new StructualText.StructualTextReader(text);
            string ret = "";
            string title="";
            int imageNum = 0;
            if (textDirectory == null)
            {
                textDirectory = System.IO.Directory.GetCurrentDirectory();
            }
            if (imageDirectory != null)
            {
                if (!System.IO.Directory.Exists(imageDirectory)) { System.IO.Directory.CreateDirectory(imageDirectory); }
                if (!imageDirectory.EndsWith("/")) { imageDirectory += "/"; }
            }
            List<XhtmlInfo> srl = new List<XhtmlInfo>();
            List<FileInfo> ims = new List<FileInfo>();
            string tempCD= System.IO.Directory.GetCurrentDirectory();
            System.IO.Directory.SetCurrentDirectory(textDirectory);
            while (sr.ReadLine())
            {
                System.IO.Directory.SetCurrentDirectory(tempCD);
                if (sr.Type == StructualText.StructualTextReader.LineType.Comment)
                {
                    ret += "<!-- " + XhtmlEscape(sr.Text) + " -->\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Text)
                {
                    ret += "<p>" + XhtmlEscape(sr.Text) + "</p>\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Math)
                {
                    ret += "<p class=\"Math\">" + XhtmlEscape(sr.Text) + "</p>\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Image)
                {
                    if (imageDirectory == null)
                    {
                        ret += "<img src=\"" + sr.Text + "\" alt=\"Image\" />";
                    }
                    else
                    {
                        string imagePath=imageDirectory +"image"+ imageNum + System.IO.Path.GetExtension(sr.Text);
                        System.IO.File.Copy(sr.Text,imagePath );
                        ret += "<img src=\"" + imagePath + "\" alt=\"Image" + imageNum + "\" /><br />\n";

                        ims.Add(new FileInfo() { FileName = imagePath, mimetype = GetImageMimetype(imagePath),ID="Image"+imageNum });
                        imageNum++;
                    }
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Structure)
                {
                    if (sr.Depth >= basicDepth && sr.Depth < basicDepth + 6)
                    {
                        ret += "<h" + (sr.Depth - basicDepth+1) + ">" + XhtmlEscape(sr.Text) + "</h" + (sr.Depth - basicDepth+1) + ">\n";
                    }
                    else if (sr.Depth >= basicDepth + 6)
                    {
                        ret += "<h6>" + XhtmlEscape(sr.Text) + "</h6>\n";
                    }
                    else if(title!="")
                    {
                        srl.Add(new XhtmlInfo(XhtmlPackage(ret, title, css), title));
                        title = sr.Text;
                        ret = "";
                    }
                    if (sr.Depth == 1 && title=="")
                    {
                        title = sr.Text;
                    }
                }
                System.IO.Directory.SetCurrentDirectory(textDirectory);
            }
            System.IO.Directory.SetCurrentDirectory(tempCD);

            srl.Add(new XhtmlInfo(XhtmlPackage(ret, title, css), title));
            return new XhtmlReturn()
            {
                Xhtml = srl,
                Images = ims
            };
            
        }
        private static string XhtmlEscape(string text)
        {
            return text.Replace("<", "&lt;").Replace(">", "&gt;").Replace("&", "&amp;").Replace(" ", "&nbsp;").Replace("\"", "&quot;");
        }
        private static string XhtmlPackage(string text,string title,string css=null)
        {
            string cssNote = "";
            if(css!=null){cssNote="<link rel=\"stylesheet\" type=\"text/css\" href=\""+css+"\" />\n";}
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.1//EN\" " +
                    "\"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd\">\n" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\" xml:lang=\"ja\">\n" +
            "<head>\n <title>" + title + "</title>\n"+cssNote+"</head><body>\n" + text + "\n</body></html>";
            //lang="ja"を省略。何らかの問題もあり得る。
        }

        public class XhtmlInfo:CommonInfo
        {
            public XhtmlInfo() : base() { }
            public XhtmlInfo(string text) : base(text) { }
            public XhtmlInfo(string text, string title) : base(text, title) { }
            public XhtmlInfo(string text, string title, string filename) : base(text, title, filename) { }
        }

        #endregion

        #region ePub//ePub出力部 zipライブラリと外部ファイル依存
        private static XmlAttribute NewAttribute(XmlDocument xd, string name, string value)
        {
            XmlAttribute txa = xd.CreateAttribute(name);
            txa.Value = value;
            return txa;
        }
        
        public static void SaveEPub(string text,string filename)
        {

            string relBasePath = "Temporary";
            string basePath = System.IO.Path.GetDirectoryName(filename) + "/"+relBasePath+"/";
            int tempNum=0;
            while(System.IO.Directory.Exists(basePath)){
                relBasePath = "Temporary"+tempNum;
                basePath = System.IO.Path.GetDirectoryName(filename) + "/" + relBasePath + "/";
                tempNum++;
            }
            const string DATAROOT = "./Template/ePub/"; 
            
            System.IO.Directory.CreateDirectory(basePath);
            try
            {
                System.IO.Directory.CreateDirectory(basePath + "OEBPS");
                string tempCD = System.IO.Directory.GetCurrentDirectory();
                System.IO.Directory.SetCurrentDirectory(basePath+"OEBPS");
                XhtmlReturn xret = Xhtml(text, 2, "default.css", "image");
                System.IO.Directory.SetCurrentDirectory(tempCD);

                List<XhtmlInfo> xhtmls = xret.Xhtml;
                List<FileInfo> images = xret.Images;

                for (int i = 0; i < xhtmls.Count; i++)
                {
                    string tempfn = basePath + "OEBPS/" + i.ToString() + ".xhtml";
                    using (System.IO.StreamWriter sw = new System.IO.StreamWriter(tempfn, false, Encoding.UTF8))
                    {
                        sw.Write(xhtmls[i].Text);
                        xhtmls[i].FileName = tempfn;
                        xhtmls[i].ID = "Section" + i;
                    }
                }

                XmlDocument xd = new XmlDocument();
                xd.Load(DATAROOT + "content.opf");
                System.Xml.XmlNamespaceManager xmlnsManager = new System.Xml.XmlNamespaceManager(xd.NameTable);
                xmlnsManager.AddNamespace("basic", "http://www.idpf.org/2007/opf");
                xmlnsManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
                xmlnsManager.AddNamespace("opf", "http://www.idpf.org/2007/opf");
                //            xd.SelectSingleNode("/basic:package/basic:metadata/dc:title").InnerText = "タイトル";
                //            xd.SelectSingleNode("/basic:package/basic:metadata/dc:creater").InnerText = "作者名";
                xd.SelectSingleNode("/basic:package/basic:metadata/dc:identifier", xmlnsManager).InnerText = "Guid:" + Guid.NewGuid();
                XmlNode txn1 = xd.SelectSingleNode("/basic:package/basic:manifest",xmlnsManager);
                XmlNode txn2 = xd.SelectSingleNode("/basic:package/basic:spine",xmlnsManager);
                foreach (XhtmlInfo xi in xhtmls)
                {
                    XmlNode txn11 = xd.CreateNode(XmlNodeType.Element, "item", "http://www.idpf.org/2007/opf");
                    txn11.Attributes.Append(NewAttribute(xd, "href", "OEBPS/" + System.IO.Path.GetFileName(xi.FileName)));
                    txn11.Attributes.Append(NewAttribute(xd,"id",xi.ID));
                    txn11.Attributes.Append(NewAttribute(xd, "media-type", "application/xhtml+xml"));
                    txn1.AppendChild(txn11);
                    XmlNode txn21 = xd.CreateNode(XmlNodeType.Element, "itemref", "http://www.idpf.org/2007/opf");
                    txn21.Attributes.Append(NewAttribute(xd, "idref", xi.ID));
                    txn2.AppendChild(txn21);
                }
                foreach (FileInfo xi in images)
                {
                    XmlNode txn11 = xd.CreateNode(XmlNodeType.Element, "item", "http://www.idpf.org/2007/opf");
                    txn11.Attributes.Append(NewAttribute(xd, "href", "OEBPS"+ xi.FileName));
                    txn11.Attributes.Append(NewAttribute(xd, "id", xi.ID));
                    txn11.Attributes.Append(NewAttribute(xd, "media-type", xi.mimetype));
                    txn1.AppendChild(txn11);
                }
                xd.Save(basePath + "content.opf");

                xd = new XmlDocument();
                xd.Load(DATAROOT + "content.ncx");
                xmlnsManager = new System.Xml.XmlNamespaceManager(xd.NameTable);
                xmlnsManager.AddNamespace("basic", "http://www.daisy.org/z3986/2005/ncx/");
                //            xd.SelectSingleNode("/basic:ncx/basic:docTitle/basic:Text").InnerText = "タイトル";
                //            xd.SelectSingleNode("/basic:ncx/basic:docAuthor/basic:Text").InnerText = "タイトル";
                XmlNode txn3 = xd.SelectSingleNode("/basic:ncx/basic:navMap",xmlnsManager);
                for (int i = 0; i < xhtmls.Count; i++)
                {
                    XhtmlInfo xi = xhtmls[i];
                    XmlNode txn31 = xd.CreateNode(XmlNodeType.Element, "navPoint", "http://www.daisy.org/z3986/2005/ncx/");
                    txn31.Attributes.Append(NewAttribute(xd, "playOrder", (i+1).ToString()));
                    txn31.Attributes.Append(NewAttribute(xd, "id", xi.ID));
                    XmlNode txn311 = xd.CreateNode(XmlNodeType.Element, "navLabel", "http://www.daisy.org/z3986/2005/ncx/");
                    XmlNode txn3111 = xd.CreateNode(XmlNodeType.Element, "text", "http://www.daisy.org/z3986/2005/ncx/");
                    txn3111.InnerText = xi.Title;
                    txn311.AppendChild(txn3111);
                    XmlNode txn312 = xd.CreateNode(XmlNodeType.Element, "content", "http://www.daisy.org/z3986/2005/ncx/");
                    txn312.Attributes.Append(NewAttribute(xd, "src", "OEBPS/" + System.IO.Path.GetFileName(xi.FileName)));
                    txn31.AppendChild(txn311);
                    txn31.AppendChild(txn312);
                    txn3.AppendChild(txn31);
                }
                xd.Save(basePath + "content.ncx");


                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.ProvisionalAlternateEncoding = System.Text.Encoding.GetEncoding("shift_jis");
                    zip.UseUnicodeAsNecessary = true;
                    zip.EmitTimesInWindowsFormatWhenSaving = false;
                    zip.EmitTimesInUnixFormatWhenSaving = false;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.None;
                    zip.AddFile(DATAROOT + "mimetype","");
                    zip.EmitTimesInWindowsFormatWhenSaving = true;
                    zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression;
                    zip.AddFile(DATAROOT + "container.xml", "META-INF");
                    zip.AddDirectory(basePath + "OEBPS","OEBPS");
                    zip.AddFile(DATAROOT + "default.css", "OEBPS");
                    zip.AddFile(basePath + "content.ncx", "");
                    zip.AddFile(basePath + "content.opf", "");

                    zip.Save(filename);
                }
            }
            catch (EntryPointNotFoundException e)
            {
                throw e;
            }
            finally
            {
                System.IO.Directory.Delete(basePath,true);
            }
        }

        #endregion

        #region Tex//LaTex用
        public enum TexType
        {
            article,report,book,
            jarticle,jreport,jbook
            ,jarticleCustom,jreportCustom
            ,tarticle,tbook
        }

        public static TexInfo Tex(string text,TexType type=TexType.jreport)
        {
            string header = "";
            string body="";
            string title="";
            if(type==TexType.article){
                header="\\documentclass{article}\n";
            }else if (type == TexType.report)
            {
                header = "\\documentclass{report}\n";
            }else if (type == TexType.book)
            {
                header = "\\documentclass{book}\n";
            }else if(type==TexType.jarticle){
                header="\\documentclass{jarticle}\n";
            }else if (type == TexType.jreport)
            {
                header = "\\documentclass{jreport}\n";
            }else if (type == TexType.jbook)
            {
                header = "\\documentclass{jbook}\n";
            }else if(type==TexType.tarticle){
                header="\\documentclass{tarticle}\n";
            }else if (type == TexType.tbook)
            {
                header = "\\documentclass{tbook}\n";
            }
            else if (type == TexType.jarticleCustom)
            {
                header = "\\documentclass{jarticle}\n\\setlength{\\textwidth}{17cm}\n\\setlength{\\textheight}{25cm}\n\\setlength{\\topmargin}{-1cm}\n\\setlength{\\oddsidemargin}{0cm}\n\\setlength{\\evensidemargin}{0cm}\n";
            }
            else if (type == TexType.jreportCustom)
            {
                header = "\\documentclass{jreport}\n\\setlength{\\textwidth}{17cm}\n\\setlength{\\textheight}{25cm}\n\\setlength{\\topmargin}{-1cm}\n\\setlength{\\oddsidemargin}{0cm}\n\\setlength{\\evensidemargin}{0cm}\n";
            }

            StructualText.StructualTextReader sr = new StructualText.StructualTextReader(text);
            while (sr.ReadLine())
            {
                if (sr.Type == StructualText.StructualTextReader.LineType.Comment)
                {
                    body+="%"+texEscape( sr.Text)+"\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Text)
                {
                    body+=texEscape(sr.Text)+"\n\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Math)
                {
                    body+="$$"+sr.Text+"$$\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Image)
                {
                    body += "%自動コメント:ここには本来画像が挿入されます。("+sr.Text+")\n";
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Structure)
                {
                    if(sr.Depth==1 && title==""){title=sr.Text;}
                    else if(sr.Depth==1){//ToDo
//                        body+="\\part{"+texEscape(sr.Text)+"}\n";
                    }
                    if(type==TexType.jarticle || type==TexType.jarticleCustom|| type==TexType.article){
                        if(sr.Depth==1){
                            body+="\\part{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==2){
                            body+="\\section{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==3){
                            body+="\\subsection{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==4){
                            body+="\\subsubsection{"+texEscape(sr.Text)+"}\n";
                        }
                        else if (sr.Depth == 5)
                        {
                            body += "\\paragraph{" + texEscape(sr.Text) + "}\n";
                        }
                        else if (sr.Depth >= 6)
                        {
                            body += "\\subparagraph{" + texEscape(sr.Text) + "}\n";
                        }
                    }else{
                        if(sr.Depth==1){
                            body+="\\part{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==2){
                            body+="\\chapter{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==3){
                            body+="\\section{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==4){
                            body+="\\subsection{"+texEscape(sr.Text)+"}\n";
                        }else if(sr.Depth==5){
                            body+="\\paragraph{"+texEscape(sr.Text)+"}\n";
                        }
                        else if (sr.Depth >= 6)
                        {
                            body += "\\subparagraph{" + texEscape(sr.Text) + "}\n";
                        }
                    }
                }
            }
            return new TexInfo( header+"\\title{"+title+"}\n\\date{"+DateTime.Now.ToString()+"}\n\\begin{document}\n\\maketitle\n"+body+
                "\\end{document}\n",title);
        }

        public class TexInfo : CommonInfo
        {
            public TexInfo():base(){}
            public TexInfo(string text) : base(text) { }
            public TexInfo(string text, string title) : base(text, title) { }
            public TexInfo(string text, string title,string filename) : base(text, title,filename) { }
        }

        private static string texEscape(string text){
            return text.Replace(@"#",@"\#").Replace(@"%",@"\%").Replace(@"_",@"\_")
                .Replace(@"{",@"\{").Replace(@">",@"\>").Replace(@"$",@"\$")
                .Replace(@"&",@"\&").Replace(@"}",@"\}")
                .Replace(@"<",@"\verb|<|").Replace(@"\",@"\verb|\|").Replace(@"^",@"\verb|^|")
                .Replace(@"~",@"\verb|~|").Replace(@"|",@"\verb+|+");
        }
#endregion

        #region XML//XML

        static public XmlDocument XML(string text)
        {
            StructualText.StructualTextReader sr = new StructualText.StructualTextReader(text);
            XmlDocument xd = new XmlDocument();
            const string xmlns = "http://www.structualtext.co.jp/2011/xml";

            List<XmlNode> chapters = new List<XmlNode>();

            XmlDeclaration xmldec = xd.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            xd.AppendChild(xmldec);

            XmlNode topxn = xd.CreateElement("document", xmlns);
            chapters.Add(topxn);
            xd.AppendChild(topxn);


            XmlNode currentxn = topxn;

            while (sr.ReadLine())
            {
                if (sr.Type == StructualText.StructualTextReader.LineType.Comment)
                {
                    XmlComment tempxn = xd.CreateComment(sr.Text);
                    currentxn.AppendChild(tempxn);
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Text)
                {
                    XmlElement tempxn = xd.CreateElement("paragraph", xmlns);
                    tempxn.InnerText = sr.Text;
                    currentxn.AppendChild(tempxn);
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Math)
                {
                    XmlElement tempxn = xd.CreateElement("mathematics", xmlns);
                    XmlElement tempxnme = xd.CreateElement("expression", xmlns);
                    tempxnme.InnerText = sr.Text;
                    XmlAttribute xi = xd.CreateAttribute("type");
                    xi.Value = "text";//tex,mathML等。
                    tempxnme.Attributes.Append(xi);
                    tempxn.AppendChild(tempxnme);
                    currentxn.AppendChild(tempxn);
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Image)
                {
                    XmlElement tempxn = xd.CreateElement("image", xmlns);
                    XmlAttribute xi = xd.CreateAttribute("src");
                    xi.Value = sr.Text;
                    XmlAttribute xim = xd.CreateAttribute("mimetype");
                    xim.Value = GetImageMimetype(sr.Text);
                    tempxn.Attributes.Append(xi);
                    tempxn.Attributes.Append(xim);
                    currentxn.AppendChild(tempxn);
                }
                else if (sr.Type == StructualText.StructualTextReader.LineType.Structure)
                {
                    for (int i = chapters.Count; i < sr.Depth; i++)
                    {
                        XmlElement xe = xd.CreateElement("chapter", xmlns);
                        //XmlAttribute xi = xd.CreateAttribute("level");
                        //xi.Value = i.ToString();
                        //xe.Attributes.Append(xi);
                        chapters[i - 1].AppendChild(xe);
                        chapters.Add(xe);
                    }
                    XmlElement xc = xd.CreateElement("chapter", xmlns);
                    //XmlAttribute xl = xd.CreateAttribute("level");
                    //xl.Value = sr.Depth.ToString();
                    XmlAttribute xt = xd.CreateAttribute("title");
                    xt.Value = sr.Text;
                    xc.Attributes.Append(xt);
                    //xc.Attributes.Append(xl);
                    chapters[sr.Depth - 1].AppendChild(xc);
                    if (sr.Depth >= chapters.Count)
                    {
                        chapters.Add(xc);
                    }
                    else
                    {
                        chapters[sr.Depth] = xc;
                    }
                    currentxn = xc;
                }

            }
            return xd;
        }



        #endregion

        public class CommonInfo
        {
            public string Text;
            public string Title;
            public string FileName;
            public string ID;

            public static implicit operator string(CommonInfo xi) { return xi.Text; }

            public CommonInfo()
            {
                this.Text = "";
                this.Title = "";
                this.FileName = "";
                this.ID = "";
            }
            public CommonInfo(string text)
            {
                this.Text = text;
                this.Title = "";
                this.FileName = "";
                this.ID = "";
            }
            public CommonInfo(string text, string title)
            {
                this.Text = text;
                this.Title = title;
                this.FileName = "";
                this.ID = "";
            }
            public CommonInfo(string text, string title,string filename)
            {
                this.Text = text;
                this.Title = title;
                this.FileName = filename;
                this.ID = "";
            }
        }

        static public string GetImageMimetype(string path)
        {
            Dictionary<string, string> mimetypeDic = new Dictionary<string, string>();
            mimetypeDic[".png"] = "image/png";
            mimetypeDic[".jpeg"] = "image/jpeg";
            mimetypeDic[".jpg"] = "image/jpeg";
            mimetypeDic[".gif"] = "image/gif";
            mimetypeDic[".svg"] = "image/svg+xml";
            mimetypeDic[".bmp"] = "image/bmp";
            mimetypeDic[".tiff"] = "image/tiff";

            return mimetypeDic.ContainsKey(System.IO.Path.GetExtension(path).ToLower()) ?
                mimetypeDic[System.IO.Path.GetExtension(path).ToLower()] :
                "image/" + System.IO.Path.GetExtension(path).ToLower().Substring(1)
                ;
        }
    }
}
