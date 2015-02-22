using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//OpenXML Related
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using M = DocumentFormat.OpenXml.Math;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace StructualTextEditer
{
    static partial class Converter
    {
        static public void SaveOfficeOpenXML(string text,string fileName)
        {
            using (var package = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                DocumentFormat.OpenXml.Packaging.MainDocumentPart maindoc = package.AddMainDocumentPart();
                OfficeOpenXMLDocument(text,maindoc);

                StyleDefinitionsPart styleDefinitionsPart1 = maindoc.AddNewPart<StyleDefinitionsPart>("rId1");
                styleDefinitionsPart1.Styles = OfficeOpenXMLDefaultStyle();
            }
        }

        static public Styles OfficeOpenXMLDefaultStyle()
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts5 = new RunFonts() { ComplexScript = "Times New Roman", AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia };
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };
            Languages languages3 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "en-US" };

            runPropertiesBaseStyle1.Append(runFonts5);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages3);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35 };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid5 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines2);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid5);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            Style style2 = new Style() { Type = StyleValues.Paragraph, StyleId = "1" };
            StyleName styleName2 = new StyleName() { Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "10" };
            UIPriority uIPriority1 = new UIPriority() { Val = 9 };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid6 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties2.Append(keepNext1);
            styleParagraphProperties2.Append(spacingBetweenLines3);
            styleParagraphProperties2.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Kern kern1 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize3 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties2.Append(runFonts6);
            styleRunProperties2.Append(bold1);
            styleRunProperties2.Append(boldComplexScript1);
            styleRunProperties2.Append(kern1);
            styleRunProperties2.Append(fontSize3);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(uIPriority1);
            style2.Append(primaryStyle2);
            style2.Append(rsid6);
            style2.Append(styleParagraphProperties2);
            style2.Append(styleRunProperties2);

            Style style3 = new Style() { Type = StyleValues.Paragraph, StyleId = "2" };
            StyleName styleName3 = new StyleName() { Val = "heading 2" };
            BasedOn basedOn2 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "20" };
            UIPriority uIPriority2 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid7 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties3.Append(keepNext2);
            styleParagraphProperties3.Append(spacingBetweenLines4);
            styleParagraphProperties3.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts7 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize4 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties3.Append(runFonts7);
            styleRunProperties3.Append(bold2);
            styleRunProperties3.Append(boldComplexScript2);
            styleRunProperties3.Append(italic1);
            styleRunProperties3.Append(italicComplexScript1);
            styleRunProperties3.Append(fontSize4);
            styleRunProperties3.Append(fontSizeComplexScript4);

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(nextParagraphStyle2);
            style3.Append(linkedStyle2);
            style3.Append(uIPriority2);
            style3.Append(semiHidden1);
            style3.Append(unhideWhenUsed1);
            style3.Append(primaryStyle3);
            style3.Append(rsid7);
            style3.Append(styleParagraphProperties3);
            style3.Append(styleRunProperties3);

            Style style4 = new Style() { Type = StyleValues.Paragraph, StyleId = "3" };
            StyleName styleName4 = new StyleName() { Val = "heading 3" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "30" };
            UIPriority uIPriority3 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid8 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            KeepNext keepNext3 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties4.Append(keepNext3);
            styleParagraphProperties4.Append(spacingBetweenLines5);
            styleParagraphProperties4.Append(outlineLevel3);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts8 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize5 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties4.Append(runFonts8);
            styleRunProperties4.Append(bold3);
            styleRunProperties4.Append(boldComplexScript3);
            styleRunProperties4.Append(fontSize5);
            styleRunProperties4.Append(fontSizeComplexScript5);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(nextParagraphStyle3);
            style4.Append(linkedStyle3);
            style4.Append(uIPriority3);
            style4.Append(semiHidden2);
            style4.Append(unhideWhenUsed2);
            style4.Append(primaryStyle4);
            style4.Append(rsid8);
            style4.Append(styleParagraphProperties4);
            style4.Append(styleRunProperties4);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "4" };
            StyleName styleName5 = new StyleName() { Val = "heading 4" };
            BasedOn basedOn4 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "40" };
            UIPriority uIPriority4 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid9 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            KeepNext keepNext4 = new KeepNext();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 3 };

            styleParagraphProperties5.Append(keepNext4);
            styleParagraphProperties5.Append(spacingBetweenLines6);
            styleParagraphProperties5.Append(outlineLevel4);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts9 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            FontSize fontSize6 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties5.Append(runFonts9);
            styleRunProperties5.Append(bold4);
            styleRunProperties5.Append(boldComplexScript4);
            styleRunProperties5.Append(fontSize6);
            styleRunProperties5.Append(fontSizeComplexScript6);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(nextParagraphStyle4);
            style5.Append(linkedStyle4);
            style5.Append(uIPriority4);
            style5.Append(semiHidden3);
            style5.Append(unhideWhenUsed3);
            style5.Append(primaryStyle5);
            style5.Append(rsid9);
            style5.Append(styleParagraphProperties5);
            style5.Append(styleRunProperties5);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "5" };
            StyleName styleName6 = new StyleName() { Val = "heading 5" };
            BasedOn basedOn5 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "50" };
            UIPriority uIPriority5 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid10 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel5 = new OutlineLevel() { Val = 4 };

            styleParagraphProperties6.Append(spacingBetweenLines7);
            styleParagraphProperties6.Append(outlineLevel5);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts10 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            FontSize fontSize7 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties6.Append(runFonts10);
            styleRunProperties6.Append(bold5);
            styleRunProperties6.Append(boldComplexScript5);
            styleRunProperties6.Append(italic2);
            styleRunProperties6.Append(italicComplexScript2);
            styleRunProperties6.Append(fontSize7);
            styleRunProperties6.Append(fontSizeComplexScript7);

            style6.Append(styleName6);
            style6.Append(basedOn5);
            style6.Append(nextParagraphStyle5);
            style6.Append(linkedStyle5);
            style6.Append(uIPriority5);
            style6.Append(semiHidden4);
            style6.Append(unhideWhenUsed4);
            style6.Append(primaryStyle6);
            style6.Append(rsid10);
            style6.Append(styleParagraphProperties6);
            style6.Append(styleRunProperties6);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "6" };
            StyleName styleName7 = new StyleName() { Val = "heading 6" };
            BasedOn basedOn6 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "60" };
            UIPriority uIPriority6 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden5 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            Rsid rsid11 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel6 = new OutlineLevel() { Val = 5 };

            styleParagraphProperties7.Append(spacingBetweenLines8);
            styleParagraphProperties7.Append(outlineLevel6);

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts11 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize8 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties7.Append(runFonts11);
            styleRunProperties7.Append(bold6);
            styleRunProperties7.Append(boldComplexScript6);
            styleRunProperties7.Append(fontSize8);
            styleRunProperties7.Append(fontSizeComplexScript8);

            style7.Append(styleName7);
            style7.Append(basedOn6);
            style7.Append(nextParagraphStyle6);
            style7.Append(linkedStyle6);
            style7.Append(uIPriority6);
            style7.Append(semiHidden5);
            style7.Append(unhideWhenUsed5);
            style7.Append(primaryStyle7);
            style7.Append(rsid11);
            style7.Append(styleParagraphProperties7);
            style7.Append(styleRunProperties7);

            Style style8 = new Style() { Type = StyleValues.Paragraph, StyleId = "7" };
            StyleName styleName8 = new StyleName() { Val = "heading 7" };
            BasedOn basedOn7 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "70" };
            UIPriority uIPriority7 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            Rsid rsid12 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel7 = new OutlineLevel() { Val = 6 };

            styleParagraphProperties8.Append(spacingBetweenLines9);
            styleParagraphProperties8.Append(outlineLevel7);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts12 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };

            styleRunProperties8.Append(runFonts12);

            style8.Append(styleName8);
            style8.Append(basedOn7);
            style8.Append(nextParagraphStyle7);
            style8.Append(linkedStyle7);
            style8.Append(uIPriority7);
            style8.Append(semiHidden6);
            style8.Append(unhideWhenUsed6);
            style8.Append(primaryStyle8);
            style8.Append(rsid12);
            style8.Append(styleParagraphProperties8);
            style8.Append(styleRunProperties8);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "8" };
            StyleName styleName9 = new StyleName() { Val = "heading 8" };
            BasedOn basedOn8 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle8 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "80" };
            UIPriority uIPriority8 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden7 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle9 = new PrimaryStyle();
            Rsid rsid13 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel8 = new OutlineLevel() { Val = 7 };

            styleParagraphProperties9.Append(spacingBetweenLines10);
            styleParagraphProperties9.Append(outlineLevel8);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic3 = new Italic();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();

            styleRunProperties9.Append(runFonts13);
            styleRunProperties9.Append(italic3);
            styleRunProperties9.Append(italicComplexScript3);

            style9.Append(styleName9);
            style9.Append(basedOn8);
            style9.Append(nextParagraphStyle8);
            style9.Append(linkedStyle8);
            style9.Append(uIPriority8);
            style9.Append(semiHidden7);
            style9.Append(unhideWhenUsed7);
            style9.Append(primaryStyle9);
            style9.Append(rsid13);
            style9.Append(styleParagraphProperties9);
            style9.Append(styleRunProperties9);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "9" };
            StyleName styleName10 = new StyleName() { Val = "heading 9" };
            BasedOn basedOn9 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle9 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "90" };
            UIPriority uIPriority9 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden8 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle10 = new PrimaryStyle();
            Rsid rsid14 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Before = "240", After = "60" };
            OutlineLevel outlineLevel9 = new OutlineLevel() { Val = 8 };

            styleParagraphProperties10.Append(spacingBetweenLines11);
            styleParagraphProperties10.Append(outlineLevel9);

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            FontSize fontSize9 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties10.Append(runFonts14);
            styleRunProperties10.Append(fontSize9);
            styleRunProperties10.Append(fontSizeComplexScript9);

            style10.Append(styleName10);
            style10.Append(basedOn9);
            style10.Append(nextParagraphStyle9);
            style10.Append(linkedStyle9);
            style10.Append(uIPriority9);
            style10.Append(semiHidden8);
            style10.Append(unhideWhenUsed8);
            style10.Append(primaryStyle10);
            style10.Append(rsid14);
            style10.Append(styleParagraphProperties10);
            style10.Append(styleRunProperties10);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "a0", Default = true };
            StyleName styleName11 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority10 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden9 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();

            style11.Append(styleName11);
            style11.Append(uIPriority10);
            style11.Append(semiHidden9);
            style11.Append(unhideWhenUsed9);

            Style style12 = new Style() { Type = StyleValues.Table, StyleId = "a1", Default = true };
            StyleName styleName12 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority11 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden10 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle11 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style12.Append(styleName12);
            style12.Append(uIPriority11);
            style12.Append(semiHidden10);
            style12.Append(unhideWhenUsed10);
            style12.Append(primaryStyle11);
            style12.Append(styleTableProperties1);

            Style style13 = new Style() { Type = StyleValues.Numbering, StyleId = "a2", Default = true };
            StyleName styleName13 = new StyleName() { Val = "No List" };
            UIPriority uIPriority12 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden11 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed11 = new UnhideWhenUsed();

            style13.Append(styleName13);
            style13.Append(uIPriority12);
            style13.Append(semiHidden11);
            style13.Append(unhideWhenUsed11);

            Style style14 = new Style() { Type = StyleValues.Paragraph, StyleId = "a3" };
            StyleName styleName14 = new StyleName() { Val = "Title" };
            BasedOn basedOn10 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle10 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "a4" };
            UIPriority uIPriority13 = new UIPriority() { Val = 10 };
            PrimaryStyle primaryStyle12 = new PrimaryStyle();
            Rsid rsid15 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Before = "240", After = "60" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel10 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties11.Append(spacingBetweenLines12);
            styleParagraphProperties11.Append(justification1);
            styleParagraphProperties11.Append(outlineLevel10);

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Kern kern2 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize10 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties11.Append(runFonts15);
            styleRunProperties11.Append(bold7);
            styleRunProperties11.Append(boldComplexScript7);
            styleRunProperties11.Append(kern2);
            styleRunProperties11.Append(fontSize10);
            styleRunProperties11.Append(fontSizeComplexScript10);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(nextParagraphStyle10);
            style14.Append(linkedStyle10);
            style14.Append(uIPriority13);
            style14.Append(primaryStyle12);
            style14.Append(rsid15);
            style14.Append(styleParagraphProperties11);
            style14.Append(styleRunProperties11);

            Style style15 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName15 = new StyleName() { Val = "表題 (文字)" };
            BasedOn basedOn11 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "a3" };
            UIPriority uIPriority14 = new UIPriority() { Val = 10 };
            Rsid rsid16 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold8 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            Kern kern3 = new Kern() { Val = (UInt32Value)28U };
            FontSize fontSize11 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties12.Append(runFonts16);
            styleRunProperties12.Append(bold8);
            styleRunProperties12.Append(boldComplexScript8);
            styleRunProperties12.Append(kern3);
            styleRunProperties12.Append(fontSize11);
            styleRunProperties12.Append(fontSizeComplexScript11);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(linkedStyle11);
            style15.Append(uIPriority14);
            style15.Append(rsid16);
            style15.Append(styleRunProperties12);

            Style style16 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
            StyleName styleName16 = new StyleName() { Val = "Subtitle" };
            BasedOn basedOn12 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle11 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "a6" };
            UIPriority uIPriority15 = new UIPriority() { Val = 11 };
            PrimaryStyle primaryStyle13 = new PrimaryStyle();
            Rsid rsid17 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { After = "60" };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };
            OutlineLevel outlineLevel11 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties12.Append(spacingBetweenLines13);
            styleParagraphProperties12.Append(justification2);
            styleParagraphProperties12.Append(outlineLevel11);

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };

            styleRunProperties13.Append(runFonts17);

            style16.Append(styleName16);
            style16.Append(basedOn12);
            style16.Append(nextParagraphStyle11);
            style16.Append(linkedStyle12);
            style16.Append(uIPriority15);
            style16.Append(primaryStyle13);
            style16.Append(rsid17);
            style16.Append(styleParagraphProperties12);
            style16.Append(styleRunProperties13);

            Style style17 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName17 = new StyleName() { Val = "副題 (文字)" };
            BasedOn basedOn13 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle13 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority16 = new UIPriority() { Val = 11 };
            Rsid rsid18 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties14.Append(runFonts18);
            styleRunProperties14.Append(fontSize12);
            styleRunProperties14.Append(fontSizeComplexScript12);

            style17.Append(styleName17);
            style17.Append(basedOn13);
            style17.Append(linkedStyle13);
            style17.Append(uIPriority16);
            style17.Append(rsid18);
            style17.Append(styleRunProperties14);

            Style style18 = new Style() { Type = StyleValues.Character, StyleId = "10", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "見出し 1 (文字)" };
            BasedOn basedOn14 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle14 = new LinkedStyle() { Val = "1" };
            UIPriority uIPriority17 = new UIPriority() { Val = 9 };
            Rsid rsid19 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold9 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            Kern kern4 = new Kern() { Val = (UInt32Value)32U };
            FontSize fontSize13 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties15.Append(runFonts19);
            styleRunProperties15.Append(bold9);
            styleRunProperties15.Append(boldComplexScript9);
            styleRunProperties15.Append(kern4);
            styleRunProperties15.Append(fontSize13);
            styleRunProperties15.Append(fontSizeComplexScript13);

            style18.Append(styleName18);
            style18.Append(basedOn14);
            style18.Append(linkedStyle14);
            style18.Append(uIPriority17);
            style18.Append(rsid19);
            style18.Append(styleRunProperties15);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "20", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "見出し 2 (文字)" };
            BasedOn basedOn15 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle15 = new LinkedStyle() { Val = "2" };
            UIPriority uIPriority18 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden12 = new SemiHidden();
            Rsid rsid20 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold10 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            Italic italic4 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            FontSize fontSize14 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties16.Append(runFonts20);
            styleRunProperties16.Append(bold10);
            styleRunProperties16.Append(boldComplexScript10);
            styleRunProperties16.Append(italic4);
            styleRunProperties16.Append(italicComplexScript4);
            styleRunProperties16.Append(fontSize14);
            styleRunProperties16.Append(fontSizeComplexScript14);

            style19.Append(styleName19);
            style19.Append(basedOn15);
            style19.Append(linkedStyle15);
            style19.Append(uIPriority18);
            style19.Append(semiHidden12);
            style19.Append(rsid20);
            style19.Append(styleRunProperties16);

            Style style20 = new Style() { Type = StyleValues.Character, StyleId = "30", CustomStyle = true };
            StyleName styleName20 = new StyleName() { Val = "見出し 3 (文字)" };
            BasedOn basedOn16 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle16 = new LinkedStyle() { Val = "3" };
            UIPriority uIPriority19 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden13 = new SemiHidden();
            Rsid rsid21 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold11 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            FontSize fontSize15 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties17.Append(runFonts21);
            styleRunProperties17.Append(bold11);
            styleRunProperties17.Append(boldComplexScript11);
            styleRunProperties17.Append(fontSize15);
            styleRunProperties17.Append(fontSizeComplexScript15);

            style20.Append(styleName20);
            style20.Append(basedOn16);
            style20.Append(linkedStyle16);
            style20.Append(uIPriority19);
            style20.Append(semiHidden13);
            style20.Append(rsid21);
            style20.Append(styleRunProperties17);

            Style style21 = new Style() { Type = StyleValues.Character, StyleId = "40", CustomStyle = true };
            StyleName styleName21 = new StyleName() { Val = "見出し 4 (文字)" };
            BasedOn basedOn17 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle17 = new LinkedStyle() { Val = "4" };
            UIPriority uIPriority20 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden14 = new SemiHidden();
            Rsid rsid22 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold12 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            FontSize fontSize16 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties18.Append(runFonts22);
            styleRunProperties18.Append(bold12);
            styleRunProperties18.Append(boldComplexScript12);
            styleRunProperties18.Append(fontSize16);
            styleRunProperties18.Append(fontSizeComplexScript16);

            style21.Append(styleName21);
            style21.Append(basedOn17);
            style21.Append(linkedStyle17);
            style21.Append(uIPriority20);
            style21.Append(semiHidden14);
            style21.Append(rsid22);
            style21.Append(styleRunProperties18);

            Style style22 = new Style() { Type = StyleValues.Character, StyleId = "50", CustomStyle = true };
            StyleName styleName22 = new StyleName() { Val = "見出し 5 (文字)" };
            BasedOn basedOn18 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle18 = new LinkedStyle() { Val = "5" };
            UIPriority uIPriority21 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden15 = new SemiHidden();
            Rsid rsid23 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold13 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            Italic italic5 = new Italic();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            FontSize fontSize17 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties19.Append(runFonts23);
            styleRunProperties19.Append(bold13);
            styleRunProperties19.Append(boldComplexScript13);
            styleRunProperties19.Append(italic5);
            styleRunProperties19.Append(italicComplexScript5);
            styleRunProperties19.Append(fontSize17);
            styleRunProperties19.Append(fontSizeComplexScript17);

            style22.Append(styleName22);
            style22.Append(basedOn18);
            style22.Append(linkedStyle18);
            style22.Append(uIPriority21);
            style22.Append(semiHidden15);
            style22.Append(rsid23);
            style22.Append(styleRunProperties19);

            Style style23 = new Style() { Type = StyleValues.Character, StyleId = "60", CustomStyle = true };
            StyleName styleName23 = new StyleName() { Val = "見出し 6 (文字)" };
            BasedOn basedOn19 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle19 = new LinkedStyle() { Val = "6" };
            UIPriority uIPriority22 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden16 = new SemiHidden();
            Rsid rsid24 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts24 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Bold bold14 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();

            styleRunProperties20.Append(runFonts24);
            styleRunProperties20.Append(bold14);
            styleRunProperties20.Append(boldComplexScript14);

            style23.Append(styleName23);
            style23.Append(basedOn19);
            style23.Append(linkedStyle19);
            style23.Append(uIPriority22);
            style23.Append(semiHidden16);
            style23.Append(rsid24);
            style23.Append(styleRunProperties20);

            Style style24 = new Style() { Type = StyleValues.Character, StyleId = "70", CustomStyle = true };
            StyleName styleName24 = new StyleName() { Val = "見出し 7 (文字)" };
            BasedOn basedOn20 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle20 = new LinkedStyle() { Val = "7" };
            UIPriority uIPriority23 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden17 = new SemiHidden();
            Rsid rsid25 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts25 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            FontSize fontSize18 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties21.Append(runFonts25);
            styleRunProperties21.Append(fontSize18);
            styleRunProperties21.Append(fontSizeComplexScript18);

            style24.Append(styleName24);
            style24.Append(basedOn20);
            style24.Append(linkedStyle20);
            style24.Append(uIPriority23);
            style24.Append(semiHidden17);
            style24.Append(rsid25);
            style24.Append(styleRunProperties21);

            Style style25 = new Style() { Type = StyleValues.Character, StyleId = "80", CustomStyle = true };
            StyleName styleName25 = new StyleName() { Val = "見出し 8 (文字)" };
            BasedOn basedOn21 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle21 = new LinkedStyle() { Val = "8" };
            UIPriority uIPriority24 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden18 = new SemiHidden();
            Rsid rsid26 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts26 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic6 = new Italic();
            ItalicComplexScript italicComplexScript6 = new ItalicComplexScript();
            FontSize fontSize19 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties22.Append(runFonts26);
            styleRunProperties22.Append(italic6);
            styleRunProperties22.Append(italicComplexScript6);
            styleRunProperties22.Append(fontSize19);
            styleRunProperties22.Append(fontSizeComplexScript19);

            style25.Append(styleName25);
            style25.Append(basedOn21);
            style25.Append(linkedStyle21);
            style25.Append(uIPriority24);
            style25.Append(semiHidden18);
            style25.Append(rsid26);
            style25.Append(styleRunProperties22);

            Style style26 = new Style() { Type = StyleValues.Character, StyleId = "90", CustomStyle = true };
            StyleName styleName26 = new StyleName() { Val = "見出し 9 (文字)" };
            BasedOn basedOn22 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle22 = new LinkedStyle() { Val = "9" };
            UIPriority uIPriority25 = new UIPriority() { Val = 9 };
            SemiHidden semiHidden19 = new SemiHidden();
            Rsid rsid27 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            RunFonts runFonts27 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };

            styleRunProperties23.Append(runFonts27);

            style26.Append(styleName26);
            style26.Append(basedOn22);
            style26.Append(linkedStyle22);
            style26.Append(uIPriority25);
            style26.Append(semiHidden19);
            style26.Append(rsid27);
            style26.Append(styleRunProperties23);

            Style style27 = new Style() { Type = StyleValues.Paragraph, StyleId = "a7" };
            StyleName styleName27 = new StyleName() { Val = "caption" };
            BasedOn basedOn23 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle12 = new NextParagraphStyle() { Val = "a" };
            UIPriority uIPriority26 = new UIPriority() { Val = 35 };
            SemiHidden semiHidden20 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed12 = new UnhideWhenUsed();
            Rsid rsid28 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            Bold bold15 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            Color color1 = new Color() { Val = "4F81BD", ThemeColor = ThemeColorValues.Accent1 };
            FontSize fontSize20 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties24.Append(bold15);
            styleRunProperties24.Append(boldComplexScript15);
            styleRunProperties24.Append(color1);
            styleRunProperties24.Append(fontSize20);
            styleRunProperties24.Append(fontSizeComplexScript20);

            style27.Append(styleName27);
            style27.Append(basedOn23);
            style27.Append(nextParagraphStyle12);
            style27.Append(uIPriority26);
            style27.Append(semiHidden20);
            style27.Append(unhideWhenUsed12);
            style27.Append(rsid28);
            style27.Append(styleRunProperties24);

            Style style28 = new Style() { Type = StyleValues.Character, StyleId = "a8" };
            StyleName styleName28 = new StyleName() { Val = "Strong" };
            BasedOn basedOn24 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority27 = new UIPriority() { Val = 22 };
            PrimaryStyle primaryStyle14 = new PrimaryStyle();
            Rsid rsid29 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            Bold bold16 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();

            styleRunProperties25.Append(bold16);
            styleRunProperties25.Append(boldComplexScript16);

            style28.Append(styleName28);
            style28.Append(basedOn24);
            style28.Append(uIPriority27);
            style28.Append(primaryStyle14);
            style28.Append(rsid29);
            style28.Append(styleRunProperties25);

            Style style29 = new Style() { Type = StyleValues.Character, StyleId = "a9" };
            StyleName styleName29 = new StyleName() { Val = "Emphasis" };
            BasedOn basedOn25 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority28 = new UIPriority() { Val = 20 };
            PrimaryStyle primaryStyle15 = new PrimaryStyle();
            Rsid rsid30 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            RunFonts runFonts28 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi };
            Bold bold17 = new Bold();
            Italic italic7 = new Italic();
            ItalicComplexScript italicComplexScript7 = new ItalicComplexScript();

            styleRunProperties26.Append(runFonts28);
            styleRunProperties26.Append(bold17);
            styleRunProperties26.Append(italic7);
            styleRunProperties26.Append(italicComplexScript7);

            style29.Append(styleName29);
            style29.Append(basedOn25);
            style29.Append(uIPriority28);
            style29.Append(primaryStyle15);
            style29.Append(rsid30);
            style29.Append(styleRunProperties26);

            Style style30 = new Style() { Type = StyleValues.Paragraph, StyleId = "aa" };
            StyleName styleName30 = new StyleName() { Val = "No Spacing" };
            BasedOn basedOn26 = new BasedOn() { Val = "a" };
            UIPriority uIPriority29 = new UIPriority() { Val = 1 };
            PrimaryStyle primaryStyle16 = new PrimaryStyle();
            Rsid rsid31 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties27.Append(fontSizeComplexScript21);

            style30.Append(styleName30);
            style30.Append(basedOn26);
            style30.Append(uIPriority29);
            style30.Append(primaryStyle16);
            style30.Append(rsid31);
            style30.Append(styleRunProperties27);

            Style style31 = new Style() { Type = StyleValues.Paragraph, StyleId = "ab" };
            StyleName styleName31 = new StyleName() { Val = "List Paragraph" };
            BasedOn basedOn27 = new BasedOn() { Val = "a" };
            UIPriority uIPriority30 = new UIPriority() { Val = 34 };
            PrimaryStyle primaryStyle17 = new PrimaryStyle();
            Rsid rsid32 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties13 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation() { Left = "720" };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties13.Append(indentation1);
            styleParagraphProperties13.Append(contextualSpacing1);

            style31.Append(styleName31);
            style31.Append(basedOn27);
            style31.Append(uIPriority30);
            style31.Append(primaryStyle17);
            style31.Append(rsid32);
            style31.Append(styleParagraphProperties13);

            Style style32 = new Style() { Type = StyleValues.Paragraph, StyleId = "ac" };
            StyleName styleName32 = new StyleName() { Val = "Quote" };
            BasedOn basedOn28 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle13 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle23 = new LinkedStyle() { Val = "ad" };
            UIPriority uIPriority31 = new UIPriority() { Val = 29 };
            PrimaryStyle primaryStyle18 = new PrimaryStyle();
            Rsid rsid33 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            Italic italic8 = new Italic();

            styleRunProperties28.Append(italic8);

            style32.Append(styleName32);
            style32.Append(basedOn28);
            style32.Append(nextParagraphStyle13);
            style32.Append(linkedStyle23);
            style32.Append(uIPriority31);
            style32.Append(primaryStyle18);
            style32.Append(rsid33);
            style32.Append(styleRunProperties28);

            Style style33 = new Style() { Type = StyleValues.Character, StyleId = "ad", CustomStyle = true };
            StyleName styleName33 = new StyleName() { Val = "引用文 (文字)" };
            BasedOn basedOn29 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle24 = new LinkedStyle() { Val = "ac" };
            UIPriority uIPriority32 = new UIPriority() { Val = 29 };
            Rsid rsid34 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties29 = new StyleRunProperties();
            Italic italic9 = new Italic();
            FontSize fontSize21 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties29.Append(italic9);
            styleRunProperties29.Append(fontSize21);
            styleRunProperties29.Append(fontSizeComplexScript22);

            style33.Append(styleName33);
            style33.Append(basedOn29);
            style33.Append(linkedStyle24);
            style33.Append(uIPriority32);
            style33.Append(rsid34);
            style33.Append(styleRunProperties29);

            Style style34 = new Style() { Type = StyleValues.Paragraph, StyleId = "21" };
            StyleName styleName34 = new StyleName() { Val = "Intense Quote" };
            BasedOn basedOn30 = new BasedOn() { Val = "a" };
            NextParagraphStyle nextParagraphStyle14 = new NextParagraphStyle() { Val = "a" };
            LinkedStyle linkedStyle25 = new LinkedStyle() { Val = "22" };
            UIPriority uIPriority33 = new UIPriority() { Val = 30 };
            PrimaryStyle primaryStyle19 = new PrimaryStyle();
            Rsid rsid35 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties14 = new StyleParagraphProperties();
            Indentation indentation2 = new Indentation() { Left = "720", Right = "720" };

            styleParagraphProperties14.Append(indentation2);

            StyleRunProperties styleRunProperties30 = new StyleRunProperties();
            Bold bold18 = new Bold();
            Italic italic10 = new Italic();
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties30.Append(bold18);
            styleRunProperties30.Append(italic10);
            styleRunProperties30.Append(fontSizeComplexScript23);

            style34.Append(styleName34);
            style34.Append(basedOn30);
            style34.Append(nextParagraphStyle14);
            style34.Append(linkedStyle25);
            style34.Append(uIPriority33);
            style34.Append(primaryStyle19);
            style34.Append(rsid35);
            style34.Append(styleParagraphProperties14);
            style34.Append(styleRunProperties30);

            Style style35 = new Style() { Type = StyleValues.Character, StyleId = "22", CustomStyle = true };
            StyleName styleName35 = new StyleName() { Val = "引用文 2 (文字)" };
            BasedOn basedOn31 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle26 = new LinkedStyle() { Val = "21" };
            UIPriority uIPriority34 = new UIPriority() { Val = 30 };
            Rsid rsid36 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties31 = new StyleRunProperties();
            Bold bold19 = new Bold();
            Italic italic11 = new Italic();
            FontSize fontSize22 = new FontSize() { Val = "24" };

            styleRunProperties31.Append(bold19);
            styleRunProperties31.Append(italic11);
            styleRunProperties31.Append(fontSize22);

            style35.Append(styleName35);
            style35.Append(basedOn31);
            style35.Append(linkedStyle26);
            style35.Append(uIPriority34);
            style35.Append(rsid36);
            style35.Append(styleRunProperties31);

            Style style36 = new Style() { Type = StyleValues.Character, StyleId = "ae" };
            StyleName styleName36 = new StyleName() { Val = "Subtle Emphasis" };
            UIPriority uIPriority35 = new UIPriority() { Val = 19 };
            PrimaryStyle primaryStyle20 = new PrimaryStyle();
            Rsid rsid37 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties32 = new StyleRunProperties();
            Italic italic12 = new Italic();
            Color color2 = new Color() { Val = "5A5A5A", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A5" };

            styleRunProperties32.Append(italic12);
            styleRunProperties32.Append(color2);

            style36.Append(styleName36);
            style36.Append(uIPriority35);
            style36.Append(primaryStyle20);
            style36.Append(rsid37);
            style36.Append(styleRunProperties32);

            Style style37 = new Style() { Type = StyleValues.Character, StyleId = "23" };
            StyleName styleName37 = new StyleName() { Val = "Intense Emphasis" };
            BasedOn basedOn32 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority36 = new UIPriority() { Val = 21 };
            PrimaryStyle primaryStyle21 = new PrimaryStyle();
            Rsid rsid38 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties33 = new StyleRunProperties();
            Bold bold20 = new Bold();
            Italic italic13 = new Italic();
            FontSize fontSize23 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties33.Append(bold20);
            styleRunProperties33.Append(italic13);
            styleRunProperties33.Append(fontSize23);
            styleRunProperties33.Append(fontSizeComplexScript24);
            styleRunProperties33.Append(underline1);

            style37.Append(styleName37);
            style37.Append(basedOn32);
            style37.Append(uIPriority36);
            style37.Append(primaryStyle21);
            style37.Append(rsid38);
            style37.Append(styleRunProperties33);

            Style style38 = new Style() { Type = StyleValues.Character, StyleId = "af" };
            StyleName styleName38 = new StyleName() { Val = "Subtle Reference" };
            BasedOn basedOn33 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority37 = new UIPriority() { Val = 31 };
            PrimaryStyle primaryStyle22 = new PrimaryStyle();
            Rsid rsid39 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties34 = new StyleRunProperties();
            FontSize fontSize24 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };
            Underline underline2 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties34.Append(fontSize24);
            styleRunProperties34.Append(fontSizeComplexScript25);
            styleRunProperties34.Append(underline2);

            style38.Append(styleName38);
            style38.Append(basedOn33);
            style38.Append(uIPriority37);
            style38.Append(primaryStyle22);
            style38.Append(rsid39);
            style38.Append(styleRunProperties34);

            Style style39 = new Style() { Type = StyleValues.Character, StyleId = "24" };
            StyleName styleName39 = new StyleName() { Val = "Intense Reference" };
            BasedOn basedOn34 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority38 = new UIPriority() { Val = 32 };
            PrimaryStyle primaryStyle23 = new PrimaryStyle();
            Rsid rsid40 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties35 = new StyleRunProperties();
            Bold bold21 = new Bold();
            FontSize fontSize25 = new FontSize() { Val = "24" };
            Underline underline3 = new Underline() { Val = UnderlineValues.Single };

            styleRunProperties35.Append(bold21);
            styleRunProperties35.Append(fontSize25);
            styleRunProperties35.Append(underline3);

            style39.Append(styleName39);
            style39.Append(basedOn34);
            style39.Append(uIPriority38);
            style39.Append(primaryStyle23);
            style39.Append(rsid40);
            style39.Append(styleRunProperties35);

            Style style40 = new Style() { Type = StyleValues.Character, StyleId = "af0" };
            StyleName styleName40 = new StyleName() { Val = "Book Title" };
            BasedOn basedOn35 = new BasedOn() { Val = "a0" };
            UIPriority uIPriority39 = new UIPriority() { Val = 33 };
            PrimaryStyle primaryStyle24 = new PrimaryStyle();
            Rsid rsid41 = new Rsid() { Val = "002206EE" };

            StyleRunProperties styleRunProperties36 = new StyleRunProperties();
            RunFonts runFonts29 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia };
            Bold bold22 = new Bold();
            Italic italic14 = new Italic();
            FontSize fontSize26 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties36.Append(runFonts29);
            styleRunProperties36.Append(bold22);
            styleRunProperties36.Append(italic14);
            styleRunProperties36.Append(fontSize26);
            styleRunProperties36.Append(fontSizeComplexScript26);

            style40.Append(styleName40);
            style40.Append(basedOn35);
            style40.Append(uIPriority39);
            style40.Append(primaryStyle24);
            style40.Append(rsid41);
            style40.Append(styleRunProperties36);

            Style style41 = new Style() { Type = StyleValues.Paragraph, StyleId = "af1" };
            StyleName styleName41 = new StyleName() { Val = "TOC Heading" };
            BasedOn basedOn36 = new BasedOn() { Val = "1" };
            NextParagraphStyle nextParagraphStyle15 = new NextParagraphStyle() { Val = "a" };
            UIPriority uIPriority40 = new UIPriority() { Val = 39 };
            SemiHidden semiHidden21 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed13 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle25 = new PrimaryStyle();
            Rsid rsid42 = new Rsid() { Val = "002206EE" };

            StyleParagraphProperties styleParagraphProperties15 = new StyleParagraphProperties();
            OutlineLevel outlineLevel12 = new OutlineLevel() { Val = 9 };

            styleParagraphProperties15.Append(outlineLevel12);

            style41.Append(styleName41);
            style41.Append(basedOn36);
            style41.Append(nextParagraphStyle15);
            style41.Append(uIPriority40);
            style41.Append(semiHidden21);
            style41.Append(unhideWhenUsed13);
            style41.Append(primaryStyle25);
            style41.Append(rsid42);
            style41.Append(styleParagraphProperties15);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);
            styles1.Append(style34);
            styles1.Append(style35);
            styles1.Append(style36);
            styles1.Append(style37);
            styles1.Append(style38);
            styles1.Append(style39);
            styles1.Append(style40);
            styles1.Append(style41);

            return styles1;
        }



        static public void OfficeOpenXMLDocument(string text,DocumentFormat.OpenXml.Packaging.MainDocumentPart mainDoc)
        {
            StructualText.StructualTextReader sr = new StructualText.StructualTextReader(text);
            Document d = new Document();
            d.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            d.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            d.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            d.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            d.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            bool isTop = true;
            int count = 2;
            
            Body b1 = new Body();
            while (sr.ReadLine())
            {
                if (sr.Type == StructualText.StructualTextReader.LineType.Comment)
                {
                }
                else
                {
                    Paragraph p = new Paragraph();
                    if (sr.Type == StructualText.StructualTextReader.LineType.Text)
                    {
                        p.Append(new Run(new Text(sr.Text)));
                    }
                    else if (sr.Type == StructualText.StructualTextReader.LineType.Image)
                    {
                        using(System.IO.Stream data =new System.IO.FileStream(sr.Text,System.IO.FileMode.Open,System.IO.FileAccess.Read)){
                            ImagePart imagePart1 = mainDoc.AddNewPart<ImagePart>(Converter.GetImageMimetype(sr.Text), "rId" + count);
                            imagePart1.FeedData(data);
                            data.Close();

                            long width, height;
                            using (System.Drawing.Bitmap bmpSrc = new System.Drawing.Bitmap(sr.Text))
                            {
                                width =5000000L;
                                height = 5000000L * bmpSrc.Height / bmpSrc.Width;
                            }

                            Run run1 = new Run();

                            Drawing drawing1 = new Drawing();

                            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
                            Wp.Extent extent1 = new Wp.Extent() { Cx = width, Cy = height };
                            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "図 "+count, Description = "" };

                            A.Graphic graphic1 = new A.Graphic();
                            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

                            Pic.Picture picture1 = new Pic.Picture();
                            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

                            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
                            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture " + count, Description ="" };

                            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

                            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
                            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

                            Pic.BlipFill blipFill1 = new Pic.BlipFill();
                            A.Blip blip1 = new A.Blip() { Embed = "rId"+count, CompressionState = A.BlipCompressionValues.Print };
                            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

                            A.Stretch stretch1 = new A.Stretch();

                            blipFill1.Append(blip1);
                            blipFill1.Append(sourceRectangle1);
                            blipFill1.Append(stretch1);

                            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                            A.Transform2D transform2D1 = new A.Transform2D();
                            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                            A.Extents extents1 = new A.Extents() { Cx = width, Cy = height };

                            transform2D1.Append(offset1);
                            transform2D1.Append(extents1);

                            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                            presetGeometry1.Append(adjustValueList1);

                            shapeProperties1.Append(transform2D1);
                            shapeProperties1.Append(presetGeometry1);

                            picture1.Append(nonVisualPictureProperties1);
                            picture1.Append(blipFill1);
                            picture1.Append(shapeProperties1);

                            graphicData1.Append(picture1);

                            graphic1.Append(graphicData1);
                            
                            inline1.Append(extent1);
                            inline1.Append(docProperties1);
                            inline1.Append(graphic1);

                            drawing1.Append(inline1);

                            run1.Append(drawing1);

                            p.Append(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
                            p.Append(run1);

                            count++;
                        }
                    }
                    else if (sr.Type == StructualText.StructualTextReader.LineType.Math)
                    {
                        M.OfficeMath om1 = new M.OfficeMath();
                        om1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");

                        M.Run r1 = new M.Run();

                        RunProperties rp1 = new RunProperties();
                        RunFonts rf1 = new RunFonts() { Ascii = "Cambria Math" };

                        rp1.Append(rf1);
                        M.Text tx1 = new M.Text();
                        tx1.Text = sr.Text;

                        r1.Append(rp1);
                        r1.Append(tx1);

                        om1.Append(r1);

                        p.Append(om1);
                    }
                    else if (sr.Type == StructualText.StructualTextReader.LineType.Structure)
                    {
                        if (sr.Depth == 1 && !isTop)
                        {
                            b1.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));

                        }
                        if (isTop) { isTop = false; }
                        ParagraphProperties wp = new ParagraphProperties();
                        ParagraphStyleId psid = new ParagraphStyleId();
                        string[] stID = { "a3", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
                        psid.SetAttribute(new OpenXmlAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", stID[Math.Min(9, sr.Depth - 1)]));
                        wp.Append(psid);
                        p.Append(wp);
                        Run r = new Run(new Text(sr.Text));
                        p.Append(r);
                        ////ToDo
                    }
                    b1.Append(p);
                }
            }
            d.Append(b1);
            mainDoc.Document = d;
        }

        static public Document sample1()
        {
            Document d = new Document();
            d.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                Body b1 = new Body();

                    Paragraph p1 = new Paragraph();

                        M.OfficeMath om1 = new M.OfficeMath();
                        om1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");

                            M.Run r1 = new M.Run();

                            RunProperties rp1 = new RunProperties();
                            RunFonts rf1 = new RunFonts() { Ascii = "Cambria Math" };

                            rp1.Append(rf1);
                            M.Text tx1 = new M.Text();
                            tx1.Text = "a*b=c";

                            r1.Append(rp1);
                            r1.Append(tx1);

                        om1.Append(r1);

                    p1.Append(om1);

                b1.Append(p1);

            d.Append(b1);
            return d;
        }
    }
}
