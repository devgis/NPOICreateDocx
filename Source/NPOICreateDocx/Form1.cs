using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Dml;
/*
 * 本例子提供的NPOI是tonyqus提供的2.0.9.0源码经过修改。
 * 例中包括：
 * 1、docx创建和存储
 * 2、页面设置
 * 3、段落创建、设置
 * 4、字体设置
 * 5、表格创建、定位
 * 6、插图操作，包括表格插图
 * vs2010
 * netframework4
 * 创建的docx在word2007可以打开
 * 2014-6-2
 * 问题：不能插入图表，如柱状图，饼图等
 */
namespace NPOICreateDocx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //图片位置
            String m_PicPath = "..\\..\\..\\pic\\";
            FileStream gfs = null;
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;
            //页面设置A4纵向
            m_SectPr.pgSz.h = (ulong)16838;
            m_SectPr.pgSz.w = (ulong)11906;
            XWPFParagraph gp = m_Docx.CreateParagraph();
            gp.GetCTPPr().AddNewJc().val = ST_Jc.center; //水平居中
            XWPFRun gr = gp.CreateRun();
            gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().hint = ST_Hint.eastAsia;
            gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)44;//2号字体
            gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)44;
            gr.GetCTR().AddNewRPr().AddNewB().val = true; //加粗
            gr.GetCTR().AddNewRPr().AddNewColor().val = "red";//字体颜色
            gr.SetText("NPOI创建Word2007Docx");
            gp = m_Docx.CreateParagraph();
            gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gr = gp.CreateRun();
            CT_RPr rpr = gr.GetCTR().AddNewRPr();
            CT_Fonts rfonts = rpr.AddNewRFonts();
            rfonts.ascii = "宋体";
            rfonts.eastAsia = "宋体";
            rpr.AddNewSz().val = (ulong)21;//5号字体
            rpr.AddNewSzCs().val = (ulong)21;
            gr.SetText("NPOI，顾名思义，就是POI的.NET版本。那POI又是什么呢？POI是一套用Java写成的库，能够帮助开 发者在没有安装微软Office的情况下读写Office 97-2003的文件，支持的文件格式包括xls, doc, ppt等 。目前POI的稳定版本中支持Excel文件格式(xls和xlsx)，其他的都属于不稳定版本（放在poi的scrachpad目录 中）。");
            //创建表
            XWPFTable table = m_Docx.CreateTable(1, 4);//创建一行4列表
            CT_Tbl m_CTTbl = m_Docx.Document.body.GetTblArray()[0];//获得文档第一张表
            CT_TblPr m_CTTblPr = m_CTTbl.AddNewTblPr();
            m_CTTblPr.AddNewTblW().w = "2000"; //表宽
            m_CTTblPr.AddNewTblW().type = ST_TblWidth.dxa;
            m_CTTblPr.tblpPr = new CT_TblPPr();//表定位
            m_CTTblPr.tblpPr.tblpX = "4003";//表左上角坐标
            m_CTTblPr.tblpPr.tblpY = "365";
            m_CTTblPr.tblpPr.tblpXSpec = ST_XAlign.Null;//若不为“Null”，则优先tblpX，即表由tblpXSpec定位
            m_CTTblPr.tblpPr.tblpYSpec = ST_YAlign.Null;//若不为“Null”，则优先tblpY，即表由tblpYSpec定位  
            m_CTTblPr.tblpPr.leftFromText = (ulong)180;
            m_CTTblPr.tblpPr.rightFromText = (ulong)180;
            m_CTTblPr.tblpPr.vertAnchor = ST_VAnchor.text;
            m_CTTblPr.tblpPr.horzAnchor = ST_HAnchor.page; 

            //表1行4列充值：a,b,c,d
            table.GetRow(0).GetCell(0).SetText("a");
            table.GetRow(0).GetCell(1).SetText("b");
            table.GetRow(0).GetCell(2).SetText("c");
            table.GetRow(0).GetCell(3).SetText("d");
            CT_Row m_NewRow = new CT_Row();//创建1行
            XWPFTableRow m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row); //必须要！！！
            XWPFTableCell cell = m_Row.CreateCell();//创建单元格，也创建了一个CT_P
            CT_Tc cttc = cell.GetCTTc();
            CT_TcPr ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";//合并3列
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "666";
            cell = m_Row.CreateCell();//创建单元格，也创建了一个CT_P
            cell.GetCTTc().GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cell.GetCTTc().GetPList()[0].AddNewR().AddNewT().Value = "e";
            //合并3列，合并2行
            //1行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            cell.SetText("f");
            cell = m_Row.CreateCell();//从第2单元格开始合并
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";//合并3列
            ctPr.AddNewVMerge().val = ST_Merge.restart;//开始合并行
            ctPr.AddNewVAlign().val = ST_VerticalJc.center;//垂直居中
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "777";
            //2行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            cell.SetText("g");
            cell = m_Row.CreateCell();//第2单元格
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            ctPr.gridSpan.val = "3";//合并3列
            ctPr.AddNewVMerge().val = ST_Merge.@continue;//继续合并行

            //表插入图片
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            //inline方式插入图片
            //gp = table.GetRow(table.Rows.Count - 1).GetCell(0).GetParagraph(table.GetRow(table.Rows.Count - 1).GetCell(0).GetCTTc().GetPList()[0]);//获得指定表单元格的段
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);  
            gr = gp.CreateRun();//创建run
            gfs = new FileStream( m_PicPath + "1.jpg", FileMode.Open, FileAccess.Read);//读取图片文件
            gr.AddPicture(gfs, (int)PictureType.PNG, "1.jpg", 500000, 500000);//插入图片
            gfs.Close();
            //Anchor方式插入图片
            CT_Anchor an = new CT_Anchor();
            an.distB = (uint)(0);
            an.distL = 114300u;
            an.distR = 114300U;
            an.distT = 0U;
            an.relativeHeight = 251658240u;
            an.behindDoc = false; //"0"
            an.locked = false;  //"0"
            an.layoutInCell = true;  //"1"
            an.allowOverlap = true;  //"1" 

            NPOI.OpenXmlFormats.Dml.CT_Point2D simplePos = new NPOI.OpenXmlFormats.Dml.CT_Point2D();
            simplePos.x = (long)0;
            simplePos.y = (long)0;
            CT_EffectExtent effectExtent = new CT_EffectExtent();
            effectExtent.b = 0L;
            effectExtent.l = 0L;
            effectExtent.r = 0L;
            effectExtent.t = 0L;
            //wrapSquare(四周)
            cell = m_Row.CreateCell();//第2单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            CT_WrapSquare wrapSquare = new CT_WrapSquare();
            wrapSquare.wrapText = ST_WrapText.bothSides;
            gfs = new FileStream(m_PicPath + "1.png", FileMode.Open, FileAccess.Read);//读取图片文件
            gr.AddPicture(gfs, (int)PictureType.PNG, "1.png", 500000, 500000, 0, 0, wrapSquare, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gfs.Close();
            //wrapTight（紧密）
            cell = m_Row.CreateCell();//第3单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            CT_WrapTight wrapTight = new CT_WrapTight();
            wrapTight.wrapText = ST_WrapText.bothSides;
            wrapTight.wrapPolygon = new CT_WrapPath();
            wrapTight.wrapPolygon.edited = false;
            wrapTight.wrapPolygon.start = new CT_Point2D();
            wrapTight.wrapPolygon.start.x = 0;
            wrapTight.wrapPolygon.start.y = 0;
            CT_Point2D lineTo = new CT_Point2D();
            wrapTight.wrapPolygon.lineTo = new List<CT_Point2D>();
            lineTo = new CT_Point2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            gfs = new FileStream(m_PicPath + "11.png", FileMode.Open, FileAccess.Read);//读取图片文件
            gr.AddPicture(gfs, (int)PictureType.PNG, "1.png", 500000, 500000, 0, 0, wrapTight, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gfs.Close();
            //wrapThrough(穿越)
            cell = m_Row.CreateCell();//第4单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            gfs = new FileStream(m_PicPath + "15.png", FileMode.Open, FileAccess.Read);//读取图片文件
            CT_WrapThrough wrapThrough = new CT_WrapThrough();
            wrapThrough.wrapText = ST_WrapText.bothSides;
            wrapThrough.wrapPolygon = new CT_WrapPath();
            wrapThrough.wrapPolygon.edited = false;
            wrapThrough.wrapPolygon.start = new CT_Point2D();
            wrapThrough.wrapPolygon.start.x = 0;
            wrapThrough.wrapPolygon.start.y = 0;
            lineTo = new CT_Point2D();
            wrapThrough.wrapPolygon.lineTo = new List<CT_Point2D>();
            lineTo = new CT_Point2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            gr.AddPicture(gfs, (int)PictureType.PNG, "15.png", 500000, 500000, 0, 0, wrapThrough, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gfs.Close();

            gp = m_Docx.CreateParagraph();
            gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gr = gp.CreateRun();
            gr.SetText("NPOI是POI项目的.NET版本。POI是一个开源的Java读写Excel、WORD等微软OLE2组件文档的项目。使用NPOI你就可以在没有安装Office或者相应环境的机器上对WORD/EXCEL文档进行读写。NPOI是构建在POI3.x版本之上的，它可以在没有安装Office的情况下对Word/Excel文档进行读写操作。");
            gp = m_Docx.CreateParagraph();
            gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gr = gp.CreateRun();
            gr.SetText("NPOI之所以强大，并不是因为它支持导出Excel，而是因为它支持导入Excel，并能“理解”OLE2文档结构，这也是其他一些Excel读写库比较弱的方面。通常，读入并理解结构远比导出来得复杂，因为导入你必须假设一切情况都是可能的，而生成你只要保证满足你自己需求就可以了，如果把导入需求和生成需求比做两个集合，那么生成需求通常都是导入需求的子集。");
            //在本段中插图-wrapSquare
            //gr = gp.CreateRun();//创建run
            wrapSquare = new CT_WrapSquare();
            wrapSquare.wrapText = ST_WrapText.bothSides;
            gfs = new FileStream(m_PicPath + "15.png", FileMode.Open, FileAccess.Read);//读取图片文件
            gr.AddPicture(gfs, (int)PictureType.PNG, "15.png", 500000, 500000, 900000, 200000, wrapSquare, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gfs.Close();
          
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms,Path.GetPathRoot(Directory.GetCurrentDirectory()) +  "\\NPOI.docx");
            
        }
        static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
        }
        protected int Indentation(String fontname, int fontsize, int Indentationfonts, FontStyle fs)
        {
            //字显示宽度，用于段首行缩进
            /*字号与fontsize关系
             * 初号（0号）=84，小初=72，1号=52，2号=44，小2=36，3号=32，小3=30，4号=28，小4=24，5号=21，小5=18，6号=15，小6=13，7号=11，8号=10
             */
            Graphics m_tmpGr = this.CreateGraphics();
            m_tmpGr.PageUnit = GraphicsUnit.Point;
            SizeF size = m_tmpGr.MeasureString("好", new Font(fontname, fontsize * 0.75F, fs));
            return (int)size.Width * Indentationfonts * 10;
        }
    }
}
