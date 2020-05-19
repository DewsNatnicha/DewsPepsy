using System;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.IO;
using System.Diagnostics;

namespace DGVPrinterHelper 
{
    #region ซับพอร์ตคลาส แต่ซับเหงื่อดิว

    /// <summary>
    /// ตั้งค่าและใช้บันทึกสำหรับการบันทึกภายใน
    /// </summary>
    class LogManager
    {
        #region จัดการบันทึก!! บันทึกไฟล์ นู่นนี่นั้น
        /// <summary>
        /// path ไปยัง log file (บันทึกไฟล์)
        /// </summary>
        private String basepath;
        public String BasePath
        {
            get { return basepath; }
            set { basepath = value; }
        }

        /// <summary>
        /// ส่วนหัวชื่อ log file
        /// </summary>
        private String logheader;
        public String LogNameHeader
        {
            get { return logheader; }
            set { logheader = value; }
        }

        private int useFrame = 1;

        /// <summary>
        /// กำหนดหมวดหมู่ข้อความการบันทึก
        /// </summary>
        public enum Categories
        {
            Info = 1,
            Warning,
            Error,
            Exception
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override path และชื่อของ log file
        /// </summary>
        /// <param name="userbasepath"></param>
        /// <param name="userlogname"></param>
        public LogManager(String userbasepath, String userlogname)
        {
            BasePath = String.IsNullOrEmpty(userbasepath) ? "." : userbasepath;
            LogNameHeader = String.IsNullOrEmpty(userlogname) ? "MsgLog" : userlogname;

            Log(Categories.Info, "********************* New Trace *********************");
        }

        /// <summary>
        /// บันทึกข้อความโดยใช้หมวดหมู่ที่ให้ไว้
        /// </summary>
        /// <param name="category"></param>
        /// <param name="msg"></param>
        public void Log(Categories category, String msg)
        {
            // เรียก stack
            StackTrace stackTrace = new StackTrace();

            // รับชื่อ method
            String caller = stackTrace.GetFrame(useFrame).GetMethod().Name;

            // เข้าสู่ระบบ
            LogWriter.Write(caller, category, msg, BasePath, LogNameHeader);

            // รีเซ็ต! ตัวแปร frame 
            useFrame = 1;
        }

        /// <summary>
        /// บันทึกข้อความที่ให้ข้อมูล
        /// </summary>
        /// <param name="msg"></param>
        public void LogInfoMsg(String msg)
        {
            useFrame++; //ชน!! ตัวแปร frame สแต็คขึ้น เพื่อข้ามรายการนี้
            Log(Categories.Info, msg);
        }

        /// <summary>
        /// Log an error message
        /// บันทึกข้อความแสดงข้อผิดพลาด
        /// </summary>
        /// <param name="msg"></param>
        public void LogErrorMsg(String msg)
        {
            useFrame++; //ชน!! ตัวแปร frame สแต็คขึ้น เพื่อข้ามรายการนี้
            Log(Categories.Error, msg);
        }

        /// <summary>
        /// บันทึก exception (ข้อยกเว้น)
        /// </summary>
        /// <param name="ex"></param>
        public void Log(Exception ex)
        {
            useFrame++; //ชน!! ตัวแปร frame สแต็คขึ้น เพื่อข้ามรายการนี้
            Log(Categories.Exception, String.Format("{0} from {1}", ex.Message, ex.Source));
        }
        #endregion
    }

    /// <summary>
    /// ทำการเขียนบันทึกการเขียน โดยใช้ข้อมูลการตั้งค่าในคลาส Log Manager
    /// </summary>
    class LogWriter
    {
        #region บันทึกการเขียน เขียนอะไรไปบ้าง? 
        /// <summary>
        /// สร้างชื่อไฟล์บันทึก ให้เป็นมาตรฐาน ด้วยรูปแบบชื่อของเรา
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static String LogFileName(String name)
        {
            return String.Format("{0}_{1:yyyyMMdd}.Log", name, DateTime.Now);
        }

        /// <summary>
        /// เขียนรายการบันทึกลงในไฟล์ --> ว่าไฟล์ที่ บันทึกจะถูกล้างและปิดเสมอ แต่ว่าข้อความจะไม่หาย
        /// </summary>
        /// <param name="from"></param>
        /// <param name="category"></param>
        /// <param name="msg"></param>
        /// <param name="path"></param>
        /// <param name="name"></param>
        public static void Write(String from, LogManager.Categories category, String msg, String path, String name)
        {
            StringBuilder line = new StringBuilder();
            line.Append(DateTime.Now.ToShortDateString().ToString());
            line.Append("-");
            line.Append(DateTime.Now.ToLongTimeString().ToString());
            line.Append(", ");
            line.Append(category.ToString().PadRight(6, ' '));
            line.Append(",");
            line.Append(from.PadRight(13, ' '));
            line.Append(",");
            line.Append(msg);
            StreamWriter w = new StreamWriter(path + "\\" + LogFileName(name), true);
            w.WriteLine(line.ToString());
            w.Flush();
            w.Close();
        }
        #endregion
    }

    /// <summary>
    /// คลาสสำหรับ ownerdraw event ระบุผู้เรียก + ข้อมูลเซลล์ปัจจุบัน + กราฟิก + ตำแหน่งที่จะวาดเซลล์
    /// </summary>
    public class DGVCellDrawingEventArgs : EventArgs
    {
        #region การวาดเซลล์
        public Graphics g;
        public RectangleF DrawingBounds; //ขอบเขตการวาด
        public DataGridViewCellStyle CellStyle;
        public int row;
        public int column;
        public Boolean Handled; //แตะยัง?

        public DGVCellDrawingEventArgs(Graphics g, RectangleF bounds, DataGridViewCellStyle style,
            int row, int column)
            : base()
        {
            this.g = g;
            DrawingBounds = bounds;
            CellStyle = style;
            this.row = row;
            this.column = column;
            Handled = false; //ยังไม่แตะ!
        }
        #endregion
    }

    /// <summary>
    /// ตัวแทน (delegate) ของ ownerdraw cells - อนุญาตให้ผู้ใช้เรียกให้วาดรูปเซลล์
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    public delegate void CellOwnerDrawEventHandler(object sender, DGVCellDrawingEventArgs e);

    /// <summary>
    /// กด method ส่วนขยาย
    /// </summary>
    public static class Extensions
    {
        #region ส่วนขยาย
        /// <summary>
        /// วิธีการขยายเพื่อพิมพ์ "รูปภาพที่ฝัง" (ImbeddedImages) ทั้งหมดในรายการที่ให้ไว้
        /// </summary>
        /// <typeparam name="?"></typeparam>
        /// <param name="list"></param>
        /// <param name="g"></param>
        /// <param name="pagewidth"></param>
        /// <param name="pageheight"></param>
        /// <param name="margins"></param>
        public static void DrawImbeddedImage<T>(this IEnumerable<T> list,
            Graphics g, int pagewidth, int pageheight, Margins margins)
        {
            foreach (T t in list) //ตัวเลข
            {
                if (t is DGVPrinter.ImbeddedImage)
                {
                    DGVPrinter.ImbeddedImage ii = (DGVPrinter.ImbeddedImage)Convert.ChangeType(t, typeof(DGVPrinter.ImbeddedImage));
                    //แก้ไข - DrawImageUnscaled ****จริงๆแล้ว เป็นการปรับขนาดรูปภาพ!!
                    //g.DrawImageUnscaled(ii.theImage, ii.upperleft(pagewidth, pageheight, margins));
                    g.DrawImage(ii.theImage,
                        new Rectangle(ii.upperleft(pagewidth, pageheight, margins),
                        new Size(ii.theImage.Width, ii.theImage.Height)));
                }
            }
        }
        #endregion

    }
    #endregion

    /// <summary>
    /// ตัวพิมพ์ Data Grid View พิมพ์ฟังก์ชั่นสำหรับ DataGridview ตั้งแต่ MS
    /// </summary>
    public class DGVPrinter
    {
        public enum Alignment { NotSet, Left, Right, Center }
        public enum Location { Header, Footer, Absolute }
        public enum SizeType { CellSize, StringSize, Porportional }
        public enum PrintLocation { All, FirstOnly, LastOnly, None }

        // ภายในคลาส / โครงสร้าง
        #region Internal Classes

        // ระบุเค้าแมว แฮร่! เค้าโครงเดิมกับหน้าใหม่ตามแถวเดิม
        enum paging { keepgoing, outofroom, datachange };

        // อนุญาตให้ผู้ใช้ระบุรูปภาพที่จะพิมพ์เป็นโลโก้ ใน ส่วนหัว/ส่วนท้าย/ลายน้ำ ตามที่พิมพ์ไว้ด้านหลังข้อความ
        public class ImbeddedImage
        {
            public Image theImage { get; set; }
            public Alignment ImageAlignment { get; set; }
            public Location ImageLocation { get; set; }
            public Int32 ImageX { get; set; }
            public Int32 ImageY { get; set; }

            internal Point upperleft(int pagewidth, int pageheight, Margins margins)
            {
                int y = 0;
                int x = 0;

                // รับตำแหน่งที่แน่นอน
                if (ImageLocation == Location.Absolute)
                    return new Point(ImageX, ImageY);

                // ตั้งค่าตำแหน่ง y ตามหัวกระดาษ/ท้ายกระดาษ
                switch (ImageLocation)
                {
                    case Location.Header:
                        y = margins.Top;
                        break;
                    case Location.Footer:
                        y = pageheight - theImage.Height - margins.Bottom;
                        break;
                    default:
                        throw new ArgumentException(String.Format("Unkown value: {0}", ImageLocation));
                }

                // กำหนดตำแหน่ง x ตามซ้าย, ขวา, กึ่งกลาง
                switch (ImageAlignment)
                {
                    case Alignment.Left:
                        x = margins.Left;
                        break;
                    case Alignment.Center:
                        x = (int)(pagewidth / 2 - theImage.Width / 2) + margins.Left;
                        break;
                    case Alignment.Right:
                        x = (int)(pagewidth - theImage.Width) + margins.Left;
                        break;
                    case Alignment.NotSet:
                        x = ImageX;
                        break;
                    default:
                        throw new ArgumentException(String.Format("Unkown value: {0}", ImageAlignment));
                }

                return new Point(x, y);
            }
        }

        public IList<ImbeddedImage> ImbeddedImageList = new List<ImbeddedImage>();

        // จัดการการพิมพ์ ความกว้างของคอลัมน์ --> คอลัมน์ที่ขยาย
        // คอลัมน์แตกออกเป็น "Page Sets" ที่จะถูกพิมพ์ทีละรายการจนกว่า คอลัมน์ทั้งหมดจะถูกพิมพ์
        class PageDef
        {
            public PageDef(Margins m, int count, int pagewidth)
            {
                columnindex = new List<int>(count);
                colstoprint = new List<object>(count);
                colwidths = new List<float>(count);
                colwidthsoverride = new List<float>(count);
                coltotalwidth = 0;
                margins = (Margins)m.Clone();
                pageWidth = pagewidth;
            }

            public List<int> columnindex;
            public List<object> colstoprint;
            public List<float> colwidths;
            public List<float> colwidthsoverride;
            public float coltotalwidth;
            public Margins margins;
            private int pageWidth;

            public int printWidth
            {
                get { return pageWidth - margins.Left - margins.Right; }
            }
        }
        
        IList<PageDef> pagesets;
        int currentpageset = 0;

        // คลาสที่จะการตั้งค่า PrintDialog ระหว่างกระบวนการพิมพ์ --> พอปริ้นแล้วมันจะขึ้นบอก (กล่องโต้ตอบการพิมพ์)
        public class PrintDialogSettingsClass
        {
            public bool AllowSelection = true;
            public bool AllowSomePages = true;
            public bool AllowCurrentPage = true;
            public bool AllowPrintToFile = false;
            public bool ShowHelp = true;
            public bool ShowNetwork = true;
            public bool UseEXDialog = true;
        }

        // คลาสระบุข้อมูลแถวสำหรับการพิมพ์
        public class rowdata
        {
            public DataGridViewRow row = null;
            public float height = 0;
            public bool pagebreak = false;
            public bool splitrow = false;
        }
        #endregion

        // ตัวแปร global
        #region global variables

        // datagridview ที่เรากำลังพิมพ์
        DataGridView dgv = null;

        // พิมพ์เอกสาร
        PrintDocument printDoc = null;

        // logging เข้าสู่ระบบ
        LogManager Logger = null;

        // พิมพ์รายการสถานะ
        Boolean EmbeddedPrinting = false;
        List<rowdata> rowstoprint;
        IList colstoprint;        //แบ่งออกเป็นชุดหน้า(pagesets)สำหรับพิมพ์
        int lastrowprinted = -1;
        int currentrow = -1;
        int fromPage = 0;
        int toPage = -1;
        const int maxPages = 2147483647; //ค่าสูงสุดของ int

        // (option) ตัวเลือกการจัดรูปแบบหน้า (page format)
        int pageHeight = 0;
        float staticheight = 0;
        float rowstartlocation = 0;
        int pageWidth = 0;
        int printWidth = 0;
        float rowheaderwidth = 0;
        int CurrentPage = 0;
        int totalpages;
        PrintRange printRange;

        // ค่าที่คำนวณได้
        //private float headerHeight = 0;
        private float footerHeight = 0;
        private float pagenumberHeight = 0;
        private float colheaderheight = 0;
        //private List<float> rowheights;
        private List<float> colwidths;
        //private List<List<SizeF>> cellsizes;

        #endregion

        // คุณสมบัติ - กำหนดโดยผู้ใช้
        #region properties

        #region global properties

        /// <summary>
        /// Enable logging of of the print process. Default is to log to a file named
        /// เปิดใช้งานการบันทึกของกระบวนการพิมพ์ เริ่มต้นคือการเข้าสู่ไฟล์ชื่อ
        /// 'DGVPrinter_yyyymmdd.Log' in the current directory. Since logging may have 
        /// 'DGVPrinter_yyyymmdd.Log' ในไดเรกทอรีปัจจุบัน ตั้งแต่การบันทึกอาจมี
        /// an impact on performance, it should be used for troubleshooting purposes only.
        /// ผลกระทบต่อประสิทธิภาพควรใช้เพื่อการแก้ไขปัญหาเท่านั้น
        /// </summary>
        protected Boolean enablelogging;
        public Boolean EnableLogging
        {
            get { return enablelogging; }
            set
            {
                enablelogging = value;
                if (enablelogging)
                {
                    Logger = new LogManager(".", "DGVPrinter");
                }
            }
        }

        /// <summary>
        /// Allow the user to change the logging directory. Setting this enables logging by default.
        /// อนุญาตให้ผู้ใช้เปลี่ยนไดเรกทอรีการบันทึก การตั้งค่านี้เปิดใช้งานการบันทึกตามค่าเริ่มต้น
        /// </summary>
        public String LogDirectory
        {
            get
            {
                if (null != Logger)
                    return Logger.BasePath;
                else
                    return null;
            }
            set
            {
                if (null == Logger)
                    EnableLogging = true;
                Logger.BasePath = value;
            }
        }

        /// <summary>
        /// OwnerDraw Event declaration. Callers can subscribe to this event to override the 
        /// OwnerDraw การประกาศเหตุการณ์ ผู้โทรสามารถสมัครรับข้อมูลกิจกรรมนี้เพื่อแทนที่
        /// cell drawing.
        /// การวาดเซลล์
        /// </summary>
        public event CellOwnerDrawEventHandler OwnerDraw;

        /// <summary>
        /// provide an override for the print preview dialog "owner" field
        /// จัดเตรียมการแทนที่สำหรับช่องโต้ตอบ "เจ้าของ" ตัวอย่างก่อนพิมพ์
        /// Note: Changed style for VS2005 compatibility
        /// หมายเหตุ: เปลี่ยนสไตล์เพื่อความเข้ากันได้ของ VS2005
        /// </summary>
        //public Form Owner เจ้าของแบบฟอร์มสาธารณะ
        //{ get; set; }
        protected Form _Owner = null;
        public Form Owner
        {
            get { return _Owner; }
            set { _Owner = value; }
        }

        /// <summary>
        /// provide an override for the print preview zoom setting
        /// จัดเตรียมการแทนที่สำหรับการตั้งค่าการซูมตัวอย่างก่อนพิมพ์
        /// Note: Changed style for VS2005 compatibility
        /// หมายเหตุ: เปลี่ยนสไตล์เพื่อความเข้ากันได้ของ VS2005
        /// </summary>
        //public Double PrintPreviewZoom สาธารณะ Double PrintPreviewZoom
        //{ get; set; }
        protected Double _PrintPreviewZoom = 1.0;
        public Double PrintPreviewZoom
        {
            get { return _PrintPreviewZoom; }
            set { _PrintPreviewZoom = value; }
        }


        /// <summary>
        /// expose printer settings to allow access to calling program
        /// เปิดเผยการตั้งค่าเครื่องพิมพ์เพื่ออนุญาตการเข้าถึงโปรแกรมการโทร
        /// </summary>
        public PrinterSettings PrintSettings
        {
            get { return printDoc.PrinterSettings; }
        }

        /// <summary>
        /// expose settings for the PrintDialog displayed to the user
        /// เปิดเผยการตั้งค่าสำหรับ PrintDialog ที่แสดงต่อผู้ใช้
        /// </summary>
        private PrintDialogSettingsClass printDialogSettings = new PrintDialogSettingsClass();
        public PrintDialogSettingsClass PrintDialogSettings
        {
            get { return printDialogSettings; }
        }

        /// <summary>
        /// Set Printer Name
        /// ตั้งชื่อเครื่องพิมพ์
        /// </summary>
        private String printerName;
        public String PrinterName
        {
            get { return printerName; }
            set { printerName = value; }
        }

        /// <summary>
        /// Allow access to the underlying print document
        /// อนุญาตให้เข้าถึงเอกสารการพิมพ์พื้นฐาน
        /// </summary>
        public PrintDocument printDocument
        {
            get { return printDoc; }
            set { printDoc = value; }
        }

        /// <summary>
        /// Allow caller to set the upper-left corner icon used
        /// อนุญาตให้ผู้โทรตั้งไอคอนมุมบนซ้ายที่ใช้
        /// in the print preview dialog
        /// ในกล่องโต้ตอบตัวอย่างก่อนพิมพ์
        /// </summary>
        private Icon ppvIcon = null;
        public Icon PreviewDialogIcon
        {
            get { return ppvIcon; }
            set { ppvIcon = value; }
        }

        /// <summary>
        /// Allow caller to set print preview dialog
        /// อนุญาตให้ผู้โทรตั้งค่ากล่องโต้ตอบตัวอย่างก่อนพิมพ์
        /// </summary>
        private PrintPreviewDialog previewdialog = null;
        public PrintPreviewDialog PreviewDialog
        {
            get { return previewdialog; }
            set { previewdialog = value; }
        }

        /// <summary>
        /// Flag to control whether or not we print the Page Header
        /// ตั้งค่าสถานะเพื่อควบคุมว่าเราพิมพ์ Page Header หรือไม่
        /// </summary>
        private Boolean printHeader = true;
        public Boolean PrintHeader
        {
            get { return printHeader; }
            set { printHeader = value; }
        }

        /// <summary>
        /// Determine the height of the header
        /// กำหนดความสูงของส่วนหัว
        /// </summary>
        private float HeaderHeight
        {
            get
            {
                float headerheight = 0;

                // Add in title and subtitle heights - this is sensitive to 
                // เพิ่มในชื่อเรื่องและความสูงของคำบรรยาย - เรื่องนี้ละเอียดอ่อน
                // wether or not titles are printed on the current page
                // พิมพ์ชื่อหรือไม่บนหน้าปัจจุบัน
                // TitleHeight and SubTitleHeight have their respective spacing
                // TitleHeight และ SubTitleHeight มีระยะห่างตามลำดับ
                // already included
                // รวมแล้ว
                headerheight += TitleHeight + SubTitleHeight;

                // Add in column header heights
                // เพิ่มความสูงส่วนหัวคอลัมน์
                if ((bool)PrintColumnHeaders)
                {
                    headerheight += colheaderheight;
                }

                // return calculated height
                // ส่งคืนความสูงที่คำนวณได้
                return headerheight;
            }
        }

        /// <summary>
        /// Flag to control whether or not we print the Page Footer
        /// ตั้งค่าสถานะเพื่อควบคุมว่าเราพิมพ์ Page Footer หรือไม่
        /// </summary>
        private Boolean printFooter = true;
        public Boolean PrintFooter
        {
            get { return printFooter; }
            set { printFooter = value; }
        }

        /// <summary>
        /// Flag to control whether or not we print the Column Header line
        /// ตั้งค่าสถานะเพื่อควบคุมว่าเราพิมพ์บรรทัดคอลัมน์ส่วนหัวหรือไม่
        /// </summary>
        private Boolean? printColumnHeaders;
        public Boolean? PrintColumnHeaders
        {
            get { return printColumnHeaders; }
            set { printColumnHeaders = value; }
        }

        /// <summary>
        /// Flag to control whether or not we print the Column Header line
        /// ตั้งค่าสถานะเพื่อควบคุมว่าเราพิมพ์บรรทัดคอลัมน์ส่วนหัวหรือไม่
        /// Defaults to False to match previous functionality
        /// ค่าเริ่มต้นเป็นเท็จเพื่อให้ตรงกับการทำงานก่อนหน้านี้
        /// </summary>
        private Boolean? printRowHeaders = false;
        public Boolean? PrintRowHeaders
        {
            get { return printRowHeaders; }
            set { printRowHeaders = value; }
        }

        /// <summary>
        /// Flag to control whether rows are printed whole or if partial
        /// ตั้งค่าสถานะเพื่อควบคุมว่าจะพิมพ์แถวทั้งหมดหรือบางส่วน
        /// rows should be printed to fill the bottom of the page. Turn this
        /// ควรพิมพ์แถวเพื่อเติมด้านล่างของหน้า หมุนสิ่งนี้
        /// "Off" (i.e. false) to print cells/rows deeper than one page
        /// "ปิด" (เช่นเท็จ) เพื่อพิมพ์เซลล์ / แถวที่ลึกกว่าหนึ่งหน้า
        /// </summary>
        private Boolean keepRowsTogether = true;
        public Boolean KeepRowsTogether
        {
            get { return keepRowsTogether; }
            set { keepRowsTogether = value; }
        }

        /// <summary>
        /// How much of a row must show on the current page before it is 
        /// จำนวนแถวต้องแสดงในหน้าปัจจุบันก่อนที่จะเป็น
        /// split when KeepRowsTogether is set to true.
        /// แยกเมื่อ KeepRowsTogether ตั้งค่าเป็นจริง
        /// </summary>
        private float keeprowstogethertolerance = 15;
        public float KeepRowsTogetherTolerance
        {
            get { return keeprowstogethertolerance; }
            set { keeprowstogethertolerance = value; }
        }

        #endregion

        // Title หัวข้อ
        #region title properties

        // override flag แทนที่ธง
        bool overridetitleformat = false;

        // formatted height of title ความสูงในรูปแบบของชื่อเรื่อง
        float titleheight = 0;

        /// <summary>
        /// Title for this report. Default is empty.
        /// ชื่อสำหรับรายงานนี้ ค่าเริ่มต้นว่างเปล่า
        /// </summary>
        private String title;
        public String Title
        {
            get { return title; }
            set
            {
                title = value;
                if (docName == null)
                {
                    printDoc.DocumentName = value;
                }
            }
        }

        /// <summary>
        /// Name of the document. Default is report title (can be empty)
        /// ชื่อของเอกสาร ค่าเริ่มต้นคือชื่อรายงาน (ว่างเปล่า)
        /// </summary>
        private String docName;
        public String DocName
        {
            get { return docName; }
            set { printDoc.DocumentName = value; docName = value; }
        }

        /// <summary>
        /// Font for the title. Default is Tahoma, 18pt.
        /// แบบอักษรสำหรับชื่อเรื่อง ค่าเริ่มต้นคือ Tahoma, 18pt
        /// </summary>
        private Font titlefont;
        public Font TitleFont
        {
            get { return titlefont; }
            set { titlefont = value; }
        }

        /// <summary>
        /// Foreground color for the title. Default is Black
        /// สีพื้นหน้าสำหรับชื่อเรื่อง ค่าเริ่มต้นคือสีดำ
        /// </summary>
        private Color titlecolor;
        public Color TitleColor
        {
            get { return titlecolor; }
            set { titlecolor = value; }
        }

        /// <summary>
        /// Allow override of the header cell format object
        /// อนุญาตการแทนที่ของวัตถุรูปแบบเซลล์ส่วนหัว
        /// </summary>
        private StringFormat titleformat;
        public StringFormat TitleFormat
        {
            get { return titleformat; }
            set { titleformat = value; overridetitleformat = true; }
        }

        /// <summary>
        /// Allow the user to override the title string alignment. Default value is 
        /// อนุญาตให้ผู้ใช้แทนที่การจัดตำแหน่งสตริงหัวเรื่อง ค่าเริ่มต้นคือ
        /// Alignment - Near; 
        /// การจัดตำแหน่ง - ใกล้;
        /// </summary>
        public StringAlignment TitleAlignment
        {
            get { return titleformat.Alignment; }
            set
            {
                titleformat.Alignment = value;
                overridetitleformat = true;
            }
        }

        /// <summary>
        /// Allow the user to override the title string format flags. Default values
        /// อนุญาตให้ผู้ใช้แทนที่ค่าสถานะรูปแบบสตริงชื่อเรื่อง ค่าเริ่มต้น
        /// are: FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        public StringFormatFlags TitleFormatFlags
        {
            get { return titleformat.FormatFlags; }
            set
            {
                titleformat.FormatFlags = value;
                overridetitleformat = true;
            }
        }

        /// <summary>
        /// Control where in the document the title prints
        /// ควบคุมตำแหน่งในเอกสารที่ชื่อเรื่องพิมพ์
        /// </summary>
        private PrintLocation titleprint = PrintLocation.All;
        public PrintLocation TitlePrint
        {
            get { return titleprint; }
            set { titleprint = value; }
        }

        /// <summary>
        /// Return the title height based whether to print it or not
        /// ส่งคืนความสูงของหัวเรื่องตามว่าจะพิมพ์หรือไม่
        /// </summary>
        private float TitleHeight
        {
            get
            {
                if (PrintLocation.All == TitlePrint)
                    return titleheight + titlespacing;

                if ((PrintLocation.FirstOnly == TitlePrint) && (1 == CurrentPage))
                    return titleheight + titlespacing;

                if ((PrintLocation.LastOnly == TitlePrint) && (totalpages == CurrentPage))
                    return titleheight + titlespacing;

                return 0;
            }
        }

        /// <summary>
        /// Mandatory spacing between the grid and the footer
        /// ระยะห่างที่บังคับระหว่างตารางและส่วนท้าย
        /// </summary>
        private float titlespacing;
        public float TitleSpacing
        {
            get { return titlespacing; }
            set { titlespacing = value; }
        }

        /// <summary>
        /// สีพื้นหลังของ Title
        /// </summary>
        private Brush titlebackground;
        public Brush TitleBackground
        {
            get { return titlebackground; }
            set { titlebackground = value; }
        }

        /// <summary>
        /// ความหนาของ Title
        /// </summary>
        private Pen titleborder;
        public Pen TitleBorder
        {
            get { return titleborder; }
            set { titleborder = value; }
        }

        #endregion

        // คุณสมบัติของ SubTitle 
        #region subtitle properties

        // override ของ subtitle
        bool overridesubtitleformat = false;

        // ความสูงของ subtitle
        float subtitleheight = 0;

        /// <summary>
        /// SubTitle ค่าเริ่มต้นว่าง!
        /// </summary>
        private String subtitle;
        public String SubTitle
        {
            get { return subtitle; }
            set { subtitle = value; }
        }

        /// <summary>
        /// แบบอักษร subtitle ค่าเริ่มต้นคือ Tahoma, 12pt
        /// </summary>
        private Font subtitlefont;
        public Font SubTitleFont
        {
            get { return subtitlefont; }
            set { subtitlefont = value; }
        }

        /// <summary>
        /// สี Foreground เป็น Black
        /// </summary>
        private Color subtitlecolor;
        public Color SubTitleColor
        {
            get { return subtitlecolor; }
            set { subtitlecolor = value; }
        }

        /// <summary>
        /// อนุญาต override ของรูปแบบเซลล์ **ส่วนหัว
        /// </summary>
        private StringFormat subtitleformat;
        public StringFormat SubTitleFormat
        {
            get { return subtitleformat; }
            set { subtitleformat = value; overridesubtitleformat = true; }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override การจัดตำแหน่ง subtitle ค่าเริ่มต้นคือ Alignment - Near; 
        /// </summary>
        public StringAlignment SubTitleAlignment
        {
            get { return subtitleformat.Alignment; }
            set
            {
                subtitleformat.Alignment = value;
                overridesubtitleformat = true;
            }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override ค่า subtitle ค่าเริ่มต้น : FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        public StringFormatFlags SubTitleFormatFlags
        {
            get { return subtitleformat.FormatFlags; }
            set
            {
                subtitleformat.FormatFlags = value;
                overridesubtitleformat = true;
            }
        }

        /// <summary>
        /// ตำแหน่งในเอกสารที่พิมพ์
        /// </summary>
        private PrintLocation subtitleprint = PrintLocation.All;
        public PrintLocation SubTitlePrint
        {
            get { return subtitleprint; }
            set { subtitleprint = value; }
        }

        /// <summary>
        /// Return ความสูงของ title ว่าจะพิมพ์หรือไม่
        /// </summary>
        private float SubTitleHeight
        {
            get
            {
                if (PrintLocation.All == SubTitlePrint)
                    return subtitleheight + subtitlespacing;

                if ((PrintLocation.FirstOnly == SubTitlePrint) && (1 == CurrentPage))
                    return subtitleheight + subtitlespacing;

                if ((PrintLocation.LastOnly == SubTitlePrint) && (totalpages == CurrentPage))
                    return subtitleheight + subtitlespacing;

                return 0;
            }
        }

        /// <summary>
        /// ระยะห่าง ระหว่างตารางและส่วนท้าย
        /// </summary>
        private float subtitlespacing;
        public float SubTitleSpacing
        {
            get { return subtitlespacing; }
            set { subtitlespacing = value; }
        }

        /// <summary>
        /// สีพื้นหลังของ Title
        /// </summary>
        private Brush subtitlebackground;
        public Brush SubTitleBackground
        {
            get { return subtitlebackground; }
            set { subtitlebackground = value; }
        }

        /// <summary>
        /// ตัวหนาของ Title
        /// </summary>
        private Pen subtitleborder;
        public Pen SubTitleBorder
        {
            get { return subtitleborder; }
            set { subtitleborder = value; }
        }

        #endregion

        // Footer
        #region footer properties

        // override ของ footer
        bool overridefooterformat = false;

        /// <summary>
        /// ส่วนท้าย ค่าเริ่มต้นว่างเปล่า
        /// </summary>
        private String footer;
        public String Footer
        {
            get { return footer; }
            set { footer = value; }
        }

        /// <summary>
        /// แบบอักษรสำหรับส่วนท้าย ค่าเริ่มต้นคือ TH SarabunPSK, 10pt
        /// </summary>
        private Font footerfont;
        public Font FooterFont
        {
            get { return footerfont; }
            set { footerfont = value; }
        }

        /// <summary>
        /// สี Foreground ของส่วนท้าย ค่าเริ่มต้น Black
        /// </summary>
        private Color footercolor;
        public Color FooterColor
        {
            get { return footercolor; }
            set { footercolor = value; }
        }

        /// <summary>
        /// อนุญาต override ของรูปแบบเซลล์ส่วนหัว (header cell)
        /// </summary>
        private StringFormat footerformat;
        public StringFormat FooterFormat
        {
            get { return footerformat; }
            set { footerformat = value; overridefooterformat = true; }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override การจัดตำแหน่ง footer ค่าเริ่มต้นคือ Alignment - Center; 
        /// </summary>
        public StringAlignment FooterAlignment
        {
            get { return footerformat.Alignment; }
            set
            {
                footerformat.Alignment = value;
                overridefooterformat = true;
            }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override ค่าสถานะรูปแบบของส่วนท้าย ค่าเริ่มต้น : FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        public StringFormatFlags FooterFormatFlags
        {
            get { return footerformat.FormatFlags; }
            set
            {
                footerformat.FormatFlags = value;
                overridefooterformat = true;
            }
        }

        /// <summary>
        /// ระยะห่าง ระหว่างตารางและส่วนท้าย
        /// </summary>
        private float footerspacing;
        public float FooterSpacing
        {
            get { return footerspacing; }
            set { footerspacing = value; }
        }

        /// <summary>
        /// ตำแหน่งในเอกสาร title ที่พิมพ์
        /// </summary>
        private PrintLocation footerprint = PrintLocation.All;
        public PrintLocation FooterPrint
        {
            get { return footerprint; }
            set { footerprint = value; }
        }

        /// <summary>
        /// กำหนดความสูงของส่วนท้าย
        /// </summary>
        private float FooterHeight
        {
            get
            {
                float footerheight = 0;

                // return ความสูงที่คำนวณได้ ถ้าเราพิมพ์ส่วนท้าย
                if ((PrintLocation.All == FooterPrint)
                    || ((PrintLocation.FirstOnly == FooterPrint) && (1 == CurrentPage))
                    || ((PrintLocation.LastOnly == FooterPrint) && (totalpages == CurrentPage)))
                {
                    // เพิ่มความสูงของข้อความท้ายกระดาษ
                    footerheight += footerHeight + FooterSpacing;
                }

                return footerheight;
            }
        }

        /// <summary>
        /// สีพื้นหลังของ Title
        /// </summary>
        private Brush footerbackground;
        public Brush FooterBackground
        {
            get { return footerbackground; }
            set { footerbackground = value; }
        }

        /// <summary>
        /// ความหนาของ Title
        /// </summary>
        private Pen footerborder;
        public Pen FooterBorder
        {
            get { return footerborder; }
            set { footerborder = value; }
        }

        #endregion

        // การกำหนดหมายเลขหน้า
        #region page number properties

        // override หมายเลขหน้า
        bool overridepagenumberformat = false;

        /// <summary>
        /// รวมหมายเลขหน้าในงานที่พิมพ์ ค่าเริ่มต้นเป็น true
        /// </summary>
        private bool pageno = true;
        public bool PageNumbers
        {
            get { return pageno; }
            set { pageno = value; }
        }

        /// <summary>
        /// แบบอักษรของหมายเลขหน้าค่าเริ่มต้นคือ TH SarabunPSK, 8pt
        /// </summary>
        private Font pagenofont;
        public Font PageNumberFont
        {
            get { return pagenofont; }
            set { pagenofont = value; }
        }

        /// <summary>
        /// สีข้อความ (เบื้องหน้า) ของหมายเลขหน้า ค่าเริ่มต้น Black
        /// </summary>
        private Color pagenocolor;
        public Color PageNumberColor
        {
            get { return pagenocolor; }
            set { pagenocolor = value; }
        }

        /// <summary>
        /// อนุญาต override ของรูปแบบobject เซลล์ส่วนหัว (header cell)
        /// </summary>
        private StringFormat pagenumberformat;
        public StringFormat PageNumberFormat
        {
            get { return pagenumberformat; }
            set { pagenumberformat = value; overridepagenumberformat = true; }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override การจัดตำแหน่งหมายเลขหน้า ค่าเริ่มต้นคือ Alignment - Near;
        /// </summary>
        public StringAlignment PageNumberAlignment
        {
            get { return pagenumberformat.Alignment; }
            set
            {
                pagenumberformat.Alignment = value;
                overridepagenumberformat = true;
            }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้ override ค่าสถานะการจัดรูปแบบหมายเลขหน้า ค่าเริ่มต้น : FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        public StringFormatFlags PageNumberFormatFlags
        {
            get { return pagenumberformat.FormatFlags; }
            set
            {
                pagenumberformat.FormatFlags = value;
                overridepagenumberformat = true;
            }
        }

        /// <summary>
        /// อนุญาตให้ผู้ใช้เลือกว่าจะมีหมายเลขหน้าอยู่ด้านบนหรือด้านล่างของหน้า 
        /// ค่าเริ่มต้นเป็น false;: หมายเลขหน้า จะอยู่ด้านล่างของหน้า
        /// </summary>
        private bool pagenumberontop = false;
        public bool PageNumberInHeader
        {
            get { return pagenumberontop; }
            set { pagenumberontop = value; }
        }

        /// <summary>
        /// ควรพิมพ์หมายเลขหน้าบนบรรทัดแยกต่างหาก หรือพิมพ์บนบรรทัดเดียวกับส่วนหัว/ส่วนท้าย ค่าเริ่มต้นเป็น false; 
        /// </summary>
        private bool pagenumberonseparateline = false;
        public bool PageNumberOnSeparateLine
        {
            get { return pagenumberonseparateline; }
            set { pagenumberonseparateline = value; }
        }

        /// <summary>
        /// แสดงหมายเลขหน้าทั้งหมดเป็นจำนวนรวม
        /// </summary>
        private bool showtotalpagenumber = false;
        public bool ShowTotalPageNumber
        {
            get { return showtotalpagenumber; }
            set { showtotalpagenumber = value; }
        }

        /// <summary>
        /// ข้อความที่คั่นหมายเลขหน้าและหมายเลขหน้าทั้งหมด ค่าเริ่มต้นเป็น ' of '
        /// </summary>
        private String pageseparator = " of ";
        public String PageSeparator
        {
            get { return pageseparator; }
            set { pageseparator = value; }
        }

        private String pagetext = "Page ";
        public String PageText
        {
            get { return pagetext; }
            set { pagetext = value; }
        }

        private String parttext = " - Part ";
        public String PartText
        {
            get { return parttext; }
            set { parttext = value; }
        }

        /// <summary>
        /// ตำแหน่งในเอกสาร title ที่พิมพ์
        /// </summary>
        private PrintLocation pagenumberprint = PrintLocation.All;
        public PrintLocation PageNumberPrint
        {
            get { return pagenumberprint; }
            set { pagenumberprint = value; }
        }

        /// <summary>
        /// กำหนดความสูงของส่วนท้าย
        /// </summary>
        private float PageNumberHeight
        {
            get
            {
                // return ความสูงที่คำนวณได้ ถ้าเราพิมพ์ส่วนท้าย
                if ((PrintLocation.All == PageNumberPrint)
                    || ((PrintLocation.FirstOnly == PageNumberPrint) && (1 == CurrentPage))
                    || ((PrintLocation.LastOnly == PageNumberPrint) && (totalpages == CurrentPage)))
                {
                    // return ความสูงของจำนวนหน้า ถ้าเรากำลังพิมพ์ บรรทัดแยก
                    // ถ้าเราไม่ได้พิมพ์บนบรรทัดที่แยกจากกัน จะหยุด!!
                    // หัวกระดาษหรือท้ายกระดาษแล้วเรายังต้องสำรองพื้นที่สำหรับหมายเลขหน้า
                    if (pagenumberonseparateline)
                        return pagenumberHeight;
                    else if (pagenumberontop && 0 == TitleHeight && 0 == SubTitleHeight)
                    {
                        return pagenumberHeight;
                    }
                    else if (!pagenumberontop && 0 == FooterHeight)
                    {
                        return footerspacing + pagenumberHeight;
                    }
                }

                return 0;
            }
        }

        #endregion

        // การพิมพ์เซลล์ส่วนหัว
        #region header cell properties

        private DataGridViewCellStyle rowheaderstyle;
        public DataGridViewCellStyle RowHeaderCellStyle
        {
            get { return rowheaderstyle; }
            set { rowheaderstyle = value; }
        }

        /// <summary>
        /// อนุญาตให้แทนที่รูปแบบเซลล์ส่วนหัวของแถว
        /// </summary>
        private StringFormat rowheadercellformat = null;
        public StringFormat GetRowHeaderCellFormat(DataGridView grid)
        {
            // get default values from provided data grid view, but only
            // รับค่าเริ่มต้นจากมุมมองตารางข้อมูลที่ให้ไว้ แต่เท่านั้น
            // if we don't already have a header cell format
            // หากเรายังไม่มีรูปแบบเซลล์ส่วนหัว
            if ((null != grid) && (null == rowheadercellformat))
            {
                buildstringformat(ref rowheadercellformat, grid.Rows[0].HeaderCell.InheritedStyle,
                    headercellalignment, StringAlignment.Near, headercellformatflags,
                    StringTrimming.Word);
            }

            // if we still don't have a header cell format, create an empty
            // หากเรายังไม่มีรูปแบบเซลล์ส่วนหัวให้สร้างที่ว่างเปล่า
            if (null == rowheadercellformat)
                rowheadercellformat = new StringFormat(headercellformatflags);

            return rowheadercellformat;
        }

        /// <summary>
        /// Default value to show in the row header cell if no value is provided in the DataGridView.
        /// ค่าเริ่มต้นที่จะแสดงในเซลล์ส่วนหัวของแถวหากไม่มีการระบุค่าไว้ใน DataGridView
        /// Defaults to one tab space
        /// เริ่มต้นที่หนึ่งแท็บพื้นที่
        /// </summary>
        private String rowheadercelldefaulttext = "\t";
        public String RowHeaderCellDefaultText
        {
            get { return rowheadercelldefaulttext; }
            set { rowheadercelldefaulttext = value; }
        }

        /// <summary>
        /// Allow override of the header cell format object
        /// อนุญาตการแทนที่ของวัตถุรูปแบบเซลล์ส่วนหัว
        /// </summary>
        private Dictionary<string, DataGridViewCellStyle> columnheaderstyles =
            new Dictionary<string, DataGridViewCellStyle>();
        public Dictionary<string, DataGridViewCellStyle> ColumnHeaderStyles
        {
            get { return columnheaderstyles; }
        }

        /// <summary>
        /// Allow override of the header cell format object
        /// อนุญาตการแทนที่ของวัตถุรูปแบบเซลล์ส่วนหัว
        /// </summary>
        private StringFormat columnheadercellformat = null;
        public StringFormat GetColumnHeaderCellFormat(DataGridView grid)
        {
            // get default values from provided data grid view, but only
            // รับค่าเริ่มต้นจากมุมมองตารางข้อมูลที่ให้ไว้ แต่เท่านั้น
            // if we don't already have a header cell format
            // หากเรายังไม่มีรูปแบบเซลล์ส่วนหัว
            if ((null != grid) && (null == columnheadercellformat))
            {
                buildstringformat(ref columnheadercellformat, grid.Columns[1].HeaderCell.InheritedStyle,
                    headercellalignment, StringAlignment.Near, headercellformatflags,
                    StringTrimming.Word);
            }

            // if we still don't have a header cell format, create an empty
            // หากเรายังไม่มีรูปแบบเซลล์ส่วนหัวให้สร้างที่ว่างเปล่า
            if (null == columnheadercellformat)
                columnheadercellformat = new StringFormat(headercellformatflags);

            return columnheadercellformat;
        }

        /// <summary>
        /// Deprecated - use HeaderCellFormat
        /// เลิกใช้ - ใช้ HeaderCellFormat
        /// Allow the user to override the header cell string alignment. Default value is 
        /// อนุญาตให้ผู้ใช้แทนที่การจัดตำแหน่งสตริงเซลล์ส่วนหัว ค่าเริ่มต้นคือ
        /// Alignment - Near; 
        /// การจัดตำแหน่ง - ใกล้;
        /// </summary>
        private StringAlignment headercellalignment;
        public StringAlignment HeaderCellAlignment
        {
            get { return headercellalignment; }
            set { headercellalignment = value; }
        }

        /// <summary>
        /// Deprecated - use HeaderCellFormat
        /// เลิกใช้ - ใช้ HeaderCellFormat
        /// Allow the user to override the header cell string format flags. Default values
        /// อนุญาตให้ผู้ใช้แทนที่ค่าสถานะการจัดรูปแบบสตริงของเซลล์ส่วนหัว ค่าเริ่มต้น
        /// are: FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        private StringFormatFlags headercellformatflags;
        public StringFormatFlags HeaderCellFormatFlags
        {
            get { return headercellformatflags; }
            set { headercellformatflags = value; }
        }
        #endregion

        // การพิมพ์แต่ละเซลล์
        #region cell properties

        /// <summary>
        /// อนุญาตให้ override รูปแบบการพิมพ์ของเซลล์
        /// </summary>
        private StringFormat cellformat = null;
        public StringFormat GetCellFormat(DataGridView grid)
        {
            // get default values from provided data grid view, but only
            // รับค่าเริ่มต้นจากมุมมองตารางข้อมูลที่ให้ไว้ แต่เท่านั้น
            // if we don't already have a cell format
            // หากเรายังไม่มีรูปแบบเซลล์
            if ((null != grid) && (null == cellformat))
            {
                buildstringformat(ref cellformat, grid.Rows[0].Cells[0].InheritedStyle,
                    cellalignment, StringAlignment.Near, cellformatflags,
                    StringTrimming.Word);
            }

            // if we still don't have a cell format, create an empty
            // หากเรายังไม่มีรูปแบบเซลล์ให้สร้างที่ว่างเปล่า
            if (null == cellformat)
                cellformat = new StringFormat(cellformatflags);

            return cellformat;
        }

        /// <summary>
        /// Deprecated - use GetCellFormat
        /// เลิกใช้ - ใช้ GetCellFormat
        /// Allow the user to override the cell string alignment. Default value is 
        /// อนุญาตให้ผู้ใช้แทนที่การจัดตำแหน่งสตริงเซลล์ ค่าเริ่มต้นคือ
        /// Alignment - Near; 
        /// การจัดตำแหน่ง - ใกล้;
        /// </summary>
        private StringAlignment cellalignment;
        public StringAlignment CellAlignment
        {
            get { return cellalignment; }
            set { cellalignment = value; }
        }

        /// <summary>
        /// Deprecated - use GetCellFormat
        /// เลิกใช้ - ใช้ GetCellFormat
        /// Allow the user to override the cell string format flags. Default values
        /// อนุญาตให้ผู้ใช้แทนที่ค่าสถานะการจัดรูปแบบสตริงของเซลล์ ค่าเริ่มต้น
        /// are: FormatFlags - NoWrap, LineLimit, NoClip
        /// </summary>
        private StringFormatFlags cellformatflags;
        public StringFormatFlags CellFormatFlags
        {
            get { return cellformatflags; }
            set { cellformatflags = value; }
        }

        /// <summary>
        /// allow the user to override the column width calcs with their own defaults
        /// อนุญาตให้ผู้ใช้แทนที่ความกว้างคอลัมน์ด้วยค่าเริ่มต้นของตนเอง
        /// </summary>
        private List<float> colwidthsoverride = new List<float>();
        private Dictionary<string, float> publicwidthoverrides = new Dictionary<string, float>();
        public Dictionary<string, float> ColumnWidths
        {
            get { return publicwidthoverrides; }
        }

        /// <summary>
        /// Allow per column style overrides
        /// อนุญาตการแทนที่สไตล์คอลัมน์
        /// </summary>
        private Dictionary<string, DataGridViewCellStyle> colstyles =
            new Dictionary<string, DataGridViewCellStyle>();
        public Dictionary<string, DataGridViewCellStyle> ColumnStyles
        {
            get { return colstyles; }
        }

        /// <summary>
        /// Allow per column style overrides
        /// อนุญาตการแทนที่สไตล์คอลัมน์
        /// </summary>
        private Dictionary<string, DataGridViewCellStyle> altrowcolstyles =
            new Dictionary<string, DataGridViewCellStyle>();
        public Dictionary<string, DataGridViewCellStyle> AlternatingRowColumnStyles
        {
            get { return altrowcolstyles; }
        }

        /// <summary>
        /// Allow the user to set columns that appear on every pageset. Only used when 
        /// อนุญาตให้ผู้ใช้ตั้งค่าคอลัมน์ที่ปรากฏในทุก ๆ หน้า ใช้เฉพาะเมื่อ
        /// the printout is wider than one page.
        /// งานพิมพ์กว้างกว่าหน้าเดียว
        /// </summary>
        private List<int> fixedcolumns = new List<int>();
        private List<string> fixedcolumnnames = new List<string>();
        public List<string> FixedColumns
        {
            get { return fixedcolumnnames; }
        }

        /// <summary>
        /// List of columns to not display in the grid view printout.
        /// รายการของคอลัมน์ที่ไม่แสดงในมุมมองกริดที่พิมพ์ออกมา
        /// </summary>
        private List<String> hidecolumns = new List<string>();
        public List<String> HideColumns
        {
            get { return hidecolumns; }
        }

        /// <summary>
        /// Insert a page break when the value in this column changes
        /// แทรกตัวแบ่งหน้าเมื่อค่าในคอลัมน์นี้เปลี่ยนแปลง
        /// </summary>
        private object oldvalue = null;
        private String breakonvaluechange;
        public String BreakOnValueChange
        {
            get { return breakonvaluechange; }
            set { breakonvaluechange = value; }
        }

        #endregion

        // ระดับหน้า
        #region page level properties

        /// <summary>
        /// ระยะขอบของหน้า ค่าเริ่มต้นคือ (60, 60, 40, 40)
        /// </summary>
        public Margins PrintMargins
        {
            get { return PageSettings.Margins; }
            set { PageSettings.Margins = value; }
        }

        /// <summary>
        /// Expose the printdocument default page settings to the caller
        /// แสดงการตั้งค่าหน้าเริ่มต้นของ printdocument แก่ผู้โทร
        /// </summary>
        public PageSettings PageSettings
        {
            get { return printDoc.DefaultPageSettings; }
        }

        /// <summary>
        /// Spread the columns porportionally accross the page. Default is false.
        /// กระจายคอลัมน์ตามสัดส่วนในหน้า ค่าเริ่มต้นเป็นเท็จ
        /// Deprecated. Please use the ColumnWidth property
        /// เลิก กรุณาใช้คุณสมบัติ ColumnWidth
        /// </summary>
        private bool porportionalcolumns = false;
        public bool PorportionalColumns
        {
            get { return porportionalcolumns; }
            set
            {
                porportionalcolumns = value;
                if (porportionalcolumns)
                    ColumnWidth = ColumnWidthSetting.Porportional;
                else
                    ColumnWidth = ColumnWidthSetting.CellWidth;
            }
        }

        /// <summary>
        /// จัดกึ่งกลางตารางของหน้า
        /// </summary>
        private Alignment tablealignment = Alignment.NotSet;
        public Alignment TableAlignment
        {
            get { return tablealignment; }
            set { tablealignment = value; }
        }

        /// <summary>
        /// Change the default row height to either the height of the string or the size of 
        /// เปลี่ยนความสูงของแถวเริ่มต้นเป็นความสูงของสตริงหรือขนาดของ
        /// the cell. Added for image cell handling; set to CellHeight for image cells
        /// เซลล์ เพิ่มสำหรับการจัดการเซลล์รูปภาพ ตั้งค่าเป็น CellHeight สำหรับเซลล์รูปภาพ
        /// </summary>
        public enum RowHeightSetting { DataHeight, CellHeight }
        private RowHeightSetting _rowheight = RowHeightSetting.DataHeight;
        public RowHeightSetting RowHeight
        {
            get { return _rowheight; }
            set { _rowheight = value; }
        }

        /// <summary>
        /// Change the default column width to be spread porportionally accross the page,
        /// เปลี่ยนความกว้างคอลัมน์เริ่มต้นที่จะแพร่กระจายตามส่วนของหน้า
        /// to the size of the grid cell or the size of the formatted data string.
        /// ตามขนาดของเซลล์กริดหรือขนาดของสตริงข้อมูลที่จัดรูปแบบ
        /// Set to CellWidth for image cells.
        /// ตั้งค่าเป็น CellWidth สำหรับเซลล์รูปภาพ
        /// </summary>
        public enum ColumnWidthSetting { DataWidth, CellWidth, Porportional }
        private ColumnWidthSetting _rowwidth = ColumnWidthSetting.CellWidth;
        public ColumnWidthSetting ColumnWidth
        {
            get { return _rowwidth; }
            set
            {
                _rowwidth = value;
                if (value == ColumnWidthSetting.Porportional)
                    porportionalcolumns = true;
                else
                    porportionalcolumns = false;
            }
        }

        #endregion

        // Utility Functions 
        #region
        /// <summary>
        /// calculate the print preview window width to show the entire page
        /// คำนวณความกว้างของหน้าต่างตัวอย่างก่อนพิมพ์เพื่อแสดงทั้งหน้า
        /// </summary>
        /// <returns></returns>
        private int PreviewDisplayWidth()
        {
            double displayWidth = printDoc.DefaultPageSettings.Bounds.Width
                + 3 * printDoc.DefaultPageSettings.HardMarginY;
            return (int)(displayWidth * PrintPreviewZoom);
        }

        /// <summary>
        /// calculate the print preview window height to show the entire page
        /// คำนวณความสูงของหน้าต่างตัวอย่างก่อนพิมพ์เพื่อแสดงทั้งหน้า
        /// </summary>
        /// <returns></returns>
        private int PreviewDisplayHeight()
        {
            double displayHeight = printDoc.DefaultPageSettings.Bounds.Height
                + 3 * printDoc.DefaultPageSettings.HardMarginX;

            return (int)(displayHeight * PrintPreviewZoom);
        }

        /// <summary>
        /// การวาดเซลล์
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnCellOwnerDraw(DGVCellDrawingEventArgs e)
        {
            if (null != OwnerDraw)
                OwnerDraw(this, e);
        }

        /// <summary>
        /// Given a row and column, get the current grid cell style, including our local 
        /// รับแถวและคอลัมน์รับสไตล์กริดเซลล์ปัจจุบันรวมถึงในพื้นที่ของเรา
        /// overrides
        /// แทนที่
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        protected DataGridViewCellStyle GetStyle(DataGridViewRow row, DataGridViewColumn col)
        {
            // set initial default ตั้งค่าเริ่มต้น
            DataGridViewCellStyle colstyle = row.Cells[col.Index].InheritedStyle.Clone();

            // check for our override ตรวจสอบการแทนที่ของเรา
            if (ColumnStyles.ContainsKey(col.Name))
            {
                colstyle = ColumnStyles[col.Name];
            }

            // check for alternating row override ตรวจสอบการแทนที่แถว
            if (0 != (row.Index & 1) && AlternatingRowColumnStyles.ContainsKey(col.Name))
            {
                colstyle = AlternatingRowColumnStyles[col.Name];
            }

            return colstyle;
        }

        /// <summary>
        /// Skim the colstoprint list for a column name and return it's index
        /// อ่านรายชื่อ colstoprint สำหรับชื่อคอลัมน์และส่งคืนดัชนี
        /// </summary>
        /// <param name="colname">Name of column to find</param>
        /// <returns>index of column</returns> ดัชนีของคอลัมน์
        protected int GetColumnIndex(string colname)
        {
            int i = 0;
            foreach (DataGridViewColumn col in colstoprint)
            {
                if (col.Name != colname)
                {
                    i++;
                }
                else
                {
                    break;
                }
            }

            // ตรวจสอบชื่อคอลัมน์
            if (i >= colstoprint.Count)
            {
                throw new Exception("Unknown Column Name: " + colname);
            }

            return i;
        }

        #endregion

        #endregion

        /// <summary>
        /// ตัวสร้างสำหรับ DGVPrinter
        /// </summary>
        public DGVPrinter()
        {
            // สร้างเอกสารที่จะพิมพ์
            printDoc = new PrintDocument();
            //printDoc.PrintPage += new PrintPageEventHandler(PrintPageEventHandler);
            //printDoc.BeginPrint += new PrintEventHandler(BeginPrintEventHandler);
            PrintMargins = new Margins(60, 60, 40, 40);

            // ตั้งค่ารูปแบบตัวอักษร
            pagenofont = new Font("TH SarabunPSK", 8, FontStyle.Regular, GraphicsUnit.Point);
            pagenocolor = Color.Black;
            titlefont = new Font("TH SarabunPSK", 18, FontStyle.Bold, GraphicsUnit.Point);
            titlecolor = Color.Black;
            subtitlefont = new Font("TH SarabunPSK", 12, FontStyle.Bold, GraphicsUnit.Point);
            subtitlecolor = Color.Black;
            footerfont = new Font("TH SarabunPSK", 10, FontStyle.Bold, GraphicsUnit.Point);
            footercolor = Color.Black;

            // ระยะห่างเริ่มต้น
            titlespacing = 0;
            subtitlespacing = 0;
            footerspacing = 0;

            // การจัดรูปแบบ
            buildstringformat(ref titleformat, null, StringAlignment.Center, StringAlignment.Center,
                StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip, StringTrimming.Word);
            buildstringformat(ref subtitleformat, null, StringAlignment.Center, StringAlignment.Center,
                StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip, StringTrimming.Word);
            buildstringformat(ref footerformat, null, StringAlignment.Center, StringAlignment.Center,
                StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip, StringTrimming.Word);
            buildstringformat(ref pagenumberformat, null, StringAlignment.Far, StringAlignment.Center,
                StringFormatFlags.NoWrap | StringFormatFlags.LineLimit | StringFormatFlags.NoClip, StringTrimming.Word);

            // ตั้งค่าออบเจ็กต์การจัดรูปแบบเป็น null เพื่อตั้งค่าสถานะว่าถูกตั้งค่าหรือไม่
            columnheadercellformat = null;
            rowheadercellformat = null;
            cellformat = null;

            // ตัวอย่างก่อนพิมพ์
            Owner = null;
            PrintPreviewZoom = 1.0;

            // properties ที่เลิกใช้แล้ว - เก็บไว้ เพื่อความเข้ากันได้
            headercellalignment = StringAlignment.Near;
            headercellformatflags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            cellalignment = StringAlignment.Near;
            cellformatflags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
        }

        // อินเทอร์เฟซหลัก - แสดงกล่องโต้ตอบ จากนั้นพิมพ์ + ดูตัวอย่าง datagridview

        /// <summary>
        /// เริ่มกระบวนการพิมพ์ --> พิมพ์ไปที่เครื่องพิมพ์
        /// </summary>
        /// <param name="dgv">DataGridView ที่จะพิมพ์</param>
        /// หมายเหตุ: ถ้ามีการเปลี่ยนแปลง กับวิธีนี้ยังต้องทำใน PrintPreviewDataGridView
        public void PrintDataGridView(DataGridView dgv)
        {
            if (EnableLogging) Logger.LogInfoMsg("PrintDataGridView process started");
            if (null == dgv) throw new Exception("Null Parameter passed to DGVPrinter.");
            if (!(typeof(DataGridView).IsInstanceOfType(dgv)))
                throw new Exception("Invalid Parameter passed to DGVPrinter.");

            // บันทึก DataGridview ที่เรากำลังพิมพ์
            this.dgv = dgv;

            // แสดงข้อความโต้ตอบ & พิมพ์
            if (DialogResult.OK == DisplayPrintDialog())
            {
                PrintNoDisplay(dgv);
            }
        }

        /// <summary>
        /// เริ่มกระบวนการพิมพ์พิมพ์ไปที่กล่องโต้ตอบตัวอย่างก่อนพิมพ์
        /// </summary>
        /// <param name="dgv">The DataGridView to print</param>
        /// หมายเหตุ: การเปลี่ยนแปลงใด ๆ ต้องทำใน PrintDataGridView
        public void PrintPreviewDataGridView(DataGridView dgv)
        {
            if (EnableLogging) Logger.LogInfoMsg("PrintPreviewDataGridView process started");
            if (null == dgv) throw new Exception("Null Parameter passed to DGVPrinter.");
            if (!(typeof(DataGridView).IsInstanceOfType(dgv)))
                throw new Exception("Invalid Parameter passed to DGVPrinter.");

            // บันทึก DataGridview ที่เรากำลังพิมพ์
            this.dgv = dgv;

            // แสดงข้อความโต้ตอบและพิมพ์
            if (DialogResult.OK == DisplayPrintDialog())
            {
                PrintPreviewNoDisplay(dgv);
            }
        }

        /// <summary>
        /// แสดง printdialog และส่งคืนผลลัพธ์ อาจเป็นได้ทั้งวิธีนี้หรือ
        /// ต้องทำการเทียบเท่าก่อนที่จะเรียกใช้ PrintNoDisplay / PrintPreviewNoDisplay methods
        /// </summary>
        /// <returns></returns>
        public DialogResult DisplayPrintDialog()
        {
            if (EnableLogging) Logger.LogInfoMsg("DisplayPrintDialog process started");
            // สร้างกล่องโต้ตอบการพิมพ์ใหม่และตั้งค่าตัวเลือก
            PrintDialog pd = new PrintDialog();
            pd.UseEXDialog = printDialogSettings.UseEXDialog;
            pd.AllowSelection = printDialogSettings.AllowSelection;
            pd.AllowSomePages = printDialogSettings.AllowSomePages;
            pd.AllowCurrentPage = printDialogSettings.AllowCurrentPage;
            pd.AllowPrintToFile = printDialogSettings.AllowPrintToFile;
            pd.ShowHelp = printDialogSettings.ShowHelp;
            pd.ShowNetwork = printDialogSettings.ShowNetwork;

            // ตั้งค่ากล่องโต้ตอบการพิมพ์ด้วยการตั้งค่าภายใน
            pd.Document = printDoc;
            if (!String.IsNullOrEmpty(printerName))
                printDoc.PrinterSettings.PrinterName = printerName;

            // แสดงข้อความและแสดงผลลัพธ์
            return pd.ShowDialog();
        }

        /// <summary>
        /// พิมพ์มุมมองตารางที่ให้ไว้ DisplayPrintDialog () หรือเทียบเท่า
        /// การตั้งค่าจะต้องเสร็จสิ้นก่อนที่จะเรียก
        /// </summary>
        /// <param name="dgv"></param>
        public void PrintNoDisplay(DataGridView dgv)
        {
            if (EnableLogging) Logger.LogInfoMsg("PrintNoDisplay process started");
            if (null == dgv) throw new Exception("Null Parameter passed to DGVPrinter.");
            if (!(dgv is DataGridView))
                throw new Exception("Invalid Parameter passed to DGVPrinter.");

            // บันทึกตารางที่เรากำลังพิมพ์
            this.dgv = dgv; 

            printDoc.PrintPage += new PrintPageEventHandler(PrintPageEventHandler);
            printDoc.BeginPrint += new PrintEventHandler(BeginPrintEventHandler);

            // setup + print
            SetupPrint();
            printDoc.Print();
        }

        /// <summary>
        /// ดูตัวอย่างตารางที่ให้ไว้ DisplayPrintDialog () หรือเทียบเท่า
        /// การตั้งค่าจะต้องเสร็จสิ้นก่อนที่จะเรียก
        /// </summary>
        /// <param name="dgv"></param>
        public void PrintPreviewNoDisplay(DataGridView dgv)
        {
            if (EnableLogging) Logger.LogInfoMsg("PrintPreviewNoDisplay process started");
            if (null == dgv) throw new Exception("Null Parameter passed to DGVPrinter.");
            if (!(dgv is DataGridView))
                throw new Exception("Invalid Parameter passed to DGVPrinter.");

            // บันทึกตารางที่เรากำลังพิมพ์
            this.dgv = dgv;

            printDoc.PrintPage += new PrintPageEventHandler(PrintPageEventHandler);
            printDoc.BeginPrint += new PrintEventHandler(BeginPrintEventHandler);

            // แสดงข้อความแสดงตัวอย่าง
            SetupPrint();

            // หากผู้ใช้ไม่ได้เตรียมกล่องโต้ตอบตัวอย่างก่อนพิมพ์ ให้สร้างขึ้นมา
            if (null == PreviewDialog)
                PreviewDialog = new PrintPreviewDialog();

            // ตั้งค่าโต้ตอบเพื่อดูตัวอย่าง
            PreviewDialog.Document = printDoc;
            PreviewDialog.UseAntiAlias = true;
            PreviewDialog.Owner = Owner;
            PreviewDialog.PrintPreviewControl.Zoom = PrintPreviewZoom;
            PreviewDialog.Width = PreviewDisplayWidth();
            PreviewDialog.Height = PreviewDisplayHeight();

            if (null != ppvIcon)
                PreviewDialog.Icon = ppvIcon;

            // แสดงข้อความ
            PreviewDialog.ShowDialog();
        }


        //---------------------------------------------------------------------
        //---------------------------------------------------------------------
        // Print Process Interface Methods
        // วิธีการพิมพ์อินเตอร์เฟสกระบวนการ
        //---------------------------------------------------------------------
        //---------------------------------------------------------------------

        // NOTE: This is retained only for backward compatibility, and should 
        // หมายเหตุ: สิ่งนี้จะถูกเก็บไว้สำหรับความเข้ากันได้แบบย้อนหลังเท่านั้นและควร
        // not be used for printing grid views that might be larger than the 
        // ไม่ใช้สำหรับการพิมพ์มุมมองตารางที่อาจมีขนาดใหญ่กว่า
        // input print area. พื้นที่พิมพ์อินพุต
        public Boolean EmbeddedPrint(DataGridView dgv, Graphics g, Rectangle area)
        {
            if (EnableLogging) Logger.LogInfoMsg("EmbeddedPrint process started");
            // verify we've been set up properly
            // ยืนยันว่าเราได้ติดตั้งอย่างถูกต้องแล้ว
            if ((null == dgv))
                throw new Exception("Null Parameter passed to DGVPrinter.");

            // set the embedded print flag ตั้งค่าสถานะการพิมพ์ฝังตัว
            EmbeddedPrinting = true;

            // save the grid we're printing บันทึกกริดที่เรากำลังพิมพ์
            this.dgv = dgv;

            //-----------------------------------------------------------------
            // Force setting for embedded printing บังคับการตั้งค่าสำหรับการพิมพ์แบบฝัง
            //-----------------------------------------------------------------

            // set margins so we print within the provided area
            // กำหนดระยะขอบเพื่อให้เราพิมพ์ภายในพื้นที่ที่จัดไว้
            Margins saveMargins = PrintMargins;
            PrintMargins.Top = area.Top;
            PrintMargins.Bottom = 0;
            PrintMargins.Left = area.Left;
            PrintMargins.Right = 0;

            // set "page" height and width to our destination area
            // กำหนด "หน้า" ความสูงและความกว้างเป็นพื้นที่ปลายทางของเรา
            pageHeight = area.Height + area.Top;
            printWidth = area.Width;
            pageWidth = area.Width + area.Left;

            // force 'off' header and footer บังคับให้ปิดส่วนหัวและส่วนท้าย
            PrintHeader = false;
            PrintFooter = false;
            pageno = false;

            //-----------------------------------------------------------------
            // Determine what's going to be printed and set the columns to print
            // กำหนดสิ่งที่จะพิมพ์และตั้งค่าคอลัมน์ที่จะพิมพ์
            //-----------------------------------------------------------------
            SetupPrint();

            //-----------------------------------------------------------------
            // Do a single "Print" and return false - we're just printing what
            // ทำ "พิมพ์" เพียงครั้งเดียวและส่งคืนค่าเท็จ - เราแค่พิมพ์อะไร
            // we can in the space provided.
            // เราทำได้ในพื้นที่ที่จัดไว้ให้
            //-----------------------------------------------------------------
            PrintPage(g);
            return false;
        }

        public void EmbeddedPrintMultipageSetup(DataGridView dgv, Rectangle area)
        {
            if (EnableLogging) Logger.LogInfoMsg("EmbeddedPrintMultipageSetup process started");
            // verify we've been set up properly ยืนยันว่าเราได้ติดตั้งอย่างถูกต้องแล้ว
            if ((null == dgv))
                throw new Exception("Null Parameter passed to DGVPrinter.");

            // set the embedded print flag ตั้งค่าสถานะการพิมพ์ฝังตัว
            EmbeddedPrinting = true;

            // save the grid we're printing บันทึกกริดที่เรากำลังพิมพ์
            this.dgv = dgv;

            //-----------------------------------------------------------------
            // Force setting for embedded printing
            // บังคับการตั้งค่าสำหรับการพิมพ์แบบฝัง
            //-----------------------------------------------------------------

            // set margins so we print within the provided area
            // กำหนดระยะขอบเพื่อให้เราพิมพ์ภายในพื้นที่ที่จัดไว้
            Margins saveMargins = PrintMargins;
            PrintMargins.Top = area.Top;
            PrintMargins.Bottom = 0;
            PrintMargins.Left = area.Left;
            PrintMargins.Right = 0;

            // set "page" height and width to our destination area
            // กำหนด "หน้า" ความสูงและความกว้างเป็นพื้นที่ปลายทางของเรา
            pageHeight = area.Height + area.Top;
            printWidth = area.Width;
            pageWidth = area.Width + area.Left;

            // force 'off' header and footer บังคับให้ปิดส่วนหัวและส่วนท้าย
            PrintHeader = false;
            PrintFooter = false;
            pageno = false;

            //-----------------------------------------------------------------
            // Determine what's going to be printed and set the columns to print
            // กำหนดสิ่งที่จะพิมพ์และตั้งค่าคอลัมน์ที่จะพิมพ์
            //-----------------------------------------------------------------
            SetupPrint();
        }

        /// <summary>
        /// BeginPrint Event Handler ตัวจัดการเหตุการณ์ BeginPrint
        /// Set values at start of print run กำหนดค่าเมื่อเริ่มการพิมพ์
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void BeginPrintEventHandler(object sender, PrintEventArgs e)
        {
            if (EnableLogging) Logger.LogInfoMsg("BeginPrintEventHandler called. Printing started.");
            // reset counters since we'll go through this twice if we print from preview
            // รีเซ็ตตัวนับเนื่องจากเราจะผ่านสองครั้งนี้ถ้าเราพิมพ์จากหน้าตัวอย่าง
            currentpageset = 0;
            lastrowprinted = -1;
            CurrentPage = 0;
        }

        /// <summary>
        /// PrintPage event handler. This routine prints one page. It will
        /// ตัวจัดการเหตุการณ์ PrintPage รูทีนนี้พิมพ์หนึ่งหน้า มันจะ
        /// skip non-printable pages if the user selected the "some pages" option
        /// ข้ามหน้าที่ไม่สามารถพิมพ์ได้หากผู้ใช้เลือกตัวเลือก "บางหน้า"
        /// on the print dialog. ในกล่องโต้ตอบการพิมพ์
        /// </summary>
        /// <param name="sender">default object from windows</param>
        /// <param name="e">Event info from Windows about the printing</param>
        public void PrintPageEventHandler(object sender, PrintPageEventArgs e)
        {
            if (EnableLogging) Logger.LogInfoMsg("PrintPageEventHandler called. Printing a page.");
            e.HasMorePages = PrintPage(e.Graphics);
        }


        //---------------------------------------------------------------------
        //---------------------------------------------------------------------
        // Internal Methods วิธีการภายใน
        //---------------------------------------------------------------------
        //---------------------------------------------------------------------

        /// <summary>
        /// Set up the print job. Save information from print dialog
        /// ตั้งค่างานพิมพ์ บันทึกข้อมูลจากกล่องโต้ตอบการพิมพ์
        /// and print document for easy access. Also sets up the rows
        /// และพิมพ์เอกสารเพื่อให้เข้าถึงได้ง่าย นอกจากนี้ยังตั้งค่าแถว
        /// and columns that will be printed. At this point, we're 
        /// และคอลัมน์ที่จะพิมพ์ ณ จุดนี้เรา
        /// collecting all columns in colstoprint. This will be broken
        /// รวบรวมคอลัมน์ทั้งหมดใน colstoprint สิ่งนี้จะหัก
        /// up into pagesets later on ขึ้นเป็นจำนวนหน้าในภายหลัง
        /// </summary>
        void SetupPrint()
        {
            if (EnableLogging)
            {
                Logger.LogInfoMsg("SetupPrint process started");
                var m = printDoc.DefaultPageSettings.Margins;
                Logger.LogInfoMsg(String.Format("Initial Printer Margins are {0}, {1}, {2}, {3} ", m.Left, m.Right, m.Top, m.Bottom));
            }

            if (null == PrintColumnHeaders)
                PrintColumnHeaders = dgv.ColumnHeadersVisible;

            if (null == PrintRowHeaders)
                PrintRowHeaders = dgv.RowHeadersVisible;

            // Set the default row header style where we don't have an override
            // กำหนดลักษณะส่วนหัวของแถวเริ่มต้นที่เราไม่มีการแทนที่
            // and we do have rows และเรามีแถว
            if ((null == RowHeaderCellStyle) && (0 != dgv.Rows.Count))
                RowHeaderCellStyle = dgv.Rows[0].HeaderCell.InheritedStyle;

            /* Functionality to come - redo of styling ฟังก์ชั่นที่จะมา - ทำซ้ำของสไตล์
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                // Set the default column styles where we've not been given an override
                if (!ColumnStyles.ContainsKey(col.Name))
                    ColumnStyles[col.Name] = dgv.Columns[col.Name].InheritedStyle;
                // Set the default column header styles where we don't have an override
                if (!ColumnHeaderStyles.ContainsKey(col.Name))
                    ColumnHeaderStyles[col.Name] = dgv.Columns[col.Name].HeaderCell.InheritedStyle;
            }
            */

            //-----------------------------------------------------------------
            // Set row and column headercell and normal cell print formats if they were not
            // ตั้งค่าส่วนหัวของแถวและคอลัมน์และรูปแบบการพิมพ์เซลล์ปกติหากไม่มี
            // explicitly set by the caller กำหนดอย่างชัดเจนโดยผู้โทร
            //-----------------------------------------------------------------
            if (null == columnheadercellformat)
                buildstringformat(ref columnheadercellformat, dgv.Columns[0].HeaderCell.InheritedStyle,
                    headercellalignment, StringAlignment.Near, headercellformatflags,
                    StringTrimming.Word);
            if (null == rowheadercellformat)
                buildstringformat(ref rowheadercellformat, RowHeaderCellStyle,
                    headercellalignment, StringAlignment.Near, headercellformatflags,
                    StringTrimming.Word);
            if (null == cellformat)
                buildstringformat(ref cellformat, dgv.DefaultCellStyle,
                    cellalignment, StringAlignment.Near, cellformatflags,
                    StringTrimming.Word);

            // รับข้อมูล ข้อจำกัด ของพื้นที่การพิมพ์จริงของเครื่องพิมพ์ --> แปลง int ทำงานกับระยะขอบ
            // หมายเหตุ: ทำเฉพาะเมื่อเราไม่ได้ทำการพิมพ์แบบฝัง

            if (!EmbeddedPrinting)
            {
                int printareawidth;
                int hardx = (int)Math.Round(printDoc.DefaultPageSettings.HardMarginX);
                int hardy = (int)Math.Round(printDoc.DefaultPageSettings.HardMarginY);
                if (printDoc.DefaultPageSettings.Landscape)
                    printareawidth = (int)Math.Round(printDoc.DefaultPageSettings.PrintableArea.Height);
                else
                    printareawidth = (int)Math.Round(printDoc.DefaultPageSettings.PrintableArea.Width);

                // กำหนดพื้นที่ ที่จะพิมพ์ที่เรากำลังทำงานอยู่ PageSettings (กว้าง+สูง)

                pageHeight = printDoc.DefaultPageSettings.Bounds.Height;
                pageWidth = printDoc.DefaultPageSettings.Bounds.Width;

                // ตั้งค่าพื้นที่ที่สามารถทำการพิมพ์ได้: ระยะขอบ

                // กำหนดระยะขอบของเครื่องพิมพ์ เป็นค่าเริ่มต้น
                PrintMargins = printDoc.DefaultPageSettings.Margins;

                // ปรับสำหรับเมื่อระยะขอบน้อยกว่าขีดจำกัด x / y ของเครื่องพิมพ์
                PrintMargins.Right = (hardx > PrintMargins.Right) ? hardx : PrintMargins.Right;
                PrintMargins.Left = (hardx > PrintMargins.Left) ? hardx : PrintMargins.Left;
                PrintMargins.Top = (hardy > PrintMargins.Top) ? hardy : PrintMargins.Top;
                PrintMargins.Bottom = (hardy > PrintMargins.Bottom) ? hardy : PrintMargins.Bottom;

                // สามารถคำนวณความกว้าง ของการพิมพ์ค่าเริ่มต้นได้อีกครั้ง โดยคำนึงถึงข้อจำกัดของเครื่องพิมพ์
                printWidth = pageWidth - PrintMargins.Left - PrintMargins.Right;
                printWidth = (printWidth > printareawidth) ? printareawidth : printWidth;

                // log margin เปลี่ยนแปลง 
                if (EnableLogging)
                {
                    Logger.LogInfoMsg(String.Format("Printer 'Hard' X limit is {0} and 'Hard' Y limit is {1}", hardx, hardy));
                    Logger.LogInfoMsg(String.Format("Printer height limit is {0} and width limit is {1}, print width is {2}",
                        pageHeight, pageWidth, printWidth));
                    Logger.LogInfoMsg(String.Format("Final overall margins are {0}, {1}, {2}, {3}",
                        PrintMargins.Left, PrintMargins.Right, PrintMargins.Top, PrintMargins.Bottom));
                    Logger.LogInfoMsg(String.Format("Table Alignment is {0}", TableAlignment.ToString()));
                }
            }

            // กำหนดหน้า / แถวที่จะพิมพ์

            // บันทึกช่วงการพิมพ์
            printRange = printDoc.PrinterSettings.PrintRange;
            if (EnableLogging) Logger.LogInfoMsg(String.Format("PrintRange is {0}", printRange));

            // หน้าที่จะพิมพ์จัดการตัวเลือก "SomePages" (บางหน้า)
            if (PrintRange.SomePages == printRange)
            {
                // set limits to only print some pages กำหนดขีด จำกัด เพื่อพิมพ์บางหน้าเท่านั้น
                fromPage = printDoc.PrinterSettings.FromPage;
                toPage = printDoc.PrinterSettings.ToPage;
            }
            else
            {
                // set extremes so that we'll print all pages ตั้งค่าสุดขั้วเพื่อที่เราจะพิมพ์หน้าทั้งหมด
                fromPage = 0;
                toPage = maxPages;
            }

            // กำหนดสิ่งที่จะพิมพ์
            SetupPrintRange();

            // ตั้งค่า override ความกว้างและคอลัมน์ คงที่
            SetupColumns();

            // เมื่อเรารู้ว่า เรากำลังพิมพ์อะไรให้วัดพื้นที่การพิมพ์ + นับหน้า

            // วัดพื้นที่การพิมพ์ (Measure)
            measureprintarea(printDoc.PrinterSettings.CreateMeasurementGraphics());

            // นับหน้า
            totalpages = Pagination();

        }

        /// <summary>
        /// ตั้งค่า override ความกว้างและรายการคอลัมน์คงที่ (FIX) ฟิตเปรี๊ยะ!!
        /// </summary>
        private void SetupColumns()
        {
            // ระบุคอลัมน์ที่คงที่ ด้วยหมายเลขคอลัมน์ในรายการที่พิมพ์
            foreach (string colname in fixedcolumnnames)
            {
                try
                {
                    fixedcolumns.Add(GetColumnIndex(colname));
                }
                catch // (Exception ex)
                {
                    // คอลัมน์ที่ขาดหายไป ให้เพิ่มลงในรายการพิมพ์แล้วลองอีกครั้ง
                    colstoprint.Add(dgv.Columns[colname]);
                    fixedcolumns.Add(GetColumnIndex(colname));
                }
            }

            // ปรับรายการ override เพื่อให้มีจำนวนรายการเท่ากับ colstoprint
            foreach (DataGridViewColumn col in colstoprint)
                if (publicwidthoverrides.ContainsKey(col.Name))
                    colwidthsoverride.Add(publicwidthoverrides[col.Name]);
                else
                    colwidthsoverride.Add(-1);

        }

        /// <summary>
        /// Determine the print range based on dialog selections and user input. The rows
        /// กำหนดช่วงการพิมพ์ตามการเลือกโต้ตอบและอินพุตของผู้ใช้ แถว
        /// and columns are sorted to ensure that the rows appear in their correct index 
        /// และคอลัมน์จะถูกจัดเรียงเพื่อให้แน่ใจว่าแถวปรากฏในดัชนีที่ถูกต้อง
        /// order and the columns appear in DisplayIndex order to account for added columns
        /// คำสั่งซื้อและคอลัมน์จะปรากฏในลำดับการแสดงผลเพื่อทำรายการคอลัมน์เพิ่ม
        /// and re-ordered columns. และสั่งซื้อคอลัมน์ใหม่
        /// </summary>
        private void SetupPrintRange()
        {
            //-----------------------------------------------------------------
            // set up the rows and columns to print ตั้งค่าแถวและคอลัมน์ที่จะพิมพ์
            //
            // Note: The "Selectedxxxx" lists in the datagridview are 'stacks' that
            // หมายเหตุ: รายการ "Selectedxxxx" ใน DataGridview คือ 'สแต็ค'
            //  have the selected items pushed in the *in the order they were selected*
            // ผลักรายการที่เลือกไว้ใน * ตามลำดับที่เลือก *
            //  i.e. not the order you want to print them in! เช่นไม่ใช่คำสั่งซื้อที่คุณต้องการพิมพ์!
            //-----------------------------------------------------------------
            SortedList temprowstoprint = null;
            SortedList tempcolstoprint = null;

            // rows to print (handles "selection" and "current page" options
            // แถวที่จะพิมพ์ (จัดการกับตัวเลือก "การเลือก" และ "หน้าปัจจุบัน"
            if (PrintRange.Selection == printRange)
            {
                temprowstoprint = new SortedList(dgv.SelectedCells.Count);
                tempcolstoprint = new SortedList(dgv.SelectedCells.Count);

                //if DGV has rows selected, it's easy, selected rows and all visible columns
                // ถ้า DGV มีการเลือกแถวมันง่ายเลือกแถวและคอลัมน์ที่มองเห็นได้ทั้งหมด
                if (0 != dgv.SelectedRows.Count)
                {
                    temprowstoprint = new SortedList(dgv.SelectedRows.Count);
                    tempcolstoprint = new SortedList(dgv.Columns.Count);

                    // sort the rows into index order จัดเรียงแถวตามลำดับดัชนี
                    temprowstoprint = new SortedList(dgv.SelectedRows.Count);
                    foreach (DataGridViewRow row in dgv.SelectedRows)
                        if (row.Visible && !row.IsNewRow)
                            temprowstoprint.Add(row.Index, row);

                    // sort the columns into display order จัดเรียงคอลัมน์ตามลำดับการแสดงผล
                    foreach (DataGridViewColumn col in dgv.Columns) if (col.Visible) tempcolstoprint.Add(col.DisplayIndex, col);
                }
                // if selected columns, then all rows, and selected columns หากคอลัมน์ที่เลือกไว้จะเป็นทุกแถวและคอลัมน์ที่เลือก
                else if (0 != dgv.SelectedColumns.Count)
                {
                    temprowstoprint = new SortedList(dgv.Rows.Count);
                    tempcolstoprint = new SortedList(dgv.SelectedColumns.Count);

                    foreach (DataGridViewRow row in dgv.Rows)
                        if (row.Visible && !row.IsNewRow)
                            temprowstoprint.Add(row.Index, row);

                    foreach (DataGridViewColumn col in dgv.SelectedColumns)
                        if (col.Visible)
                            tempcolstoprint.Add(col.DisplayIndex, col);
                }
                // we just have a bunch of selected cells so we have to do some work เรามีเซลล์ที่คัดสรรแล้วเราต้องทำงานบางอย่าง
                else
                {
                    // set up sorted lists. the selectedcells method does not guarantee
                    // ตั้งค่ารายการที่เรียงลำดับ เมธอด selectcells ไม่รับประกัน
                    // that the cells will always be in left-right top-bottom order. 
                    // ว่าเซลล์จะอยู่ในลำดับซ้ายบนขวาล่างเสมอ
                    temprowstoprint = new SortedList(dgv.SelectedCells.Count);
                    tempcolstoprint = new SortedList(dgv.SelectedCells.Count);

                    // for each selected cell, add unique rows and columns
                    // สำหรับแต่ละเซลล์ที่เลือกให้เพิ่มแถวและคอลัมน์ที่ไม่ซ้ำกัน
                    int displayindex, colindex, rowindex;
                    foreach (DataGridViewCell cell in dgv.SelectedCells)
                    {
                        displayindex = cell.OwningColumn.DisplayIndex;
                        colindex = cell.ColumnIndex;
                        rowindex = cell.RowIndex;

                        // add unique rows เพิ่มแถวที่ไม่ซ้ำกัน
                        if (!temprowstoprint.Contains(rowindex))
                        {
                            DataGridViewRow row = dgv.Rows[rowindex];
                            if (row.Visible && !row.IsNewRow)
                                temprowstoprint.Add(rowindex, dgv.Rows[rowindex]);
                        }
                        // add unique columns เพิ่มคอลัมน์ที่ไม่ซ้ำกัน
                        if (!tempcolstoprint.Contains(displayindex))
                            tempcolstoprint.Add(displayindex, dgv.Columns[colindex]);
                    }
                }
            }
            // if current page was selected, print visible columns for the
            // หากเลือกหน้าปัจจุบันให้พิมพ์คอลัมน์ที่มองเห็นได้สำหรับ
            // displayed rows  แถวที่แสดง              
            else if (PrintRange.CurrentPage == printRange)
            {
                // create lists สร้างรายการ
                temprowstoprint = new SortedList(dgv.DisplayedRowCount(true));
                tempcolstoprint = new SortedList(dgv.Columns.Count);

                // select all visible rows on displayed page เลือกแถวที่มองเห็นได้ทั้งหมดในหน้าแสดง
                for (int i = dgv.FirstDisplayedScrollingRowIndex;
                    i < dgv.FirstDisplayedScrollingRowIndex + dgv.DisplayedRowCount(true);
                    i++)
                {
                    DataGridViewRow row = dgv.Rows[i];
                    if (row.Visible) temprowstoprint.Add(row.Index, row);
                }

                // select all visible columns เลือกคอลัมน์ที่มองเห็นได้ทั้งหมด
                foreach (DataGridViewColumn col in dgv.Columns) if (col.Visible) tempcolstoprint.Add(col.DisplayIndex, col);
            }
            // this is the default for print all - everything marked visible will be printed
            // นี่คือค่าเริ่มต้นสำหรับการพิมพ์ทั้งหมด - ทุกอย่างที่ทำเครื่องหมายมองเห็นได้จะถูกพิมพ์
            // this is also used when printing specific pages or page ranges as we won't know
            // สิ่งนี้ยังใช้เมื่อพิมพ์หน้าหรือช่วงหน้าเฉพาะตามที่เราไม่รู้
            // what to print until we size all the rows สิ่งที่จะพิมพ์จนกว่าเราจะกำหนดขนาดของแถวทั้งหมด
            else
            {
                temprowstoprint = new SortedList(dgv.Rows.Count);
                tempcolstoprint = new SortedList(dgv.Columns.Count);

                // select all visible rows and all visible columns - but don't include the new 'data entry row' 
                // เลือกแถวที่มองเห็นได้ทั้งหมดและคอลัมน์ที่มองเห็นได้ทั้งหมด แต่ไม่รวม 'แถวป้อนข้อมูล' ใหม่
                foreach (DataGridViewRow row in dgv.Rows) if (row.Visible && !row.IsNewRow) temprowstoprint.Add(row.Index, row);

                // sort the columns into display order จัดเรียงคอลัมน์ตามลำดับการแสดงผล
                foreach (DataGridViewColumn col in dgv.Columns) if (col.Visible) tempcolstoprint.Add(col.DisplayIndex, col);
            }

            // move rows and columns into global containers ย้ายแถวและคอลัมน์ไปยังคอนเทนเนอร์ทั่วโลก
            rowstoprint = new List<rowdata>(temprowstoprint.Count);
            foreach (object item in temprowstoprint.Values) rowstoprint.Add(new rowdata() { row = (DataGridViewRow)item });

            colstoprint = new List<DataGridViewColumn>(tempcolstoprint.Count);
            foreach (object item in tempcolstoprint.Values) colstoprint.Add(item);

            // remove "hidden" columns from list of columns to print ลบคอลัมน์ "ซ่อน" ออกจากรายการคอลัมน์ที่จะพิมพ์
            foreach (String columnname in HideColumns)
            {
                colstoprint.Remove(dgv.Columns[columnname]);
            }

            if (EnableLogging) Logger.LogInfoMsg(String.Format("Grid Printout Range is {0} columns", colstoprint.Count));
            if (EnableLogging) Logger.LogInfoMsg(String.Format("Grid Printout Range is {0} rows", rowstoprint.Count));

        }

        /// <summary>
        /// Centralize the string format settings. Build a string format object
        /// รวมศูนย์การตั้งค่ารูปแบบสตริง สร้างวัตถุรูปแบบสตริง
        /// using passed in settings, (allowing a user override of a single setting)
        /// ใช้ผ่านการตั้งค่า (อนุญาตให้ผู้ใช้แทนที่การตั้งค่าเดียว)
        /// and get the alignment from the cell control style.
        /// และรับการจัดตำแหน่งจากสไตล์การควบคุมเซลล์
        /// </summary>
        /// <param name="format">String format, ref parameter with return settings</param>
        /// <param name="controlstyle">DataGridView style to apply (if available)</param>
        /// <param name="alignment">Override text Alignment</param>
        /// <param name="linealignment">Override line alignment</param>
        /// <param name="flags">String format flags</param>
        /// <param name="trim">Override string trimming flags</param>
        /// <returns></returns>
        private void buildstringformat(ref StringFormat format, DataGridViewCellStyle controlstyle,
            StringAlignment alignment, StringAlignment linealignment, StringFormatFlags flags,
            StringTrimming trim)
        {
            // allocate format if it doesn't already exist จัดสรรรูปแบบหากยังไม่มีอยู่
            if (null == format)
                format = new StringFormat();

            // Set defaults
            format.Alignment = alignment;
            format.LineAlignment = linealignment;
            format.FormatFlags = flags;
            format.Trimming = trim;

            // Check on right-to-left flag. This is set at the grid level, but doesn't show up 
            // ตรวจสอบการตั้งค่าสถานะจากขวาไปซ้าย สิ่งนี้ถูกตั้งค่าที่ระดับกริด แต่ไม่ปรากฏขึ้น
            // as a cell format. Urgh. เป็นรูปแบบของเซลล์ Urgh
            if ((null != dgv) && (RightToLeft.Yes == dgv.RightToLeft))
                format.FormatFlags |= StringFormatFlags.DirectionRightToLeft;

            // use cell alignment to override defaulted alignments ใช้การจัดตำแหน่งเซลล์เพื่อแทนที่การจัดแนวเริ่มต้น
            if (null != controlstyle)
            {
                // Adjust the format based on the control settings, bias towards centered
                // ปรับรูปแบบตามการตั้งค่าการควบคุมอคติต่อกึ่งกลาง
                DataGridViewContentAlignment cellalign = controlstyle.Alignment;
                if (cellalign.ToString().Contains("Center")) format.Alignment = StringAlignment.Center;
                else if (cellalign.ToString().Contains("Left")) format.Alignment = StringAlignment.Near;
                else if (cellalign.ToString().Contains("Right")) format.Alignment = StringAlignment.Far;

                if (cellalign.ToString().Contains("Top")) format.LineAlignment = StringAlignment.Near;
                else if (cellalign.ToString().Contains("Middle")) format.LineAlignment = StringAlignment.Center;
                else if (cellalign.ToString().Contains("Bottom")) format.LineAlignment = StringAlignment.Far;
            }
        }

        /// <summary>
        /// Calculate cell size based on data versus size settings
        /// คำนวณขนาดของเซลล์ตามข้อมูลเปรียบเทียบกับการตั้งค่าขนาด
        /// </summary>
        /// <param name="g">Current graphics context</param> บริบทกราฟิกปัจจุบัน
        /// <param name="cell">Cell being measured</param> เซลล์ถูกวัด
        /// <param name="index">Column index of cell being measured</param> ดัชนีคอลัมน์ของเซลล์ที่ถูกวัด
        /// <param name="cellstyle">Computed Style of cell being measured</param> รูปแบบการคำนวณของเซลล์ที่ถูกวัด
        /// <param name="basewidth">Initial width for size calculation</param> ความกว้างเริ่มต้นสำหรับการคำนวณขนาด
        /// <param name="format">Computed string format for cell data</param> รูปแบบสตริงที่คำนวณได้สำหรับข้อมูลเซลล์
        /// <returns>Size of printed cell</returns> ขนาดของเซลล์ที่พิมพ์
        private SizeF calccellsize(Graphics g, DataGridViewCell cell, DataGridViewCellStyle cellstyle,
            float basewidth, float overridewidth, StringFormat format)
        {
            // Start with the grid view cell size เริ่มต้นด้วยขนาดเซลล์มุมมองกริด
            SizeF size = new SizeF(cell.Size);

            // If we need to do any calculated cell sizes, we need to measure the cell contents
            // หากเราต้องการทำการคำนวณขนาดเซลล์ใด ๆ เราจำเป็นต้องวัดเนื้อหาของเซลล์
            if ((RowHeightSetting.DataHeight == RowHeight) ||
                (ColumnWidthSetting.DataWidth == ColumnWidth) ||
                (ColumnWidthSetting.Porportional == ColumnWidth))
            {
                SizeF datasize;

                //-------------------------------------------------------------
                // Measure cell contents วัดเนื้อหาของเซลล์
                //-------------------------------------------------------------
                if (("DataGridViewImageCell" == dgv.Columns[cell.ColumnIndex].CellType.Name)
                    && ("Image" == cell.ValueType.Name || "Byte[]" == cell.ValueType.Name))
                {
                    // image to measure ภาพที่จะวัด
                    Image img;

                    // if we don't actually have a value, then just exit with a minimum size.
                    // หากเราไม่มีค่าจริง ๆ แล้วก็ออกด้วยขนาดที่เล็กที่สุด
                    if ((null == cell.Value) || (typeof(DBNull) == cell.Value.GetType()))
                        return new SizeF(1, 1);

                    // Check on type of image cell value - may not be an actual "image" type
                    // ตรวจสอบประเภทของค่าเซลล์ภาพ - อาจไม่ใช่ประเภท "ภาพ" จริง
                    if ("Image" == cell.ValueType.Name || "Object" == cell.ValueType.Name)
                    {
                        // if it's an "image" type, then load it directly
                        // หากเป็นประเภท "รูปภาพ" ให้โหลดโดยตรง
                        img = (System.Drawing.Image)cell.Value;
                    }
                    else if ("Byte[]" == cell.ValueType.Name)
                    {
                        // if it's not an "image" type (i.e. loaded from a database to a bound column)
                        // หากไม่ใช่ประเภท "รูปภาพ"(เช่นโหลดจากฐานข้อมูลไปยังคอลัมน์ที่ถูกผูกไว้)
                        // convert the underlying byte array to an image แปลงอาร์เรย์ไบต์พื้นฐานเป็นรูปภาพ
                        ImageConverter ic = new ImageConverter();
                        img = (Image)ic.ConvertFrom((byte[])cell.Value);
                    }
                    else
                        throw new Exception(String.Format("Unknown image cell underlying type: {0} in column {1}",
                            cell.ValueType.Name, cell.ColumnIndex));

                    // size to print is size of image ขนาดที่จะพิมพ์คือขนาดของภาพ
                    datasize = img.Size;
                }
                else
                {
                    float width = (-1 != overridewidth) ? overridewidth : basewidth;

                    // measure the data for each column, keep widths and biggest height
                    // วัดข้อมูลสำหรับแต่ละคอลัมน์รักษาความกว้างและความสูงที่ใหญ่ที่สุด
                    datasize = g.MeasureString(cell.EditedFormattedValue.ToString(), cellstyle.Font,
                        new SizeF(width, maxPages), format);

                    // if we have excessively large cell, limit it to one page width
                    // หากเรามีเซลล์ที่ใหญ่เกินไปให้ จำกัด ไว้ที่ความกว้างหนึ่งหน้า
                    if (printWidth < datasize.Width)
                        datasize = g.MeasureString(cell.FormattedValue.ToString(), cellstyle.Font,
                        new SizeF(pageWidth - cellstyle.Padding.Left - cellstyle.Padding.Right, maxPages),
                        format);
                }

                //-------------------------------------------------------------
                // Add in padding for data based cell sizes and porportional columns
                // เพิ่มการเติมสำหรับขนาดเซลล์ตามข้อมูลและคอลัมน์ตามสัดส่วน
                //-------------------------------------------------------------

                // set cell height to string height if indicated ตั้งค่าความสูงของเซลล์เป็นความสูงของสตริงถ้าระบุ
                if (RowHeightSetting.DataHeight == RowHeight)
                    size.Height = datasize.Height + cellstyle.Padding.Top + cellstyle.Padding.Bottom;

                // set cell width to calculated width if indicated ตั้งค่าความกว้างของเซลล์เป็นความกว้างที่คำนวณได้หากระบุไว้
                if ((ColumnWidthSetting.DataWidth == ColumnWidth) ||
                    (ColumnWidthSetting.Porportional == ColumnWidth))
                    size.Width = datasize.Width + cellstyle.Padding.Left + cellstyle.Padding.Right;
            }

            return size;
        }

        /// <summary>
        /// Recalculate row heights for cells whose width is greater than the set column width. 
        /// คำนวณแถวความสูงใหม่สำหรับเซลล์ที่มีความกว้างมากกว่าความกว้างคอลัมน์ที่ตั้งไว้
        /// Called when column widths are changed in order to flow text down the page instead of 
        /// เรียกว่าเมื่อมีการเปลี่ยนแปลงความกว้างของคอลัมน์เพื่อที่จะไหลข้อความลงหน้าแทน
        /// accross. ข้าม
        /// </summary>
        /// <param name="g">Graphics Context for measuring image columns</param> บริบทกราฟิกสำหรับการวัดคอลัมน์ภาพ
        /// <param name="colindex">column index in colstoprint</param> ดัชนีคอลัมน์ใน colstoprint
        /// <param name="newcolwidth">new column width</param> ความกว้างคอลัมน์ใหม่
        private void RecalcRowHeights(Graphics g, int colindex, float newcolwidth)
        {
            DataGridViewCell cell = null;
            float finalsize = 0F;

            // search calculated cell sizes for widths larger than our new width
            // ค้นหาขนาดเซลล์ที่คำนวณได้สำหรับความกว้างที่ใหญ่กว่าความกว้างใหม่ของเรา
            for (int i = 0; i < rowstoprint.Count; i++)
            {
                cell = ((DataGridViewRow)rowstoprint[i].row).Cells[((DataGridViewColumn)colstoprint[colindex]).Index];

                if (RowHeightSetting.DataHeight == RowHeight)
                {
                    StringFormat currentformat = null;

                    // get column style รับสไตล์คอลัมน์
                    DataGridViewCellStyle colstyle = GetStyle(((DataGridViewRow)rowstoprint[i].row), ((DataGridViewColumn)colstoprint[colindex]));

                    // build the cell style and font สร้างลักษณะเซลล์และแบบอักษร
                    buildstringformat(ref currentformat, colstyle, cellformat.Alignment, cellformat.LineAlignment,
                        cellformat.FormatFlags, cellformat.Trimming);

                    // recalculate cell size using new width. This will flow data down the page and 
                    // คำนวณขนาดเซลล์ใหม่โดยใช้ความกว้างใหม่ จะเป็นการไหลข้อมูลลงสู่หน้ากระดาษและ
                    // change the row height เปลี่ยนความสูงของแถว
                    SizeF size = calccellsize(g, cell, colstyle, newcolwidth, colwidthsoverride[colindex], currentformat);

                    finalsize = size.Height;
                }
                else
                {
                    finalsize = cell.Size.Height;
                }

                // change the saved row height based on the recalculated size เปลี่ยนความสูงของแถวที่บันทึกไว้ตามขนาดการคำนวณใหม่
                rowstoprint[i].height = (rowstoprint[i].height < finalsize ? finalsize : rowstoprint[i].height);
            }
        }


        /// <summary>
        /// Scan all the rows and columns to be printed and calculate the 
        /// สแกนแถวและคอลัมน์ทั้งหมดที่จะพิมพ์และคำนวณ
        /// overall individual column width (based on largest column value), 
        /// ความกว้างแต่ละคอลัมน์โดยรวม (ขึ้นอยู่กับค่าคอลัมน์ที่ใหญ่ที่สุด)
        /// the header sizes, and determine all the row heights.
        /// ขนาดส่วนหัวและกำหนดความสูงของแถวทั้งหมด
        /// </summary>
        /// <param name="g">The graphics context for all measurements</param> บริบทกราฟิกสำหรับการวัดทั้งหมด
        private void measureprintarea(Graphics g)
        {
            int i, j;
            colwidths = new List<float>(colstoprint.Count);
            footerHeight = 0;

            // temp variables ตัวแปรอุณหภูมิ
            DataGridViewColumn col;
            DataGridViewRow row;

            //-----------------------------------------------------------------
            // measure the page headers and footers, including the grid column header cells
            // วัดส่วนหัวและส่วนท้ายของหน้ารวมถึงเซลล์ส่วนหัวคอลัมน์ตาราง
            //-----------------------------------------------------------------

            // set initial column sizes based on column titles กำหนดขนาดคอลัมน์เริ่มต้นตามชื่อคอลัมน์
            for (i = 0; i < colstoprint.Count; i++)
            {
                col = (DataGridViewColumn)colstoprint[i];

                //-------------------------------------------------------------
                // Build String format and Cell style สร้างรูปแบบสตริงและลักษณะเซลล์
                //-------------------------------------------------------------

                // get gridview style, and override if we have a set style for this column
                // รับสไตล์ gridview และแทนที่ถ้าเรามีสไตล์ที่กำหนดไว้สำหรับคอลัมน์นี้
                StringFormat currentformat = null;
                DataGridViewCellStyle headercolstyle = col.HeaderCell.InheritedStyle.Clone();
                if (ColumnHeaderStyles.ContainsKey(col.Name))
                {
                    headercolstyle = columnheaderstyles[col.Name];

                    // build the cell style and font  สร้างลักษณะเซลล์และแบบอักษร
                    buildstringformat(ref currentformat, headercolstyle, cellformat.Alignment, cellformat.LineAlignment,
                        cellformat.FormatFlags, cellformat.Trimming);
                }
                else if (col.HasDefaultCellStyle)
                {
                    // build the cell style and font  สร้างลักษณะเซลล์และแบบอักษร
                    buildstringformat(ref currentformat, headercolstyle, cellformat.Alignment, cellformat.LineAlignment,
                        cellformat.FormatFlags, cellformat.Trimming);
                }
                else
                {
                    currentformat = columnheadercellformat;
                }

                //-------------------------------------------------------------
                // Calculate and accumulate column header width and height
                // คำนวณและสะสมความกว้างและความสูงของส่วนหัวคอลัมน์
                //-------------------------------------------------------------
                SizeF size = col.HeaderCell.Size;

                // deal with overridden col widths จัดการกับความกว้างของคอลัมน์ที่ถูกแทนที่
                float usewidth = 0;
                if (0 <= colwidthsoverride[i])
                    //usewidth = colwidthsoverride[i];
                    colwidths.Add(colwidthsoverride[i]);            // override means set that size
                else if ((ColumnWidthSetting.CellWidth == ColumnWidth) || (ColumnWidthSetting.Porportional == ColumnWidth))
                {
                    usewidth = col.HeaderCell.Size.Width;
                    // calculate the size of column header cells คำนวณขนาดของเซลล์ส่วนหัวคอลัมน์
                    size = calccellsize(g, col.HeaderCell, headercolstyle, usewidth, colwidthsoverride[i], columnheadercellformat);
                    colwidths.Add(col.Width);                       // otherwise use the data width มิฉะนั้นใช้ความกว้างของข้อมูล
                } 
                else
                {
                    usewidth = printWidth;
                    // calculate the size of column header cells คำนวณขนาดของเซลล์ส่วนหัวคอลัมน์
                    size = calccellsize(g, col.HeaderCell, headercolstyle, usewidth, colwidthsoverride[i], columnheadercellformat);
                    colwidths.Add(size.Width);
                }

                // accumulate heights, saving largest for data sized option
                // สะสมความสูงประหยัดที่สุดสำหรับตัวเลือกขนาดข้อมูล
                if (RowHeightSetting.DataHeight == RowHeight)
                    colheaderheight = (colheaderheight < size.Height ? size.Height : colheaderheight);
                else
                    colheaderheight = col.HeaderCell.Size.Height;
            }

            //-----------------------------------------------------------------
            // measure the page number วัดหมายเลขหน้า
            //-----------------------------------------------------------------

            if (pageno)
            {
                pagenumberHeight = (g.MeasureString("Page", pagenofont, printWidth, pagenumberformat)).Height;
            }

            //-----------------------------------------------------------------
            // Calc height of header. ความสูงของส่วนหัวของ Calc
            // Header height is height of page number, title, subtitle and height of column headers
            // ความสูงส่วนหัวคือความสูงของหมายเลขหน้าชื่อเรื่องคำบรรยายและความสูงของส่วนหัวคอลัมน์
            //-----------------------------------------------------------------
            if (PrintHeader)
            {
                // calculate title and subtitle heights คำนวณชื่อและความสูงของคำบรรยาย
                titleheight = (g.MeasureString(title, titlefont, printWidth, titleformat)).Height;
                subtitleheight = (g.MeasureString(subtitle, subtitlefont, printWidth, subtitleformat)).Height;
            }

            //-----------------------------------------------------------------
            // measure the footer, if one is provided. Include the page number if we're printing
            // วัดส่วนท้ายถ้ามีให้ รวมหมายเลขหน้าหากเรากำลังพิมพ์
            // it on the bottom มันอยู่ด้านล่าง
            //-----------------------------------------------------------------
            if (PrintFooter)
            {
                if (!String.IsNullOrEmpty(footer))
                {
                    footerHeight += (g.MeasureString(footer, footerfont, printWidth, footerformat)).Height;
                }

                footerHeight += footerspacing;
            }

            //-----------------------------------------------------------------
            // Calculate column widths, adjusting for porportional columns
            // คำนวณความกว้างของคอลัมน์ปรับสำหรับคอลัมน์ตามสัดส่วน
            // and datawidth columns. Row heights are calculated later
            // และคอลัมน์วันที่ คำนวณความสูงของแถวในภายหลัง
            //-----------------------------------------------------------------
            for (i = 0; i < rowstoprint.Count; i++)
            {
                row = (DataGridViewRow)rowstoprint[i].row;

                // add row headers if they're visible เพิ่มส่วนหัวของแถวหากมองเห็นได้
                if ((bool)PrintRowHeaders)
                {
                    // provide a default 'blank' value to prevent a 0 length if we're supposed to show
                    // ระบุค่าเริ่มต้น 'blank' เพื่อป้องกันความยาว 0 ถ้าเราควรจะแสดง
                    // row headers ส่วนหัวของแถว
                    String rowheadertext = String.IsNullOrEmpty(row.HeaderCell.FormattedValue.ToString())
                        ? rowheadercelldefaulttext : row.HeaderCell.FormattedValue.ToString();

                    SizeF rhsize = g.MeasureString(rowheadertext,
                        row.HeaderCell.InheritedStyle.Font);
                    rowheaderwidth = (rowheaderwidth < rhsize.Width) ? rhsize.Width : rowheaderwidth;
                }

                // calculate widths for each column. We're looking for the largest width needed for
                // คำนวณความกว้างสำหรับแต่ละคอลัมน์ เรากำลังมองหาความกว้างที่ใหญ่ที่สุดที่จำเป็นสำหรับ
                // all the rows of data. แถวข้อมูลทั้งหมด
                for (j = 0; j < colstoprint.Count; j++)
                {
                    col = (DataGridViewColumn)colstoprint[j];

                    //-------------------------------------------------------------
                    // Build string format and cell style  สร้างรูปแบบสตริงและลักษณะเซลล์
                    //-------------------------------------------------------------

                    // get gridview style, and override if we have a set style for this column
                    // รับสไตล์ gridview และแทนที่ถ้าเรามีสไตล์ที่กำหนดไว้สำหรับคอลัมน์นี้
                    StringFormat currentformat = null;
                    DataGridViewCellStyle colstyle = GetStyle(row, col); // = row.Cells[col.Index].InheritedStyle.Clone();

                    // build the cell style and font สร้างลักษณะเซลล์และแบบอักษร
                    buildstringformat(ref currentformat, colstyle, cellformat.Alignment, cellformat.LineAlignment,
                        cellformat.FormatFlags, cellformat.Trimming);

                    //-------------------------------------------------------------
                    // Calculate and accumulate cell widths and heights คำนวณและสะสมความกว้างและความสูงของเซลล์
                    //-------------------------------------------------------------
                    float basewidth;

                    // get the default width, depending on overrides. Only calculate data
                    // รับความกว้างเริ่มต้นขึ้นอยู่กับการแทนที่ คำนวณข้อมูลเท่านั้น
                    // sizes for DataWidth column setting. ขนาดสำหรับการตั้งค่าคอลัมน์ DataWidth
                    if (0 <= colwidthsoverride[j])
                        // set overridden column width กำหนดความกว้างคอลัมน์แทนที่
                        basewidth = colwidthsoverride[j];
                    else if ((ColumnWidthSetting.CellWidth == ColumnWidth) || (ColumnWidthSetting.Porportional == ColumnWidth))
                        // set default to same as title cell width ตั้งค่าเริ่มต้นเป็นเช่นเดียวกับความกว้างของเซลล์ชื่อเรื่อง
                        basewidth = colwidths[j];
                    else
                    {
                        // limit to one page จำกัด หนึ่งหน้า
                        basewidth = printWidth;

                        // remove padding ลบการขยาย
                        basewidth -= colstyle.Padding.Left + colstyle.Padding.Right;

                        // calc cell size ขนาดเซลล์แคลเซียม
                        SizeF size = calccellsize(g, row.Cells[col.Index], colstyle,
                            basewidth, colwidthsoverride[j], currentformat);

                        basewidth = size.Width;
                    }

                    // if width is not overridden and we're using data width then accumulate column widths
                    // หากความกว้างไม่ถูกแทนที่และเราใช้ความกว้างของข้อมูลก็จะสะสมความกว้างของคอลัมน์
                    if (!(0 <= colwidthsoverride[j]) && (ColumnWidthSetting.DataWidth == ColumnWidth))
                        colwidths[j] = colwidths[j] < basewidth ? basewidth : colwidths[j];
                }
            }

            //-----------------------------------------------------------------
            // Break the columns accross page sets. This is the key to printing แบ่งคอลัมน์ออกเป็นชุดหน้า นี่คือกุญแจสำคัญในการพิมพ์
            // where the total width is wider than one page. โดยที่ความกว้างทั้งหมดกว้างกว่าหนึ่งหน้า
            //-----------------------------------------------------------------

            // assume everything will fit on one page สมมติว่าทุกอย่างจะพอดีกับหนึ่งหน้า
            pagesets = new List<PageDef>();
            pagesets.Add(new PageDef(PrintMargins, colstoprint.Count, pageWidth));
            int pset = 0;

            // Account for row headers บัญชีสำหรับส่วนหัวของแถว
            pagesets[pset].coltotalwidth = rowheaderwidth;

            // account for 'fixed' columns - these appear on every pageset บัญชีสำหรับคอลัมน์ 'คงที่' - สิ่งเหล่านี้จะปรากฏในทุกหน้าเว็บ
            for (j = 0; j < fixedcolumns.Count; j++)
            {
                int fixedcol = fixedcolumns[j];
                pagesets[pset].columnindex.Add(fixedcol);
                pagesets[pset].colstoprint.Add(colstoprint[fixedcol]);
                pagesets[pset].colwidths.Add(colwidths[fixedcol]);
                pagesets[pset].colwidthsoverride.Add(colwidthsoverride[fixedcol]);
                pagesets[pset].coltotalwidth += (colwidthsoverride[fixedcol] >= 0)
                    ? colwidthsoverride[fixedcol] : colwidths[fixedcol];
            }

            // check on fixed columns ตรวจสอบคอลัมน์คงที่
            if (printWidth < (pagesets[pset].coltotalwidth))
            {
                throw new Exception("Fixed column widths exceed the page width.");
            }

            // split remaining columns into page sets แยกคอลัมน์ที่เหลือออกเป็นชุดหน้า
            float columnwidth;
            for (i = 0; i < colstoprint.Count; i++)
            {
                // skip 'fixed' columns since we've already accounted for them ข้ามคอลัมน์ 'คงที่' เนื่องจากเราคิดไปแล้ว
                if (fixedcolumns.Contains(i))
                    continue;

                // get initial column width รับความกว้างคอลัมน์เริ่มต้น
                columnwidth = (colwidthsoverride[i] >= 0)
                    ? colwidthsoverride[i] : colwidths[i];

                // See if the column width takes us off the page - Except for the 
                // ดูว่าความกว้างของคอลัมน์พาเราออกหน้า - ยกเว้นสำหรับ
                // first column. This will prevent printing an empty page!! Otherwise,
                // คอลัมน์แรก สิ่งนี้จะป้องกันการพิมพ์หน้าว่าง !! มิฉะนั้น,
                // columns longer than the page width are printed on their own page
                // คอลัมน์ที่ยาวกว่าความกว้างของหน้าจะถูกพิมพ์ลงบนหน้าของตัวเอง
                if (printWidth < (pagesets[pset].coltotalwidth + columnwidth) && i != 0)
                {
                    pagesets.Add(new PageDef(PrintMargins, colstoprint.Count, pageWidth));
                    pset++;

                    // Account for row headers บัญชีสำหรับส่วนหัวของแถว
                    pagesets[pset].coltotalwidth = rowheaderwidth;

                    // account for 'fixed' columns - these appear on every pageset
                    // บัญชีสำหรับคอลัมน์ 'คงที่' - สิ่งเหล่านี้จะปรากฏในทุกหน้าเว็บ
                    for (j = 0; j < fixedcolumns.Count; j++)
                    {
                        int fixedcol = fixedcolumns[j];
                        pagesets[pset].columnindex.Add(fixedcol);
                        pagesets[pset].colstoprint.Add(colstoprint[fixedcol]);
                        pagesets[pset].colwidths.Add(colwidths[fixedcol]);
                        pagesets[pset].colwidthsoverride.Add(colwidthsoverride[fixedcol]);
                        pagesets[pset].coltotalwidth += (colwidthsoverride[fixedcol] >= 0)
                            ? colwidthsoverride[fixedcol] : colwidths[fixedcol];
                    }

                    // check on fixed columns ตรวจสอบคอลัมน์คงที่
                    if (printWidth < (pagesets[pset].coltotalwidth))
                    {
                        throw new Exception("Fixed column widths exceed the page width.");
                    }
                }

                // update page set definition อัพเดตนิยามชุดเพจ
                pagesets[pset].columnindex.Add(i);
                pagesets[pset].colstoprint.Add(colstoprint[i]);
                pagesets[pset].colwidths.Add(colwidths[i]);
                pagesets[pset].colwidthsoverride.Add(colwidthsoverride[i]);
                pagesets[pset].coltotalwidth += columnwidth;
            }

            // for right to left language, reverse the column order for each page set
            // สำหรับภาษาจากขวาไปซ้ายกลับลำดับคอลัมน์สำหรับชุดหน้าแต่ละชุด
            if (RightToLeft.Yes == dgv.RightToLeft)
            {
                for (pset = 0; pset < pagesets.Count; pset++)
                {
                    pagesets[pset].columnindex.Reverse();
                    pagesets[pset].colstoprint.Reverse();
                    pagesets[pset].colwidths.Reverse();
                    pagesets[pset].colwidthsoverride.Reverse();
                }
            }

            for (i = 0; i < pagesets.Count; i++)
            {
                PageDef pageset = pagesets[i];
                if (EnableLogging)
                {
                    String columnlist = "";

                    Logger.LogInfoMsg(String.Format("PageSet {0} Information ----------------------------------------------", i));

                    // list out all the columns printed on this page since we may have fixed columns to account for
                    // ทำรายการคอลัมน์ทั้งหมดที่พิมพ์บนหน้านี้เนื่องจากเราอาจมีคอลัมน์คงที่ที่จะพิจารณา
                    for (int k = 0; k < pageset.colstoprint.Count; k++)
                        columnlist = String.Format("{0},{1}", columnlist,
                            ((DataGridViewColumn)(pageset.colstoprint[k])).Index);
                    Logger.LogInfoMsg(String.Format("Measured columns {0}", columnlist.Substring(1)));
                    columnlist = "";

                    // list original column widths for this page แสดงรายการความกว้างคอลัมน์ดั้งเดิมสำหรับหน้านี้
                    for (int k = 0; k < pageset.colstoprint.Count; k++)
                        columnlist = String.Format("{0},{1}", columnlist, pageset.colwidths[k]);
                    Logger.LogInfoMsg(String.Format("Original Column Widths: {0}", columnlist.Substring(1)));
                    columnlist = "";

                    // list column width override values ค่าการแทนที่ความกว้างของคอลัมน์รายการ
                    for (int k = 0; k < pageset.colstoprint.Count; k++)
                        columnlist = String.Format("{0},{1}", columnlist, pageset.colwidthsoverride[k]);
                    Logger.LogInfoMsg(String.Format("Overridden Column Widths: {0}", columnlist.Substring(1)));
                    columnlist = "";
                }

                //-----------------------------------------------------------------
                // Adjust column widths and table margins for each page ปรับความกว้างคอลัมน์และระยะขอบตารางสำหรับแต่ละหน้า
                //-----------------------------------------------------------------
                AdjustPageSets(g, pageset);

                //-----------------------------------------------------------------
                // Log Pagesets เข้าสู่ระบบ Pagesets
                //-----------------------------------------------------------------
                if (EnableLogging)
                {
                    String columnlist = "";

                    // list final column widths for this page รายการความกว้างคอลัมน์สุดท้ายสำหรับหน้านี้
                    for (int k = 0; k < pageset.colstoprint.Count; k++)
                        columnlist = String.Format("{0},{1}", columnlist, pageset.colwidths[k]);
                    Logger.LogInfoMsg(String.Format("Final Column Widths: {0}", columnlist.Substring(1)));
                    columnlist = "";

                    Logger.LogInfoMsg(String.Format("pageset print width is {0}, total column width to be printed is {1}",
                        pageset.printWidth, pageset.coltotalwidth));
                }
            }
        }

        /// <summary>
        /// Adjust column widths for fixed and porportional columns, set the 
        /// ปรับความกว้างของคอลัมน์สำหรับคอลัมน์คงที่และสัดส่วนตั้งค่า
        /// margins to enforce the selected tablealignment.
        /// ระยะขอบเพื่อบังคับใช้การจัดตำแหน่งแบบตารางที่เลือก
        /// </summary>
        /// <param name="g">The graphics context for all measurements</param> บริบทกราฟิกสำหรับการวัดทั้งหมด
        /// <param name="pageset">The pageset to adjust</param> ชุดหน้าเพื่อปรับ
        private void AdjustPageSets(Graphics g, PageDef pageset)
        {
            int i;
            float fixedcolwidth = rowheaderwidth;
            float remainingcolwidth = 0;
            float ratio;

            //-----------------------------------------------------------------
            // Adjust the column widths in the page set to their final values, ปรับความกว้างคอลัมน์ในชุดหน้าเว็บเป็นค่าสุดท้าย
            // accounting for overridden widths and porportional column stretching การบัญชีสำหรับความกว้างที่ถูกเขียนทับและการยืดคอลัมน์สัดส่วน
            //-----------------------------------------------------------------

            // calculate the amount of space reserved for fixed width columns คำนวณจำนวนเนื้อที่ที่สงวนไว้สำหรับคอลัมน์ความกว้างคงที่
            for (i = 0; i < pageset.colwidthsoverride.Count; i++)
                if (pageset.colwidthsoverride[i] >= 0)
                    fixedcolwidth += pageset.colwidthsoverride[i];

            // calculate the amount space requested for non-overridden columns คำนวณพื้นที่จำนวนที่ร้องขอสำหรับคอลัมน์ที่ไม่ถูกแทนที่
            for (i = 0; i < pageset.colwidths.Count; i++)
                if (pageset.colwidthsoverride[i] < 0)
                    remainingcolwidth += pageset.colwidths[i];

            // calculate the ratio for porportional columns, use 1 for คำนวณอัตราส่วนสำหรับคอลัมน์ที่มีสัดส่วนใช้ 1 สำหรับ
            // non-overridden columns or not porportional คอลัมน์ที่ไม่ถูกแทนที่หรือไม่เป็นสัดส่วน
            if ((porportionalcolumns || ColumnWidthSetting.Porportional == ColumnWidth) &&
                0 < remainingcolwidth)
                ratio = ((float)printWidth - fixedcolwidth) / (float)remainingcolwidth;
            else
                ratio = (float)1.0;

            // reset all column widths for override and/or porportionality. coltotalwidth รีเซ็ตความกว้างคอลัมน์ทั้งหมดสำหรับการแทนที่และ / หรือ porportionality coltotalwidth
            // for each pageset should be <= pageWidth  สำหรับแต่ละชุดของหน้าควรเป็น <= pageWidth
            pageset.coltotalwidth = rowheaderwidth;
            for (i = 0; i < pageset.colwidths.Count; i++)
            {
                if (pageset.colwidthsoverride[i] >= 0)
                    // use set width ใช้ความกว้างชุด
                    pageset.colwidths[i] = pageset.colwidthsoverride[i];
                else if (ColumnWidthSetting.Porportional == ColumnWidth)
                    // change the width by the ratio เปลี่ยนความกว้างตามอัตราส่วน
                    pageset.colwidths[i] = pageset.colwidths[i] * ratio;
                else if (pageset.colwidths[i] > printWidth - pageset.coltotalwidth)
                    pageset.colwidths[i] = printWidth - pageset.coltotalwidth;

                //recalculate any rows that need to flow down the page คำนวณแถวใด ๆ ที่จำเป็นต้องไหลลงหน้า
                RecalcRowHeights(g, pageset.columnindex[i], pageset.colwidths[i]);

                pageset.coltotalwidth += pageset.colwidths[i];

            }

            //-----------------------------------------------------------------
            // Table Alignment - now that we have the column widths established
            // การจัดตำแหน่งตาราง - ตอนนี้เราได้สร้างความกว้างของคอลัมน์แล้ว
            // we can reset the table margins to get left, right and centered
            // เราสามารถรีเซ็ตระยะขอบของตารางให้เป็นซ้ายขวาและกึ่งกลาง
            // for the table on the page สำหรับตารางในหน้า
            //-----------------------------------------------------------------

            // Reset Print Margins based on table alignment รีเซ็ตระยะขอบการพิมพ์ตามการจัดตาราง
            if (Alignment.Left == tablealignment)
            {
                // Bias table to the left by setting "right" value ตารางอคติทางซ้ายโดยการตั้งค่า "ขวา"
                pageset.margins.Right = pageWidth - pageset.margins.Left - (int)pageset.coltotalwidth;
                if (0 > pageset.margins.Right) pageset.margins.Right = 0;
            }
            else if (Alignment.Right == tablealignment)
            {
                // Bias table to the right by setting "left" value ตารางอคติทางด้านขวาโดยการตั้งค่า "ซ้าย"
                pageset.margins.Left = pageWidth - pageset.margins.Right - (int)pageset.coltotalwidth;
                if (0 > pageset.margins.Left) pageset.margins.Left = 0;
            }
            else if (Alignment.Center == tablealignment)
            {
                // Bias the table to the center by setting left and right equal ตั้งค่าตารางไปที่กึ่งกลางโดยตั้งค่าซ้ายและขวาเท่ากัน
                pageset.margins.Left = (pageWidth - (int)pageset.coltotalwidth) / 2;
                if (0 > pageset.margins.Left) pageset.margins.Left = 0;
                pageset.margins.Right = pageset.margins.Left;
            }
        }

        /// <summary>
        /// Set page breaks for the rows to be printed, and count total pages
        /// กำหนดตัวแบ่งหน้าสำหรับแถวที่จะพิมพ์และนับหน้าทั้งหมด
        /// </summary>
        private int Pagination()
        {
            float pos = 0;
            paging newpage = paging.keepgoing;

            //// if we're printing by pages, the total pages is the last page to 
            /// หากเรากำลังพิมพ์โดยหน้าหน้าทั้งหมดเป็นหน้าสุดท้าย
            //// print
            //if (toPage < maxPages)
            //    return toPage;

            // Start counting pages at 1 เริ่มนับหน้าที่ 1
            CurrentPage = 1;

            // Calculate where to stop printing the grid - count up from the bottom of the page.
            // คำนวณตำแหน่งที่จะหยุดการพิมพ์กริด - นับจากด้านล่างของหน้า
            staticheight = pageHeight - FooterHeight - pagesets[currentpageset].margins.Bottom; //PrintMargins.Bottom;

            // add in the page number height - doesn't matter at this point if it's printing on top or bottom
            // เพิ่มความสูงหมายเลขหน้า - ไม่สำคัญที่จุดนี้หากกำลังพิมพ์ที่ด้านบนหรือด้านล่าง
            staticheight -= PageNumberHeight;

            // Calculate where to start printing the grid for page 1 คำนวณตำแหน่งที่จะเริ่มพิมพ์กริดสำหรับหน้า 1
            pos = PrintMargins.Top + HeaderHeight;

            // set starting value for 'break on value change' column กำหนดค่าเริ่มต้นสำหรับคอลัมน์ 'หยุดการเปลี่ยนแปลงค่า'
            if (!String.IsNullOrEmpty(breakonvaluechange))
            {
                oldvalue = rowstoprint[0].row.Cells[breakonvaluechange].EditedFormattedValue;
            }

            // if we're printing by rows, sum up rowheights until we're done. หากเรากำลังพิมพ์ตามแถวให้สรุปความสูงของแถวจนกว่าเราจะพิมพ์เสร็จ
            for (int currentrow = 0; currentrow < (rowstoprint.Count); currentrow++)
            {
                // end of page: Count the page and reset to top of next page จุดสิ้นสุดของหน้า: นับหน้าและรีเซ็ตเป็นด้านบนของหน้าถัดไป
                if (pos + rowstoprint[currentrow].height >= staticheight)
                {
                    newpage = paging.outofroom;
                }

                // if we're breaking on value change in a column then watch that column ถ้าเราทำลายการเปลี่ยนแปลงค่าในคอลัมน์จากนั้นดูคอลัมน์นั้น
                if ((!String.IsNullOrEmpty(breakonvaluechange)) &&
                    (!oldvalue.Equals(rowstoprint[currentrow].row.Cells[breakonvaluechange].EditedFormattedValue)))
                {
                    newpage = paging.datachange;
                    oldvalue = rowstoprint[currentrow].row.Cells[breakonvaluechange].EditedFormattedValue;
                }

                // if we need to start a new page, count it and reset counters หากเราต้องการเริ่มหน้าใหม่ให้นับและรีเซ็ตตัวนับ
                if (newpage != paging.keepgoing)
                {
                    // note page break ตัวแบ่งหน้าโน้ต
                    rowstoprint[currentrow].pagebreak = true;

                    // count the page นับหน้า
                    CurrentPage++;

                    // if we're printing by pages, stop when we pass our limit หากเรากำลังพิมพ์โดยหน้าหยุดเมื่อเราผ่านขีด จำกัด ของเรา
                    if (CurrentPage > toPage) 
                    {
                        // we're done เราเสร็จแล้ว
                        return toPage;
                    }

                    // reset the counter - depending on setting รีเซ็ตตัวนับ - ขึ้นอยู่กับการตั้งค่า
                    if (KeepRowsTogether
                        || newpage == paging.datachange
                        || (newpage == paging.outofroom && (staticheight - pos) < KeepRowsTogetherTolerance))
                    {
                        // if we are keeping rows together and too little would be showing, put whole row on next page
                        // หากเรารวมแถวเข้าด้วยกันและแสดงน้อยเกินไปให้ใส่ทั้งแถวในหน้าถัดไป
                        pos = rowstoprint[currentrow].height;
                    }
                    else
                    {
                        // note page split แยกหน้าโน๊ต
                        rowstoprint[currentrow].splitrow = true;

                        // if we're not keeping rows together, only put remainder on next page หากเราไม่ได้เรียงแถวกันให้วางเศษเหลือไว้ในหน้าถัดไป
                        pos = pos + rowstoprint[currentrow].height - staticheight;
                    }

                    // Recalculate where to stop printing the grid because available space can change w/ dynamic header/footers.
                    // คำนวณตำแหน่งที่จะหยุดการพิมพ์กริดเนื่องจากพื้นที่ที่มีอยู่สามารถเปลี่ยนแปลงได้ด้วย / dynamic header / footers
                    staticheight = pageHeight - FooterHeight - pagesets[currentpageset].margins.Bottom; //PrintMargins.Bottom; พิมพ์ระยะขอบด้านล่าง;

                    // add in the page number height - doesn't matter at this point if it's printing on top or bottom
                    // เพิ่มความสูงหมายเลขหน้า - ไม่สำคัญที่จุดนี้หากกำลังพิมพ์ที่ด้านบนหรือด้านล่าง
                    staticheight += PageNumberHeight;


                    // พื้นที่คงที่ที่ด้านบนของหน้า
                    pos += PrintMargins.Top + HeaderHeight + PageNumberHeight;
                }
                else
                {
                    // add row space เพิ่มพื้นที่แถว
                    pos += rowstoprint[currentrow].height;
                }

                // reset สถานะใหม่
                newpage = paging.keepgoing;
            }

            // return หน้าที่นับ
            return CurrentPage;
        }

        /// <summary>
        /// ตรวจสอบหน้าเพิ่มเติม เรียกในตอนท้ายของการพิมพ์ชุดหน้า (ถ้ามีการตั้งค่าหน้าอื่นให้พิมพ์เราจะคืนค่าจริง)
        /// </summary>
        private bool DetermineHasMorePages()
        {
            currentpageset++;
            if (currentpageset < pagesets.Count)
            {
                //currentpageset--;   // ลดค่ากลับไปเป็นหมายเลขหน้าที่ถูกต้อง
                return true;        
            }
            else
                return false;
        }

        /// <summary>
        /// พิมพ์หนึ่งหน้า จะข้ามหน้าที่ไม่สามารถพิมพ์ได้ หากผู้ใช้เลือกตัวเลือก "บางหน้า"("some pages") ในกล่องโต้ตอบการพิมพ์ เรียกระหว่าง Print event
        /// </summary>
        /// <param name="g">Graphics ที่จะพิมพ์</param>
        private bool PrintPage(Graphics g)
        {
            // ตามที่เรากำหนด + การบันทึก
            int firstrow = 0;

            // ตั้งค่าสถานะสำหรับกระบวนการพิมพ์ต่อเนื่อง / สิ้นสุด
            bool HasMorePages = false;

            // การจัดการการพิมพ์หน้าทั้งหมด
            bool printthispage = false;

            // ตำแหน่งการพิมพ์ปัจจุบันภายในหน้าเดียว
            float printpos = pagesets[currentpageset].margins.Top;

            // เพิ่มจำนวนหน้าและตรวจสอบช่วงหน้า
            CurrentPage++;
            if (EnableLogging) Logger.LogInfoMsg(String.Format("Print Page processing page {0} -----------------------", CurrentPage));
            if ((CurrentPage >= fromPage) && (CurrentPage <= toPage))
                printthispage = true;

            // คำนวณพื้นที่แนวตั้งที่มีอยู่ - ที่เราหยุดพิมพ์แถว
            // หมายเหตุ: เว้นช่องว่างสำหรับหมายเลขหน้าหากอยู่ด้านล่าง
            staticheight = pageHeight - FooterHeight - pagesets[currentpageset].margins.Bottom;
            if (!pagenumberontop)
                staticheight -= PageNumberHeight;

            // นับพื้นที่ที่ใช้ในขณะที่เราทำงานจนหมดข้อมูล
            float used = 0;

            // ข้อมูลแถวปัจจุบัน
            rowdata thisrow = null;

            // ข้อมูลแถวถัดไป (lookahead)
            rowdata nextrow = null;

            // สแกนความสูงจนกว่าเราจะปิดหน้านี้ (ไม่ใช่การพิมพ์)

            while (!printthispage)
            {
                if (EnableLogging) Logger.LogInfoMsg(String.Format("Print Page skipping page {0} part {1}", CurrentPage, currentpageset + 1));

                // คำนวณและเพิ่มส่วนที่เราไม่ได้พิมพ์
                printpos = pagesets[currentpageset].margins.Top + HeaderHeight + PageNumberHeight;

                // เราทำกับหน้านี้หรือไม่?
                bool pagecomplete = false;
                currentrow = lastrowprinted + 1;

                // for logging สำหรับการเข้าสู่ระบบ
                firstrow = currentrow;

                do
                {
                    thisrow = rowstoprint[currentrow];

                    // เป็นพื้นที่ที่แถวนี้จะใช้ในหน้านี้
                    used = (thisrow.height - rowstartlocation) > (staticheight - printpos)
                            ? (staticheight - printpos) : thisrow.height - rowstartlocation;
                    printpos += used;

                    // ดูที่แถวถัดไปและเริ่มตรวจสอบว่าเราอยู่นอกจากนี้หรือไม่ & ต้องนับหน้า
                    lastrowprinted++;
                    currentrow++;
                    nextrow = (currentrow < rowstoprint.Count) ? rowstoprint[currentrow] : null;
                    if (null != nextrow && nextrow.pagebreak) // ตัวแบ่งหน้าก่อนแถวถัดไป
                    {
                        pagecomplete = true;

                        if (nextrow.splitrow)
                        {
                            // สำหรับแถวบางส่วนจะไปในหน้านี้ (แถวที่พิมพ์ไม่หมด)
                            rowstartlocation += (nextrow.height - rowstartlocation) > (staticheight - printpos)
                                ? (staticheight - printpos) : nextrow.height - rowstartlocation;
                        }
                    }
                    else
                    {
                        // เสร็จแถว ให้ reset startlocation และนับแถวนี้
                        rowstartlocation = 0;
                    }

                    // ถ้าเราไม่มีข้อมูล (ไม่มีแถว) 
                    if ((0 == rowstartlocation) && lastrowprinted >= rowstoprint.Count - 1)
                        pagecomplete = true;

                } while (!pagecomplete);

                // log rows ข้ามแล้ว
                if (EnableLogging) Logger.LogInfoMsg(String.Format("Print Page skipped rows {0} to {1}", firstrow, currentrow));

                // ข้ามไปที่หน้าถัดไป & ดูว่าอยู่ในช่วงที่พิมพ์หรือไม่
                CurrentPage++;

                if ((CurrentPage >= fromPage) && (CurrentPage <= toPage))
                    printthispage = true;

                if (0 != rowstartlocation)
                {
                    // เรายังไม่ได้ทำแถวนี้!!!
                    HasMorePages = true;
                }
                // เสร็จแล้ว!! ดูว่ามีจำนวนหน้าที่พิมพ์อีกหรือไม่
                else if ((lastrowprinted >= rowstoprint.Count - 1) || (CurrentPage > toPage))
                {
                    // reset ชุดหน้าถัดไป / แจ้งให้ผู้ใช้ทราบว่า เราดำเนินการเสร็จแล้ว!
                    HasMorePages = DetermineHasMorePages();

                    // reset ตัวนับ (ถ้าเราพิมพ์จากหน้าตัวอย่าง)
                    lastrowprinted = -1;
                    CurrentPage = 0;

                    return HasMorePages;
                }
            }

            if (EnableLogging)
            {
                Logger.LogInfoMsg(String.Format("Print Page printing page {0} part {1}", CurrentPage, currentpageset + 1));
                var m = pagesets[currentpageset].margins;
                Logger.LogInfoMsg(String.Format("Current Margins are {0}, {1}, {2}, {3}", m.Left, m.Right, m.Top, m.Bottom));
            }

            // พิมพ์ให้อยู่ในตำแหน่งคงที่

            // พิมพ์ "สมบูรณ์"("absolute") เพื่อให้ที่เราพิมพ์จะ 'อยู่ด้านบน'('on top')
            ImbeddedImageList.Where(p => p.ImageLocation == Location.Absolute).DrawImbeddedImage(g, pagesets[currentpageset].printWidth,
                pageHeight, pagesets[currentpageset].margins);

            // print headers ส่วนหัวพิมพ์

            // reset งานที่พิมพ์ เนื่องจากอาจมีการเปลี่ยนแปลง 'ข้ามหน้า'('skip pages') ด้านบน
            printpos = pagesets[currentpageset].margins.Top;

            // ข้ามส่วนหัว (headers) ถ้าการตั้งค่าสถานะเป็น false
            if (PrintHeader)
            {
                // พิมพ์ "ส่วนหัว"("header") เพื่อให้ที่เราพิมพ์จะ 'อยู่ด้านบน'('on top')
                ImbeddedImageList.Where(p => p.ImageLocation == Location.Header).DrawImbeddedImage(g, pagesets[currentpageset].printWidth,
                    pageHeight, pagesets[currentpageset].margins);

                // พิมพ์หมายเลขหน้า (ถ้าผู้ใช้เลือก)
                if (pagenumberontop)
                {
                    printpos = PrintPageNo(g, printpos);
                }

                // พิมพ์ title (ถ้ามี)
                if (0 != TitleHeight && !String.IsNullOrEmpty(title))
                    printsection(g, ref printpos, title, titlefont,
                        titlecolor, titleformat, overridetitleformat,
                        pagesets[currentpageset],
                        titlebackground, titleborder);

                // การเว้นวรรค title
                printpos += TitleHeight;

                // พิมพ์ subtitle (ถ้ามี)
                if (0 != SubTitleHeight && !String.IsNullOrEmpty(subtitle))
                    printsection(g, ref printpos, subtitle, subtitlefont,
                        subtitlecolor, subtitleformat, overridesubtitleformat,
                        pagesets[currentpageset],
                        subtitlebackground, subtitleborder);

                // การเว้นวรรค subtitle
                printpos += SubTitleHeight;
            }

            // พิมพ์ส่วนหัวคอลัมน์ (ยึดตามที่เราใส่ในตาราง)
            if ((bool)PrintColumnHeaders)
            {
                // พิมพ์ส่วนหัวของคอลัมน์
                printcolumnheaders(g, ref printpos, pagesets[currentpageset]);
            }

            // พิมพ์แถวจนกว่าจะพิมพ์หมด
            bool continueprinting = true;
            currentrow = lastrowprinted + 1;

            // for logging สำหรับการเข้าสู่ระบบ
            firstrow = currentrow;

            if (currentrow >= rowstoprint.Count)
            {
                // ระบุว่าเราพิมพ์เสร็จแล้ว!!!!
                continueprinting = false;
            }

            while (continueprinting)
            {
                thisrow = rowstoprint[currentrow];

                // พิมพ์แถว(row) ที่เราทำ + พื้นที่ที่ใช้
                used = printrow(g, printpos, (DataGridViewRow)(thisrow.row),
                    pagesets[currentpageset], rowstartlocation);
                printpos += used;

                // เริ่มตรวจสอบว่าจะพิมพ์แถวถัดไปหรือไม่ (ถ้าเรามีแถวถัดไป) มีต่อก็พิมพ์ ไม่ทิ้งเด้ออ =w=
                lastrowprinted++;
                currentrow++;
                nextrow = (currentrow < rowstoprint.Count) ? rowstoprint[currentrow] : null;
                if (null != nextrow && nextrow.pagebreak)
                {
                    continueprinting = false;

                    // พิมพ์แถวบางส่วน(แถวที่พิมพ์ไม่หมด) ก่อนที่จะ break
                    if (nextrow.splitrow)
                    {
                        // พิมพ์ที่เราได้ทำในหน้านี้ + พิมพ์ส่วนที่เหลือในหน้าถัดไป
                        rowstartlocation += printrow(g, printpos, (DataGridViewRow)(nextrow.row),
                            pagesets[currentpageset], rowstartlocation);
                    }
                }
                else
                {
                    // เสร็จสร้างแถวปั๊บ!!! ให้รีเซ็ตเป็นค่าเริ่มต้น
                    rowstartlocation = 0;
                }

                // ถ้าเราไม่มีข้อมูล (ไม่มีแถวอีก!!) พิมพ์แถวหมดแล้วอ่ะน้อออ
                if ((0 == rowstartlocation) && lastrowprinted >= rowstoprint.Count - 1)
                    continueprinting = false;

            }

            // บันทึกแถว (log rows) ข้ามแล้ว!!!
            if (EnableLogging)
            {
                Logger.LogInfoMsg(String.Format("Print Page printed rows {0} to {1}", firstrow, currentrow));
                PageDef pageset = pagesets[currentpageset];
                String columnlist = "";

                // ทำรายการคอลัมน์ทั้งหมดที่พิมพ์บนหน้านี้ 
                for (int i = 0; i < pageset.colstoprint.Count; i++)
                    columnlist = String.Format("{0},{1}", columnlist,
                        ((DataGridViewColumn)(pageset.colstoprint[i])).Index);

                Logger.LogInfoMsg(String.Format("Print Page printed columns {0}", columnlist.Substring(1)));
            }

            // พิมพ์ส่วนท้าย (print footer)
            if (PrintFooter)
            {
                // พิมพ์ "ส่วนท้าย" ("footer") เพื่อให้ที่เราพิมพ์จะ 'อยู่ด้านบน' ('on top')
                ImbeddedImageList.Where(p => p.ImageLocation == Location.Footer).DrawImbeddedImage(g, pagesets[currentpageset].printWidth,
                    pageHeight, pagesets[currentpageset].margins);

                //หมายเหตุ: บังคับให้ printpos ที่ด้านล่างของหน้า
                printpos = pageHeight - footerHeight - pagesets[currentpageset].margins.Bottom;  // - margins.Top

                // เพิ่มระยะห่าง
                printpos += footerspacing;

                // พิมพ์หมายเลขหน้า ถ้ามีอยู่
                if (!pagenumberontop)
                {
                    printpos = PrintPageNo(g, printpos);
                }

                if (0 != FooterHeight)
                    printfooter(g, ref printpos, pagesets[currentpageset]);
            }

            // ตรวจสอบดูว่านี่นะ!! เป็นหน้าสุดท้ายที่จะพิมพ์

            if (0 != rowstartlocation)
            {
                // เรายังไม่ได้ทำ rowนี้ (แถวนี้) 
                HasMorePages = true;
            }

            // เสร็จแล้ว!! เซ็ตชุดหน้านี้ เพื่อดูว่ามีจำนวนหน้าที่พิมพ์อีกหรือไม่
            if ((CurrentPage >= toPage) || (lastrowprinted >= rowstoprint.Count - 1))
            {
                // reset สำหรับชุดหน้าถัดไป / แจ้งให้ผู้ใช้ทราบว่าเราดำเนินการเสร็จแล้ว
                HasMorePages = DetermineHasMorePages();

                // reset ตัวนับ ถ้าเราพิมพ์จากหน้าตัวอย่าง
                rowstartlocation = 0;
                lastrowprinted = -1;
                CurrentPage = 0;
            }
            else
            {
                // เรายังไม่ได้ทำ
                HasMorePages = true;
            }

            return HasMorePages;
        }

        /// <summary>
        /// พิมพ์หมายเลขหน้า
        /// </summary>
        /// <param name="g"></param>
        /// <param name="printpos"></param>
        /// <returns></returns>
        private float PrintPageNo(Graphics g, float printpos)
        {
            if (pageno)
            {
                String pagenumber = pagetext + CurrentPage.ToString(CultureInfo.CurrentCulture);
                if (showtotalpagenumber)
                {
                    pagenumber += pageseparator + totalpages.ToString(CultureInfo.CurrentCulture);
                }
                if (1 < pagesets.Count)
                    pagenumber += parttext + (currentpageset + 1).ToString(CultureInfo.CurrentCulture);

                // จากนั้นก็พิมพ์
                printsection(g, ref printpos,
                    pagenumber, pagenofont, pagenocolor, pagenumberformat,
                    overridepagenumberformat, pagesets[currentpageset],
                    null, null);

                // ถ้าหมายเลขหน้าไม่ได้อยู่ในบรรทัดแยกกัน --> อย่า "ใช้" เป็นพื้นที่แนวตั้ง
                if (pagenumberonseparateline)
                    printpos += pagenumberHeight;
            }
            return printpos;
        }

        /// <summary>
        /// พิมพ์ส่วนหัว/ส่วนท้าย ของหมายเลขหน้า + ชื่อเรื่อง
        /// </summary>
        /// <param name="g">Graphic ที่จะพิมพ์</param> 
        /// <param name="pos">ตามพื้นที่แนวตั้งที่ใช้ ตำแหน่ง 'y'</param> 
        /// <param name="text">text ที่จะพิมพ์</param>
        /// <param name="font">แบบอักษรที่ใช้สำหรับการพิมพ์</param> 
        /// <param name="color">สีที่จะพิมพ์</param>
        /// <param name="format">String format for text</param>
        /// <param name="useroverride">เป็นจริงถ้าผู้ใช้ล้างการจัดตำแหน่ง</param>
        /// <param name="margins">ระยะขอบการพิมพ์ของตาราง</param>
        /// <param name="background">เติมพื้นหลัง; (อาจเป็นไม่มี ถ้าไม่มีพื้นหลัง)</param>
        /// <param name="border">ความหนา; (อาจเป็นไม่มี ถ้าไม่มีขอบ)</param>
        private void printsection(Graphics g, ref float pos, string text,
            Font font, Color color, StringFormat format, bool useroverride, PageDef pageset,
            Brush background, Pen border)
        {
            // วัดขนาด (measure string สายวัด)
            SizeF printsize = g.MeasureString(text, font, pageset.printWidth, format);

            // สร้างพื้นที่ที่จะพิมพ์ภายใน
            RectangleF printarea = new RectangleF((float)pageset.margins.Left, pos, (float)pageset.printWidth,
               printsize.Height);

            // วาดพื้นหลัง ถ้ามีการใช้ Brush
            if (null != background)
            {
                g.FillRectangle(background, printarea);
            }

            // วาดเส้นขอบ ถ้ามี Pen ไว้
            if (null != border)
            {
                g.DrawRectangle(border, printarea.X, printarea.Y, printarea.Width, printarea.Height);
            }

            // พิมพ์จริง!!!!!!!
            g.DrawString(text, font, new SolidBrush(color), printarea, format);
        }

        /// <summary>
        /// พิมพ์ส่วนท้าย จะจัดการระยะห่างส่วนท้ายและพิมพ์หมายเลขหน้า ที่ด้านล่างของหน้า (ถ้าหมายเลขหน้าไม่ได้อยู่ในส่วนหัว)
        /// </summary>
        /// <param name="g">Graphic context to print in</param> Graphic ที่จะพิมพ์
        /// <param name="pos">Track vertical space used; 'y' location</param> ตามพื้นที่แนวตั้งที่ใช้ ที่ตำแหน่ง 'y'
        /// <param name="margins">The table's print margins</param> ระยะขอบการพิมพ์ของตาราง
        private void printfooter(Graphics g, ref float pos, PageDef pageset)
        {

            // พิมพ์ส่วนท้าย
            printsection(g, ref pos, footer, footerfont, footercolor, footerformat,
                overridefooterformat, pageset, footerbackground, footerborder);
        }

        /// <summary>
        /// พิมพ์ส่วนหัวคอลัมน์ ข้อมูลรูปแบบการพิมพ์ส่วนใหญ่จะถูกดึงจาก DataGridView
        /// source DataGridView.
        /// </summary>
        /// <param name="g">Graphics Context to print within</param> บริบทกราฟิกที่จะพิมพ์ภายใน
        /// <param name="pos">Track vertical space used; 'y' location</param> ติดตามพื้นที่แนวตั้งที่ใช้ ตำแหน่ง 'y'
        /// <param name="pageset">Current pageset - defines columns and margins</param> ชุดหน้าปัจจุบัน - กำหนดคอลัมน์และระยะขอบ
        private void printcolumnheaders(Graphics g, ref float pos, PageDef pageset)
        {
            // ตามตำแหน่งการพิมพ์ทั่วทั้งหน้า ตำแหน่งเริ่มต้น --> ซ้าย
            // ปรับสำหรับส่วนหัวของแถว (ความกว้างของแถวคือ 0 ถ้าส่วนหัวของแถวไม่ได้ถูกพิมพ์)
            float xcoord = pageset.margins.Left + rowheaderwidth;

            // ตั้งค่า pen สำหรับการวาดเส้นตาราง
            Pen lines = new Pen(dgv.GridColor, 1);
 
            // พิมพ์ส่วนหัวคอลัมน์
            DataGridViewColumn col;
            for (int i = 0; i < pageset.colstoprint.Count; i++)
            {
                col = (DataGridViewColumn)pageset.colstoprint[i];

                // ความกว้างของเซลล์ คำนวณคอลัมน์ที่ใหญ่กว่าพื้นที่การพิมพ์!
                float cellwidth = (pageset.colwidths[i] > pageset.printWidth - rowheaderwidth ?
                    pageset.printWidth - rowheaderwidth : pageset.colwidths[i]);

                // รับสไตล์ของคอลัมน์
                DataGridViewCellStyle style = col.HeaderCell.InheritedStyle.Clone();
                if (ColumnHeaderStyles.ContainsKey(col.Name))
                {
                    style = ColumnHeaderStyles[col.Name];
                }

                // ตั้งค่าพื้นที่พิมพ์ของแต่ละเซลล์นี้ สำหรับเซลล์ที่ใหญ่กว่า (กว่าพื้นที่พิมพ์!)
                RectangleF cellprintarea = new RectangleF(xcoord, pos, cellwidth, colheaderheight);

                DrawCell(g, cellprintarea, style, col.HeaderCell, 0, columnheadercellformat, lines);

                xcoord += pageset.colwidths[i];
            }

            // เสร็จสิ้นทั้งหมดใช้พื้นที่แนวตั้ง "ใช้" รวมถึงพื้นที่ของเส้นขอบ
            pos += colheaderheight +
                (dgv.ColumnHeadersBorderStyle != DataGridViewHeaderBorderStyle.None ? lines.Width : 0);
        }

        /// <summary>
        /// พิมพ์หนึ่งแถวของ DataGridView ดึงข้อมูลรูปแบบการพิมพ์จาก DataGridView
        /// </summary>
        /// <param name="g">Graphics Context to print within</param>
        /// <param name="pos">Track vertical space used; 'y' location</param>
        /// <param name="row">The row that will be printed</param>
        /// <param name="pageset">Current Pageset - defines columns and margins</param>
        /// <param name="startline">Line no. in row to start printing text at</param>
        private float printrow(Graphics g, float finalpos, DataGridViewRow row, PageDef pageset,
            float startlocation)
        {
            // ตามตำแหน่งการพิมพ์ทั้งหน้า
            float xcoord = pageset.margins.Left;
            float pos = finalpos;

            // ตั้งค่า Pen สำหรับการวาดเส้นตาราง (สีของเส้น)
            Pen lines = new Pen(dgv.GridColor, 1);

            // คำนวณความกว้างของแถว ของคอลัมน์ที่กว้างกว่าพื้นที่พิมพ์!
            float rowwidth = (pageset.coltotalwidth > pageset.printWidth ? pageset.printWidth : pageset.coltotalwidth);

            // คำนวณความสูงของแถวเป็นพิกเซลเพื่อพิมพ์
            float rowheight = (rowstoprint[currentrow].height - startlocation) > (staticheight - pos)
                ? (staticheight - pos) : rowstoprint[currentrow].height - startlocation;

            // พิมพ์พื้นหลัง Row

            // รับสไตล์แถวปัจจุบัน + สไตล์ส่วนหัวปัจจุบัน
            DataGridViewCellStyle rowstyle = row.InheritedStyle.Clone();
            DataGridViewCellStyle headerstyle = row.HeaderCell.InheritedStyle.Clone();

            // กำหนดพิมพ์
            RectangleF printarea = new RectangleF(xcoord, pos, rowwidth,
                rowheight);

            // เติมพื้นหลังแถวเป็นสีเริ่มต้น
            g.FillRectangle(new SolidBrush(rowstyle.BackColor), printarea);

            // พิมพ์ส่วนหัวของแถว ถ้าตั้งค่าเป็น visible
            if ((bool)PrintRowHeaders)
            {
                // ตั้งค่าพื้นที่การพิมพ์ของแต่ละเซลล์
                RectangleF headercellprintarea = new RectangleF(xcoord, pos,
                    rowheaderwidth, rowheight);

                DrawCell(g, headercellprintarea, headerstyle, row.HeaderCell, startlocation,
                    rowheadercellformat, lines);

                // ตามพื้นที่แนวนอนที่ใช้
                xcoord += rowheaderwidth;
            }

            //  พิมพ์แถว (row) : เขียน/วาดแต่ละเซลล์
            DataGridViewColumn col;
            for (int i = 0; i < pageset.colstoprint.Count; i++)
            {
                // เข้าถึงเซลล์และคอลัมน์ที่กำลังพิมพ์
                col = (DataGridViewColumn)pageset.colstoprint[i];
                DataGridViewCell cell = row.Cells[col.Index];

                // ความกว้างของเซลล์ คำนวณคอลัมน์ที่ใหญ่กว่าพื้นที่การพิมพ์!
                float cellwidth = (pageset.colwidths[i] > pageset.printWidth - rowheaderwidth ?
                    pageset.printWidth - rowheaderwidth : pageset.colwidths[i]);

                // SLG 01112010 - วาดเฉพาะคอลัมน์ที่มีความกว้างจริง
                if (cellwidth > 0)
                {
                    // เอาสไตล์คอลัมน์ DGV และดูว่าเรามีการoverrideของคอลัมน์นี้หรือไม่
                    StringFormat finalformat = null;
                    Font cellfont = null;
                    DataGridViewCellStyle colstyle = GetStyle(row, col); // = row.Cells[col.Index].InheritedStyle.Clone(); 

                    // กำหนดรูปแบบ
                    buildstringformat(ref finalformat, colstyle, cellformat.Alignment, cellformat.LineAlignment,
                        cellformat.FormatFlags, cellformat.Trimming);
                    cellfont = colstyle.Font;

                    // กำหนดพื้นที่พิมพ์ โดยรวมแต่ละเซลล์
                    RectangleF cellprintarea = new RectangleF(xcoord, pos, cellwidth,
                        rowheight);

                    DrawCell(g, cellprintarea, colstyle, cell, startlocation, finalformat, lines);
                }
                // ตามพื้นที่แนวนอนที่ใช้
                xcoord += pageset.colwidths[i];
            }

            // row (แถว) ใช้พื้นที่แนวตั้ง "ใช้"
            return rowheight;
        }

        /// <summary>
        /// อนุญาตให้ override การวาดเซลล์ นี่คือการซับพอตตารางที่มี onPaint
        /// overridden สิ่งต่างๆ --> เช่น รูปภาพในแถวส่วนหัวและการพิมพ์ในแนวตั้ง
        /// </summary>
        /// <param name="g"></param>
        /// <param name="rowindex"></param>
        /// <param name="columnindex"></param>
        /// <param name="rectf"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        Boolean DrawOwnerDrawCell(Graphics g, int rowindex, int columnindex, RectangleF rectf,
            DataGridViewCellStyle style)
        {
            DGVCellDrawingEventArgs args = new DGVCellDrawingEventArgs(g, rectf, style,
                rowindex, columnindex);
            OnCellOwnerDraw(args);
            return args.Handled;
        }

        /// <summary>
        /// วาดเซลล์ ใช้สำหรับส่วนหัวคอลัมน์และแถวของแต่ละเซลล์
        /// </summary>
        /// <param name="g"></param>
        /// <param name="cellprintarea"></param>
        /// <param name="style"></param>
        /// <param name="cell"></param>
        /// <param name="startlocation"></param>
        /// <param name="cellformat"></param>
        /// <param name="lines"></param>
        void DrawCell(Graphics g, RectangleF cellprintarea, DataGridViewCellStyle style,
            DataGridViewCell cell, float startlocation, StringFormat cellformat, Pen lines)
        {
            //(แทนที่) ก็จะเอาออกไป๊!!!
            if (!DrawOwnerDrawCell(g, cell.RowIndex, cell.ColumnIndex, cellprintarea, style))
            {
                // บันทึกขอบเขตของรูปวาดต้นฉบับ original clipping bounds
                RectangleF clip = g.ClipBounds;

                // เติมพื้นหลังของเซลล์แบบเต็ม - ใช้สไตล์ที่เลือก
                //g.FillRectangle(new SolidBrush(colstyle.BackColor), cellprintarea);
                g.FillRectangle(new SolidBrush(style.BackColor), cellprintarea);

                // รีเซ็ตพื้นที่การพิมพ์ เซลล์แต่ละเซลล์โดยปรับ 'inward' ('ขาเข้า') สำหรับการขยายเซลล์
                RectangleF paddedcellprintarea = new RectangleF(cellprintarea.X + style.Padding.Left,
                    cellprintarea.Y + style.Padding.Top,
                    cellprintarea.Width - style.Padding.Right - style.Padding.Left,
                    cellprintarea.Height - style.Padding.Bottom - style.Padding.Top);

                // ตั้งค่าการตัดเป็นพื้นที่การพิมพ์ปัจจุบัน --> เช่น เซลล์ของเรา
                g.SetClip(cellprintarea);

                // กำหนดพื้นที่การพิมพ์ *actual*  ตามการเริ่มต้นที่กำหนด 
                // โดยลบตำแหน่งเริ่มต้น + เพิ่มความสูงของพื้นที่พิมพ์ จากการค่าเริ่มต้น
                RectangleF actualprint = new RectangleF(paddedcellprintarea.X, paddedcellprintarea.Y - startlocation,
                    paddedcellprintarea.Width, paddedcellprintarea.Height + startlocation);

                // วาดรูปแบบของเซลล์
                if (0 <= cell.RowIndex && 0 <= cell.ColumnIndex)
                {
                    if ("DataGridViewImageCell" == dgv.Columns[cell.ColumnIndex].CellType.Name)
                    {
                        // วาดภาพของ image cells
                        DrawImageCell(g, (DataGridViewImageCell)cell, actualprint);
                    }
                    else if ("DataGridViewCheckBoxCell" == dgv.Columns[cell.ColumnIndex].CellType.Name)
                    {
                        // วาดช่องทำ checkbox สำหรับ checkbox cells
                        DrawCheckBoxCell(g, (DataGridViewCheckBoxCell)cell, actualprint);
                    }
                    else
                    {
                        // จัดการกับการวาดภาพสำหรับ textbox, button, combobox แล้วก็ประเภท link cell (ไปนู่นนี่นั้น เว็บนู่นนี่นั้น ฟอร์มนู่นนี่นั้น จะไปไหนก็ไป๊!!!)

                        // วาดข้อความของเซลล์ที่ row / col
                        g.DrawString(cell.FormattedValue.ToString(), style.Font,
                            new SolidBrush(style.ForeColor), actualprint, cellformat);
                    }
                }
                else
                {
                    // วาดข้อความของเซลล์ที่ row / col
                    g.DrawString(cell.FormattedValue.ToString(), style.Font,
                        new SolidBrush(style.ForeColor), actualprint, cellformat);
                }

                // รีเซ็ต clipping bounds เป็น "normal"
                g.SetClip(clip);

                // วาดเส้นขอบ/ความหนา (borders) - เริ่มต้นการตั้งค่าเส้นขอบของ dgv และใช้พื้นที่การพิมพ์ของเซลล์ที่ไม่ได้เพิ่มมา
                if (dgv.CellBorderStyle != DataGridViewCellBorderStyle.None)
                    g.DrawRectangle(lines, cellprintarea.X, cellprintarea.Y, cellprintarea.Width, cellprintarea.Height);
            }
        }

        /// <summary>
        /// วาดรูปร่างของเซลล์ ที่เป็น checkbox เป็นเครื่องหมาย
        /// </summary>
        /// <param name="g"></param>
        /// <param name="checkboxcell"></param>
        /// <param name="rectf"></param>
        void DrawCheckBoxCell(Graphics g, DataGridViewCheckBoxCell checkboxcell, RectangleF rectf)
        {

            // สร้างgraphics ที่ไม่พิมพ์ซึ่งจะวาดตัว checkbox เป็นเครื่องหมาย
            Image i = new Bitmap((int)rectf.Width, (int)rectf.Height);
            Graphics tg = Graphics.FromImage(i);

            // กำหนด checked กับ ไม่checked (หรือ checkboxes 3 สถานะ)
            CheckBoxState state = CheckBoxState.UncheckedNormal;
            if (checkboxcell.ThreeState)
            {
                if (((CheckState)checkboxcell.EditedFormattedValue) == CheckState.Checked)
                    state = CheckBoxState.CheckedNormal;
                else if (((CheckState)checkboxcell.EditedFormattedValue) == CheckState.Indeterminate)
                    state = CheckBoxState.MixedNormal;
            }
            else
            {
                if ((Boolean)checkboxcell.EditedFormattedValue)
                    state = CheckBoxState.CheckedNormal;
            }

            // รับขนาดและที่ตั้งเพื่อพิมพ์ ทำ checkbox เป็นเครื่องหมาย - กึ่งกลางปัจจุบัน อาจเปลี่ยนแปลงได้ (ตามขนาดนู่นนี่นั้น ตอนที่พิมพ์)
            Size size = CheckBoxRenderer.GetGlyphSize(tg, state);
            int x = ((int)rectf.Width - size.Width) / 2;
            int y = ((int)rectf.Height - size.Height) / 2;

            // วาด checkbox ทำเครื่องหมายใน graphics ชั่วคราว
            CheckBoxRenderer.DrawCheckBox(tg, new Point(x, y), state);

            // คำนวณค่าเริ่มต้นของการวาดภาพ ตามการจัดเรียงเซลล์
            switch (checkboxcell.InheritedStyle.Alignment)
            {
                case DataGridViewContentAlignment.BottomCenter:
                    rectf.Y += y;
                    break;
                case DataGridViewContentAlignment.BottomLeft:
                    rectf.X -= x;
                    rectf.Y += y;
                    break;
                case DataGridViewContentAlignment.BottomRight:
                    rectf.X += x;
                    rectf.Y += y;
                    break;
                case DataGridViewContentAlignment.MiddleCenter:
                    break;
                case DataGridViewContentAlignment.MiddleLeft:
                    rectf.X -= x;
                    break;
                case DataGridViewContentAlignment.MiddleRight:
                    rectf.X += x;
                    break;
                case DataGridViewContentAlignment.TopCenter:
                    rectf.Y -= y;
                    break;
                case DataGridViewContentAlignment.TopLeft:
                    rectf.X -= x;
                    rectf.Y -= y;
                    break;
                case DataGridViewContentAlignment.TopRight:
                    rectf.X += x;
                    rectf.Y -= y;
                    break;
                case DataGridViewContentAlignment.NotSet:
                    break;
            }

            // วาดภาพแล้วทำ checkbox เป็นเครื่องหมาย เพื่อพิมพ์ออก
            g.DrawImage(i, rectf);

            // ล้างทิ้ง!! ออกไป๊!!!!!!
            tg.Dispose();
            i.Dispose();
        }

        /// <summary>
        /// วาดเซลล์ ที่มีภาพกำหนดไว้
        /// </summary>
        /// <param name="g"></param>
        /// <param name="imagecell"></param>
        /// <param name="rectf"></param>
        void DrawImageCell(Graphics g, DataGridViewImageCell imagecell, RectangleF rectf)
        {
            // ภาพที่จะวาด
            Image img;

            // ถ้าเราไม่มีค่า --> ก็แค่ออก (ก็คนไม่จำเป็นต้องเดินจากไป T^T)
            if ((null == imagecell.Value) || (typeof(DBNull) == imagecell.Value.GetType()))
                return;

            // ตรวจสอบประเภทของค่าเซลล์ภาพ - อาจไม่ใช่ประเภท "image" 
            if ("Image" == imagecell.ValueType.Name)
            {
                // หากเป็นประเภท "รูปภาพ" ให้โหลดโดยตรง
                img = (System.Drawing.Image)imagecell.Value;
            }
            else if ("Byte[]" == imagecell.ValueType.Name)
            {
                // หากไม่ใช่ประเภท "รูปภาพ" (เช่น โหลดจากฐานข้อมูลไปยังคอลัมน์ที่กำหนดไว้)
                // แปลงอาร์เรย์ไบต์เป็นรูปภาพ
                ImageConverter ic = new ImageConverter();
                img = (Image)ic.ConvertFrom((byte[])imagecell.Value);
            }
            else
                throw new Exception(String.Format("Unknown image cell underlying type: {0} in column {1}",
                    imagecell.ValueType.Name, imagecell.ColumnIndex));

            // ขอบเขตการตัด ส่วนของภาพ เพื่อให้พอดีกับสี่เหลี่ยมที่วาด
            Rectangle src = new Rectangle();

            // คำนวณ delta
            int dx = 0;
            int dy = 0;

            // วาดขนาดปกติ, ตัดกับเซลล์
            if ((DataGridViewImageCellLayout.Normal == imagecell.ImageLayout) ||
                (DataGridViewImageCellLayout.NotSet == imagecell.ImageLayout))
            {
                // คำนวณ delta เริ่มต้น ใช้เพื่อขยับภาพ
                dx = img.Width - (int)rectf.Width;
                dy = img.Height - (int)rectf.Height;

                // กำหนดความกว้างและความสูง ไปที่เซลล์
                if (0 > dx) rectf.Width = src.Width = img.Width; else src.Width = (int)rectf.Width;
                if (0 > dy) rectf.Height = src.Height = img.Height; else src.Height = (int)rectf.Height;

            }
            else if (DataGridViewImageCellLayout.Stretch == imagecell.ImageLayout)
            {
                // ยืดภาพให้พอดีกับขนาดของเซลล์
                src.Width = img.Width;
                src.Height = img.Height;

                // เปลี่ยน delta เริ่มต้นเป็น 0 --> จึงไม่ขยับรูปภาพ
                dx = 0;
                dy = 0;
            }
            else // DataGridViewImageCellLayout.Zoom
            {
                // ปรับขนาดภาพให้พอดีกับเซลล์
                src.Width = img.Width;
                src.Height = img.Height;

                float vertscale = rectf.Height / src.Height;
                float horzscale = rectf.Width / src.Width;
                float scale;

                
                // ใช้ตัวคูณขนาดที่เล็กลง เพื่อให้ชพอดีกับเซลล์ (แนวตั้ง*แนวนอน)
                if (vertscale > horzscale)
                {
                    // ใช้ระดับแนวนอน อย่าขยับภาพในแนวนอน!!! ห้ามแก้!!!!!
                    scale = horzscale;
                    dx = 0;
                    dy = (int)((src.Height * scale) - rectf.Height);
                }
                else
                {
                    // ใช้ขนาดแนวตั้ง อย่าขยับภาพในแนวตั้ง!!!! ห้ามแก้!!!!!
                    scale = vertscale;
                    dy = 0;
                    dx = (int)((src.Width * scale) - rectf.Width);
                }

                // กำหนดขนาด ให้ตรงกับภาพที่ปรับขนาด
                rectf.Width = src.Width * scale;
                rectf.Height = src.Height * scale;
            }

            // คำนวณการวาดภาพ ขึ้นอยู่กับจุดเริ่มต้นที่เรากำหนดไว้แล้ว
            switch (imagecell.InheritedStyle.Alignment)
            {
                case DataGridViewContentAlignment.BottomCenter:
                    if (0 > dy) rectf.Y -= dy; else src.Y = dy;
                    if (0 > dx) rectf.X -= dx / 2; else src.X = dx / 2;
                    break;
                case DataGridViewContentAlignment.BottomLeft:
                    if (0 > dy) rectf.Y -= dy; else src.Y = dy;
                    src.X = 0;
                    break;
                case DataGridViewContentAlignment.BottomRight:
                    if (0 > dy) rectf.Y -= dy; else src.Y = dy;
                    if (0 > dx) rectf.X -= dx; else src.X = dx;
                    break;
                case DataGridViewContentAlignment.MiddleCenter:
                    if (0 > dy) rectf.Y -= dy / 2; else src.Y = dy / 2;
                    if (0 > dx) rectf.X -= dx / 2; else src.X = dx / 2;
                    break;
                case DataGridViewContentAlignment.MiddleLeft:
                    if (0 > dy) rectf.Y -= dy / 2; else src.Y = dy / 2;
                    src.X = 0;
                    break;
                case DataGridViewContentAlignment.MiddleRight:
                    if (0 > dy) rectf.Y -= dy / 2; else src.Y = dy / 2;
                    if (0 > dx) rectf.X -= dx; else src.X = dx;
                    break;
                case DataGridViewContentAlignment.TopCenter:
                    src.Y = 0;
                    if (0 > dx) rectf.X -= dx / 2; else src.X = dx / 2;
                    break;
                case DataGridViewContentAlignment.TopLeft:
                    src.Y = 0;
                    src.X = 0;
                    break;
                case DataGridViewContentAlignment.TopRight:
                    src.Y = 0;
                    if (0 > dx) rectf.X -= dx; else src.X = dx;
                    break;
                case DataGridViewContentAlignment.NotSet:
                    if (0 > dy) rectf.Y -= dy / 2; else src.Y = dy / 2;
                    if (0 > dx) rectf.X -= dx / 2; else src.X = dx / 2;
                    break;
            }

            // วาดได้แล้ววววววววววววววววววววว
            g.DrawImage(img, rectf, src, GraphicsUnit.Pixel);

            //เสร็จเถอะ พอแล้วว 
        }
    }
}
