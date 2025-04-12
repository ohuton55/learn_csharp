using System.Drawing.Printing;
using System.Management;
using System.Transactions;
using Microsoft.Office.Interop.Excel;
using PdfiumViewer;
//using Spire.Pdf;
//using Spire.Pdf.Print;

namespace PrintDuplexTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private int currentPage = 0;
        private int totalPages = 2; // Total number of pages to print
        private Workbook workbook = null;
        string printerName = string.Empty;
        string pdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Book1.pdf");

        private void printButton_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string defaultFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Book1.xlsx");
                workbook = excelApp.Workbooks.Open(defaultFilePath);
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    workbook.Worksheets[i].PageSetup.PrintArea = "A1:M29";
                    workbook.Worksheets[i].PageSetup.FitToPagesWide = 1; // fit to 1 page wide
                    workbook.Worksheets[i].PageSetup.Orientation = XlPageOrientation.xlPortrait; // fit to 1 page tall
                    workbook.Worksheets[i].PageSetup.CenterVertically = true; // center vertically
                    workbook.Worksheets[i].PageSetup.CenterHorizontally = true; // center horizontally
                }

                workbook.Worksheets.Select();
                workbook.ExportAsFixedFormat(
                    Type: XlFixedFormatType.xlTypePDF,
                    Filename: pdfPath,
                    Quality: XlFixedFormatQuality.xlQualityStandard,
                    IgnorePrintAreas: false
                );

                workbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                PrintProcess();
            }
            catch
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            
        }
        private void PrintProcess()
        {
            try
            {
                currentPage = 0; //reset page count
                PageSettings pagesetting = new PageSettings();
                pagesetting.Margins = new Margins(0, 0, 0, 0); // Set margins to zero

                PrintDocument pd = new PrintDocument();
                pd.OriginAtMargins = true; // Set to true to use the margins set above
                pd.DefaultPageSettings = pagesetting;

                //pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                pd.PrintPage += new PrintPageEventHandler(pd_testPrintPage);
                printerName = comboBox1.SelectedItem.ToString();
                pd.PrinterSettings.PrinterName = printerName; // Set your printer name here
                if (!pd.PrinterSettings.CanDuplex)
                {
                    MessageBox.Show("This printer does not support duplex printing.");
                    return;
                }
                pd.PrinterSettings.Duplex = Duplex.Vertical; // Set duplex mode to Vertical
                if (pd.PrinterSettings.IsValid)
                {
                    pd.Print();
                }

            }
            catch
            {
                MessageBox.Show("An error occurred while printing.");
            }
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            try
            {
                using (PdfDocument document = PdfDocument.Load(pdfPath))
                {
                    var image = document.Render(currentPage, 300, 300, true);

                    // 印刷できる範囲（用紙 - プリンターのマージン）を取得
                    var bounds = e.MarginBounds;

                    float scaleHorizontal = (float)bounds.Height / image.Height; // 水平方向に合わせるスケール
                    float scaleVertical = (float)bounds.Width / image.Width; // 垂直方向に合わせるスケール
                    float scale = Math.Min(scaleHorizontal, scaleVertical); // 縦横比を維持するために小さい方を選択

                    int scaleWidth = (int)(image.Width * scale);        // スケール後の幅
                    int scaleHeight = (int)(image.Height * scale);      // スケール後の高さ

                    int offsetX = bounds.X + (bounds.Width - scaleWidth) / 2; // X軸のオフセット
                    //int offsetX = bounds.X / 2; // X軸のオフセット
                    int offsetY = bounds.Y + (bounds.Height - scaleHeight) / 2; // Y軸のオフセット
                    //int offsetY = bounds.Y / 2; // Y軸のオフセット

                    e.Graphics.DrawImage(image, offsetX, offsetY, scaleWidth, scaleHeight);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            currentPage++;

            e.HasMorePages = (currentPage < totalPages);
        }

        private void pd_testPrintPage(object sender, PrintPageEventArgs e)
        {
            // 印刷可能な領域を枠で囲む。
            e.Graphics.DrawRectangle(
                new Pen(Color.Blue, 2),
                e.PageSettings.PrintableArea.X,
                e.PageSettings.PrintableArea.Y,
                e.PageSettings.PrintableArea.Width,
                e.PageSettings.PrintableArea.Height
                );
        }

        private void AddSettingInitializeComponent()
        {
            ManagementScope objScope = new ManagementScope(ManagementPath.DefaultPath);
            objScope.Connect();

            SelectQuery selectQuery = new SelectQuery("SELECT * FROM Win32_Printer");
            ManagementObjectSearcher MOS = new ManagementObjectSearcher(objScope, selectQuery);
            ManagementObjectCollection MOC = MOS.Get();
            foreach (ManagementObject managementObject in MOC)
            {
                // Get the printer name
                comboBox1.Items.Add(managementObject["Name"].ToString());
            }
            MOC.Dispose();
            MOS.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AddSettingInitializeComponent();
        }
    }
}
