using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        string htmlFilePath = @"C:\Users\hntrk\source\repos\ohuton55\cs\learn_csharp\HtmlConvertToExcel\sample.html";
        Console.WriteLine(htmlFilePath);
        
        // Excelアプリケーションの作成
        Application excelApp = new Application();
        excelApp.Visible = false; // バックグラウンドで実行
        
        try
        {
            // 新しいワークブックを作成
            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
            
            // A1セルを選択（手動操作の再現）
            Microsoft.Office.Interop.Excel.Range cell = worksheet.Range["A1"];
            cell.Select();
            // HTMLファイルを開く（Workbooks.Openメソッドを使用）
            Workbook htmlWorkbook = excelApp.Workbooks.Open(
                htmlFilePath,
                UpdateLinks: 0,
                ReadOnly: true,
                Format: 6, // 6はHTMLフォーマットを意味します
                Editable: false
            );
            
            // 必要に応じて、データを元のワークブックにコピーするコードをここに追加
            
            // 保存
            string outputPath = Path.ChangeExtension(htmlFilePath, ".xlsx");
            htmlWorkbook.SaveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
            
            Console.WriteLine($"HTMLファイルが正常にインポートされ、{outputPath}として保存されました。");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラーが発生しました: {ex.Message}");
        }
        finally
        {
            // Excelを閉じる（重要！）
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}