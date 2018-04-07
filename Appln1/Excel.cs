using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Xps.Packaging;
using Microsoft.Office.Interop.Excel;
using Exc= Microsoft.Office.Interop.Excel;

namespace Appln1
{
    class Excel : Visitable
    {
        private String filename;
        public IDocumentPaginatorSource info;

        public Excel()
        {
            filename = "";
        }

        public void SetFilename(String f)
        {
            filename = f;
        }

        public String GetFilename()
        {
            return filename;
        }

        public void accept(Visitor v)
        {
            info = v.visit(this);
        }

        public XpsDocument ConvertExcelToXps(string excelFilename, string xpsFilename)
        {
            Exc.Application excelApp = new Exc.Application();

            excelApp.DisplayAlerts = false;

            excelApp.Visible = false;

            Workbook excelWorkbook = excelApp.Workbooks.Open(excelFilename);

            // add below

            xpsFilename = (new DirectoryInfo(excelFilename)).FullName;

            xpsFilename = xpsFilename.Replace(new FileInfo(excelFilename).Extension, "") + ".xps";

            excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypeXPS, Filename: xpsFilename, OpenAfterPublish: false);

            excelWorkbook.Close(false, null, null);

            excelApp.Quit();

            // release com object

            Marshal.ReleaseComObject(excelApp);

            excelApp = null;

            XpsDocument xpsexcel = new XpsDocument(xpsFilename, FileAccess.Read, CompressionOption.NotCompressed);

            return xpsexcel;

        }
    }
}
