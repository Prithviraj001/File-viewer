using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using System.Windows.Xps.Packaging;

namespace Appln1
{
    class OpenVisitor : Visitor
    {
        public BitmapImage visit(Image Im)
        {
            string filename = Im.GetFilename();

            try
            {
                if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                {
                    MessageBox.Show("FILE DOESN'T EXIsTS");
                    return null;
                }
                else
                {
                    BitmapImage temp = new BitmapImage();
                    temp.BeginInit();
                    temp.UriSource = new Uri(filename);
                    temp.EndInit();
                    return temp;
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at OpenImageAction : " + exp.Message);
                return null;     
            }
        }

        public IDocumentPaginatorSource visit(Excel e)
        {
            string filename = e.GetFilename();

            try
            {
                if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                {
                    MessageBox.Show("FILE DOESN'T EXIsTS");
                    return null;
                }
                else
                {
                    string convertedXpsDoc = string.Concat(Path.GetTempPath(), "\\", Guid.NewGuid().ToString(), ".xps");
                    XpsDocument xpsDocument = e.ConvertExcelToXps(filename, convertedXpsDoc);
                    if (xpsDocument == null)
                    {
                        return null;
                    }

                    return xpsDocument.GetFixedDocumentSequence();
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at OpenWordAction : " + exp.Message);
                return null;
            }
        }

        public IDocumentPaginatorSource visit(Word w)
        {
            string filename = w.GetFilename();

            try
            {
                if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                {
                    MessageBox.Show("FILE DOESN'T EXIsTS");
                    return null;
                }
                else
                {
                    string convertedXpsDoc = string.Concat(Path.GetTempPath(), "\\", Guid.NewGuid().ToString(), ".xps");
                    XpsDocument xpsDocument = w.ConvertWordToXps(filename, convertedXpsDoc);
                    if (xpsDocument == null)
                    {
                        return null;
                    }

                    return xpsDocument.GetFixedDocumentSequence();
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at OpenWordAction : " + exp.Message);
                return null;
            }
        }

        public string visit(PDF pdf)
        {
            string filename = pdf.GetFilename();

            try
            {
                if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                {
                    MessageBox.Show("FILE DOESN'T EXIsTS");
                    return null;
                }
                else
                {
                    return filename;
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at OpenPDFAction : " + exp.Message);
                return null;
            }
        }

        public string visit(Video v)
        {
            string filename = v.GetFilename();

            try
            {
                if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
                {
                    MessageBox.Show("FILE DOESN'T EXIsTS");
                    return null;
                }
                else
                {
                    return filename;
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine("Exception at OpenVideoAction : " + exp.Message);
                return null;
            }
        }
    }
}
