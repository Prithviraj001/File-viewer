using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace Appln1
{
    class Image : Visitable
    {
        private string filename;
        public BitmapImage bi;

        public Image()
        {
            filename = "";
        }

        public void SetFilename(string f)
        {
            filename = f;
        }

        public string GetFilename()
        {
            return filename;
        }


        public void accept(Visitor v)
        {
            bi = v.visit(this);
        }
    }
}
