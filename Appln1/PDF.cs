using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Appln1
{
    class PDF : Visitable
    {
        private string filename;
        public string info;

        public PDF()
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
            info = v.visit(this);
        }
    }
}
