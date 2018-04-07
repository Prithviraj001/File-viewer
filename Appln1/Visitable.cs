using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace Appln1
{
    interface Visitable
    {
        void accept(Visitor v);
    }
}
