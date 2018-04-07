using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

namespace Appln1
{
    interface Visitor
    {
        BitmapImage visit(Image Im);
        IDocumentPaginatorSource visit(Excel e);
        IDocumentPaginatorSource visit(Word w);
        string visit(PDF pdf);
        string visit(Video v);

    }
}
