using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOnlineListDownloader
{
    class LocalList
    {
        public LinkedList<LocalListItem> Items
        {
            get;
            private set;
        }

        public string Title
        {
            get;
            set;
        }

        public LocalList()
        {
            Items = new LinkedList<LocalListItem>();
        }
    }
}
