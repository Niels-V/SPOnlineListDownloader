using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOnlineListDownloader
{
    class LocalListItem
    {
        public Dictionary<string, string> Fields { get; private set; }

        public LocalListItem(Dictionary<string,string> fields)
        {
            Fields = new Dictionary<string, string>();
            foreach (var field in fields)
            {
                Fields.Add(field.Key, field.Value);
            }
        }
    }
}
