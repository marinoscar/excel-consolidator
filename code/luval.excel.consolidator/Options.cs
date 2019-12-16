using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace luval.excel.consolidator
{
    public class Options
    {
        public int HeaderStartRow { get; set; }
        public int HeaderEndRow { get; set; }
        public int DataStartRow { get; set; }
        public int DataStartColumn { get; set; }
        public int DataEndColumn { get; set; }

        public Options()
        {
            HeaderStartRow = 1;
            HeaderEndRow = 8;
            DataStartRow = 22;
            DataStartColumn = 1;
            DataEndColumn = 31;
        }
    }
}
