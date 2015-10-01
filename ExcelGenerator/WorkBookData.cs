using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    public class WorkBookData
    {
        public DataSet data { get; set; }
        public List<string> sheetNames { get; set; }

        public WorkBookData(DataSet data, List<string> sheetNames)
        {
            this.data = data;
            this.sheetNames = sheetNames;
        }
    }
}
