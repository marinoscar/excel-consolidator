using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.excel.consolidator
{
    public class Arguments : Options
    {
        private List<string> _args;
        public Arguments(IEnumerable<string> args)
        {
            _args = new List<string>(args);
            if (!string.IsNullOrWhiteSpace(GetVal("-hsr")))
                HeaderStartRow = Convert.ToInt32(GetVal("-hsr"));
            if (!string.IsNullOrWhiteSpace(GetVal("-her")))
                HeaderEndRow = Convert.ToInt32(GetVal("-her"));
            if (!string.IsNullOrWhiteSpace(GetVal("-dsr")))
                DataStartRow = Convert.ToInt32(GetVal("-dsr"));
            if (!string.IsNullOrWhiteSpace(GetVal("-dsc")))
                DataStartColumn = Convert.ToInt32(GetVal("-dsc"));
            if (!string.IsNullOrWhiteSpace(GetVal("-dec")))
                DataEndColumn = Convert.ToInt32(GetVal("-dec"));
            if (!string.IsNullOrWhiteSpace(GetVal("-i")))
                InputFolder = GetVal("-i");
            else
                InputFolder = Environment.CurrentDirectory;
            if (!string.IsNullOrWhiteSpace(GetVal("-o")))
                OutputFile = GetVal("-o");
            else
                OutputFile = "CONSOLIDATED.XLSX";
            if (!string.IsNullOrWhiteSpace(GetVal("-l")))
                LogFile = GetVal("-l");
            else
                LogFile = "LOG.TXT";
        }

        public string InputFolder { get; set; }
        public string OutputFile { get; set; }
        public string LogFile { get; set; }

        private string GetVal(string switchVal)
        {
            if (!_args.Contains(switchVal)) return null;
            var idx = _args.IndexOf(switchVal);
            if (idx + 1 >= _args.Count) return null;
            return _args[idx + 1];
        }
    }
}
