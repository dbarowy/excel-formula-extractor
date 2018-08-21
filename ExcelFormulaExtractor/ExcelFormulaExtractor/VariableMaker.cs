using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFormulaExtractor
{
    class VariableMaker
    {
        private int variable = 0;

        public string nextVariable()
        {
            var v = variable++;
            return "x" + v.ToString();
        }
    }
}
