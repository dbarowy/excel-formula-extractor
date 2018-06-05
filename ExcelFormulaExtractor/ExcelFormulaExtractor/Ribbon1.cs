using System;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace ExcelFormulaExtractor
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private string prettyPrintFormulaDict(Dictionary<AST.Address,string> d)
        {
            var sb = new StringBuilder();
            foreach (var kvp in d)
            {
                sb.Append(kvp.Key.A1Local());
                sb.Append(" : ");
                sb.Append(kvp.Value);
                sb.Append("\n");
            }
            return sb.ToString();
        }

        private string prettyPrintExpressionDict(Dictionary<AST.Address, string> d1, Dictionary<AST.Address, AST.Expression> d2)
        {
            var sb = new StringBuilder();
            foreach (var kvp in d2)
            {
                sb.Append(kvp.Key.A1Local());
                sb.Append(" :\n");
                sb.Append(d1[kvp.Key]);
                sb.Append(" rewritten to:\n");
                sb.Append(kvp.Value.ToFormula);
                sb.Append("\n\n");
            }
            return sb.ToString();
        }

        private void extract_Click(object sender, RibbonControlEventArgs e)
        {
            Application app = Globals.ThisAddIn.Application.Application;
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            var graph = new Depends.DAG(wb, app, ignore_parse_errors: false, dagBuilt: new DateTime());
            var formulas =
                graph
                .getAllFormulaAddrs()
                .Select(addr =>
                    new Tuple<AST.Address, string>(addr, graph.getFormulaAtAddress(addr))
                ).ToDictionary(tup => tup.Item1, tup => tup.Item2);
            System.Windows.Forms.MessageBox.Show(prettyPrintFormulaDict(formulas));
            var asts =
                formulas
                .Select(kvp => new Tuple<AST.Address, AST.Expression>(kvp.Key, inlineExpression(kvp.Key, graph)))
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);
            System.Windows.Forms.MessageBox.Show(prettyPrintExpressionDict(formulas, asts));
        }

        private AST.Expression inlineExpression(AST.Address addr, Depends.DAG graph)
        {
            // get top-level AST
            var ast = graph.getASTofFormulaAt(addr);

            // merge subtrees
            return ExpressionTools.flattenedExpression(ast, graph);
        }
    }
}
