using System;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Countable = ExceLint.Countable;
using FPCoreOption = Microsoft.FSharp.Core.FSharpOption<FPCoreAST.FPCore>;

namespace ExcelFormulaExtractor
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private Application getApp()
        {
            return Globals.ThisAddIn.Application.Application;
        }

        private Workbook getWorkbook()
        {
            return Globals.ThisAddIn.Application.ActiveWorkbook;
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

        private Dictionary<AST.Address, string> getAllFormulas(Depends.DAG graph, bool showProgress)
        {
            var frms = graph.getAllFormulaAddrs();

            if (showProgress)
            {
                Progress p = new Progress("Marshaling", frms.Length);
                p.Show();
                p.Refresh();

                var d =
                frms
                .Select(addr => {
                    var t = new Tuple<AST.Address, string>(addr, graph.getFormulaAtAddress(addr));
                    p.increment();
                    return t;
                }).ToDictionary(tup => tup.Item1, tup => tup.Item2);

                p.Hide();
                
                return d;
            } else
            {
                return
                frms
                .Select(addr => {
                    return new Tuple<AST.Address, string>(addr, graph.getFormulaAtAddress(addr));
                }).ToDictionary(tup => tup.Item1, tup => tup.Item2);
            }
        }

        private void toCSV(string[][] table)
        {
            var sfd = new System.Windows.Forms.SaveFileDialog();
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var f = sfd.OpenFile();
                for (int r = 0; r < table.Length; r++)
                {
                    var sb = new StringBuilder();
                    for (int c = 0; c < table[r].Length; c++)
                    {
                        string cell = table[r][c];
                        sb.Append("\"");
                        sb.Append(cell);
                        sb.Append("\"");
                        if (c < table[r].Length - 1)
                        {
                            sb.Append(",");
                        }
                        else
                        {
                            sb.Append("\n");
                        }
                    }
                    var b = Encoding.ASCII.GetBytes(sb.ToString());
                    f.Write(b, 0, b.Length);
                }

                f.Close();
            }
        }

        private Dictionary<List<AST.Address>, FPCoreOption> newConvertFormulaGroups(
            Dictionary<Countable, List<AST.Address>> grps,
            Dictionary<AST.Address, ExpressionTools.EData> fexprs,
            Depends.DAG graph,
            bool showProgress)
        {
            throw new Exception("heh");
        }

        private string[][] coresToTable(
            Dictionary<AST.Address, string> formulas,
            Dictionary<AST.Address, ExpressionTools.EData> exprs,
            Dictionary<List<AST.Address>, FPCoreOption> cores)
        {
            var cores_arr = cores.ToArray();

            int COLS = 3;
            string[][] output = new string[cores.Count + 1][];
            for (int i = 0; i < cores.Count + 1; i++)
            {
                output[i] = new string[COLS];

                // header
                if (i == 0)
                {
                    output[0][0] = "address";
                    output[0][1] = "formula";
                    //output[0][2] = "inlined";
                    output[0][2] = "fpcore";
                } else
                {
                    var addrs = cores_arr[i - 1].Key;
                    output[i][0] = String.Join("; ", addrs.Select(addr => addr.A1Local()));
                    output[i][1] = String.Join("; ", addrs.Select(addr => formulas[addr]));
                    if (FPCoreOption.get_IsSome(cores[addrs]))
                    {
                        output[i][2] = cores[addrs].Value.ToExpr(0);
                    } else
                    {
                        output[i][2] = "No conversion available for this formula.";
                    }
                }
            }

            return output;
        }

        private void ExtractThis_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void ExtractToFPCore_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void checkForUnsupportedFormulas_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void extractTest_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            // get all formulas
            var formulas = getAllFormulas(graph, showProgress: true);

            // extract all
            var fpcores = ExtractionLogic.Extract.extractAll(graph, formulas);

            var output = String.Join("\n\n", fpcores);
            System.Windows.Forms.MessageBox.Show(output);
        }
    }
}
