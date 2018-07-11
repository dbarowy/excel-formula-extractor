using System;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using PreList = System.Collections.Generic.List<System.Collections.Generic.Dictionary<AST.Address, double>>;
using Vector = ExceLint.Vector.Vector;
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
                    p.increment();
                    return new Tuple<AST.Address, string>(addr, graph.getFormulaAtAddress(addr));
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

        private Dictionary<AST.Address, ExpressionTools.FExpression> inlineFormulas(
                Dictionary<AST.Address,string> formulas,
                Depends.DAG graph,
                bool showProgress
            )
        {
            if (showProgress)
            {
                var p = new Progress("Inline", formulas.Count);
                p.Show();
                p.Refresh();

                var d =
                formulas
                .Select(kvp => {
                    p.increment();
                    return new Tuple<AST.Address, ExpressionTools.FExpression>(kvp.Key, inlineExpression(kvp.Key, graph));
                })
                .Where(e => e != null)
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);

                p.Hide();

                return d;
            } else
            {
                return
                formulas
                .Select(kvp =>
                    new Tuple<AST.Address, ExpressionTools.FExpression>(kvp.Key, inlineExpression(kvp.Key, graph))
                )
                .Where(e => e != null)
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);
            }
        }

        private void toCSV(string[][] table)
        {
            var sfd = new System.Windows.Forms.SaveFileDialog();
            sfd.ShowDialog();
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

        private ExpressionTools.FExpression inlineExpression(AST.Address addr, Depends.DAG graph)
        {
            // get top-level AST
            var ast = graph.getASTofFormulaAt(addr);

            // merge subtrees
            return ExpressionTools.flattenExpression(ast, graph);
        }

        private Dictionary<AST.Address, FPCoreAST.FPCore> convertFormulas(Dictionary<AST.Address, ExpressionTools.FExpression> fexprs)
        {
            return
                fexprs
                    .Select(kvp => {
                        try
                        {
                            var prelist = new PreList();
                            prelist.Add(kvp.Value.Data);
                            var fpc = XL2FPCore.FormulaToFPCore(kvp.Value.Expression, prelist);
                            return new Tuple<AST.Address, FPCoreAST.FPCore>(kvp.Key, fpc);
                        } catch (XL2FPCore.InvalidExpressionException)
                        {
                            return null;
                        }
                        
                    })
                    .Where(e => e != null)
                    .ToDictionary(tup => tup.Item1, tup => tup.Item2);
        }

        private Dictionary<List<AST.Address>, FPCoreOption> convertFormulaGroups(
            Dictionary<Countable, List<AST.Address>> grps,
            Dictionary<AST.Address, ExpressionTools.FExpression> fexprs,
            Depends.DAG graph,
            bool showProgress)
        {
            var d = new Dictionary<List<AST.Address>, FPCoreOption>();

            Progress p = null;
            if (showProgress)
            {
                p = new Progress("Convert", grps.Count);
                p.Show();
                p.Refresh();
            }

            foreach(var grp in grps)
            {
                var vector = grp.Key;
                var formulas = grp.Value;
                var data = new PreList();
                foreach (var f in formulas)
                {
                    data.Add(fexprs[f].Data);
                }

                Func<FPCoreAST.FPCore> convert = () =>
                {
                    var ast = graph.getASTofFormulaAt(formulas.First());
                    return XL2FPCore.FormulaToFPCore(ast, data);
                };

                try
                {
                    if (showProgress)
                    {
                        p.increment();
                    }
                    d.Add(formulas, FPCoreOption.Some(convert()));
                } catch(Exception e)
                {
                    d.Add(formulas, FPCoreOption.None);
                }
            }

            if (showProgress)
            {
                p.Hide();
            }

            return d;
        }

        private string[][] coresToTable(
            Dictionary<AST.Address, string> formulas,
            Dictionary<AST.Address, ExpressionTools.FExpression> exprs,
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
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address addr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, getWorkbook());

            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            if (!graph.isFormula(addr))
            {
                System.Windows.Forms.MessageBox.Show("Selected cell is not a formula.");
                return;
            }

            // get original expression
            var f = graph.getFormulaAtAddress(addr);

            // get inlined AST
            var fexpr = inlineExpression(addr, graph);

            // init list
            var prelist = new PreList();
            prelist.Add(fexpr.Data);

            // convert to FPCore
            var fpc = XL2FPCore.FormulaToFPCore(fexpr.Expression, prelist);

            // stringify FPCore
            var f_in = fpc.ToExpr(0);

            // print
            System.Windows.Forms.MessageBox.Show("cell: " + addr.A1Local() + "\n\n" + f + "\n\nconverted to\n\n" + f_in);
        }

        private Dictionary<Countable, List<AST.Address>> groupFormulasByVector(Dictionary<AST.Address,string> addrs, Depends.DAG graph, bool showProgress)
        {
            Dictionary<AST.Address, Countable> vs;
            if (showProgress)
            {
                var p = new Progress("Group", addrs.Count);
                p.Show();
                p.Refresh();

                vs = addrs
                .Select(kvp => {
                    p.increment();
                    return new Tuple<AST.Address, Countable>(kvp.Key, Vector.run(kvp.Key, graph).ToCVectorResultant);
                })
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);

                p.Hide();
            } else
            {
                vs = addrs
                .Select(kvp => new Tuple<AST.Address, Countable>(kvp.Key, Vector.run(kvp.Key, graph).ToCVectorResultant))
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);
            }
            

            var grps = vs.GroupBy(kvp => kvp.Value);

            var d = new Dictionary<Countable, List<AST.Address>>();

            foreach (var grp in grps)
            {
                var xs = new List<AST.Address>();
                foreach (var kvp in grp)
                {
                    xs.Add(kvp.Key);
                }
                d.Add(grp.Key, xs);
            }

            return d;
        }

        private void ExtractToFPCore_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            // get all formulas
            var formulas = getAllFormulas(graph, showProgress: true);

            // get inlined ASTs
            var fexprs = inlineFormulas(formulas, graph, showProgress: true);

            // which formulas are the same?
            var fgrps = groupFormulasByVector(formulas, graph, showProgress: true);

            // for each group, generate a single formula with a bunch of preconditions
            var fpcores = convertFormulaGroups(fgrps, fexprs, graph, showProgress: true);

            // get outputs as table
            var table = coresToTable(formulas, fexprs, fpcores);

            // prompt user to save as CSV
            toCSV(table);
        }

        private void checkForUnsupportedFormulas_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            // get all formulas
            var formulas = getAllFormulas(graph, showProgress: true);

            // get inlined ASTs
            var fexprs = inlineFormulas(formulas, graph, showProgress: true);

            // which formulas are the same?
            var fgrps = groupFormulasByVector(formulas, graph, showProgress: true);

            // for each group, generate a single formula with a bunch of preconditions
            var fpcores = convertFormulaGroups(fgrps, fexprs, graph, showProgress: true);

            // find all of the formulas that convert to None
            var failures = fpcores.Where(kvp => FPCoreOption.get_IsNone(kvp.Value));

            // display on screen
            if (failures.Count() == 0)
            {
                System.Windows.Forms.MessageBox.Show("All " + formulas.Count() + " formulas are supported.");
            } else
            {
                // convert formulas to string
                var msg = String.Join("\n", failures.Select(kvp => {
                    var addr = kvp.Key.First();
                    var form = graph.getFormulaAtAddress(addr);
                    var a1local = addr.A1Local();
                    string output = a1local + ": " + form;
                    return output;
                }));

                System.Windows.Forms.MessageBox.Show(msg);
            }
        }
    }
}
