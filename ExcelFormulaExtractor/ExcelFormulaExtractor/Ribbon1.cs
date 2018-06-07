﻿using System;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

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

        private Dictionary<AST.Address, string> getAllFormulas(Depends.DAG graph)
        {
            return
                graph
                .getAllFormulaAddrs()
                .Select(addr =>
                    new Tuple<AST.Address, string>(addr, graph.getFormulaAtAddress(addr))
                ).ToDictionary(tup => tup.Item1, tup => tup.Item2);
        }

        private Dictionary<AST.Address, AST.Expression> inlineFormulas(
                Dictionary<AST.Address,string> formulas,
                Depends.DAG graph
            )
        {
            return
                formulas
                .Select(kvp => 
                    new Tuple<AST.Address, AST.Expression>(kvp.Key, inlineExpression(kvp.Key, graph))
                )
                .Where(e => e != null)
                .ToDictionary(tup => tup.Item1, tup => tup.Item2);
        }

        private void toCSV(string[][] table)
        {
            //int offset = 0;
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
                //offset += b.Length;
            }

            f.Close();
        }

        private AST.Expression inlineExpression(AST.Address addr, Depends.DAG graph)
        {
            // get top-level AST
            var ast = graph.getASTofFormulaAt(addr);

            // merge subtrees
            return ExpressionTools.flattenedExpression(ast, graph);

        }

        private Dictionary<AST.Address, FPCoreAST.FPCore> convertFormulas(Dictionary<AST.Address, AST.Expression> exprs)
        {
            return
                exprs
                    .Select(kvp => {
                        try
                        {
                            var fpc = XL2FPCore.FormulaToFPCore(kvp.Value);
                            return new Tuple<AST.Address, FPCoreAST.FPCore>(kvp.Key, fpc);
                        } catch (XL2FPCore.InvalidExpressionException e)
                        {
                            return null;
                        }
                        
                    })
                    .Where(e => e != null)
                    .ToDictionary(tup => tup.Item1, tup => tup.Item2);
        }

        private string[][] coresToTable(Dictionary<AST.Address, string> formulas, Dictionary<AST.Address, AST.Expression> exprs, Dictionary<AST.Address, FPCoreAST.FPCore> cores)
        {
            var cores_arr = cores.ToArray();

            int COLS = 4;
            string[][] output = new string[cores.Count + 1][];
            for (int i = 0; i < cores.Count + 1; i++)
            {
                output[i] = new string[COLS];

                // header
                if (i == 0)
                {
                    output[0][0] = "address";
                    output[0][1] = "formula";
                    output[0][2] = "inlined";
                    output[0][3] = "fpcore";
                } else
                {
                    var addr = cores_arr[i - 1].Key;
                    output[i][0] = addr.A1Local();
                    output[i][1] = formulas[addr];
                    output[i][2] = exprs[addr].ToFormula;
                    output[i][3] = cores[addr].ToExpr(0);
                }
            }

            return output;
        }

        private void extract_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            // get all formulas
            var formulas = getAllFormulas(graph);

            // print
            System.Windows.Forms.MessageBox.Show(prettyPrintFormulaDict(formulas));

            // get inlined ASTs
            var asts = inlineFormulas(formulas, graph);

            // print
            System.Windows.Forms.MessageBox.Show(prettyPrintExpressionDict(formulas, asts));
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
            var ast = inlineExpression(addr, graph);

            // convert AST back to expression
            var f_in = ast.ToFormula;

            // print
            System.Windows.Forms.MessageBox.Show("cell: " + addr.A1Local() + "\n\n" + f + "\n\nconverted to\n\n" + f_in);
        }

        private void ExtractToFPCore_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = new Depends.DAG(getWorkbook(), getApp(), ignore_parse_errors: false, dagBuilt: new DateTime());

            // get all formulas
            var formulas = getAllFormulas(graph);

            // get inlined ASTs
            var asts = inlineFormulas(formulas, graph);

            // convert to FPCore
            var fpcores = convertFormulas(asts);

            // get outputs as table
            var table = coresToTable(formulas, asts, fpcores);

            // prompt user to save as CSV
            toCSV(table);
        }
    }
}
