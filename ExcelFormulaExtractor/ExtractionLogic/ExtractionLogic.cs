using System;
using System.Linq;
using System.Collections.Generic;
using PreList = System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, double>>;
using Vector = ExceLint.Vector.Vector;
using Countable = ExceLint.Countable;
using MemoDBO = System.Collections.Generic.Dictionary<AST.Address, System.Tuple<AST.Expression, Microsoft.FSharp.Collections.FSharpMap<AST.Address, double>>>;
using MemoDBOpt = Microsoft.FSharp.Core.FSharpOption<System.Collections.Generic.Dictionary<AST.Address, System.Tuple<AST.Expression, Microsoft.FSharp.Collections.FSharpMap<AST.Address, double>>>>;
using Fingerprint = ExceLint.Countable;
using Source = AST.Address;
using Functions = System.Collections.Generic.Dictionary<ExceLint.Countable, System.Collections.Generic.List<ExpressionTools.Vector[]>>;
using Variables = System.Collections.Generic.Dictionary<ExceLint.Countable, System.Collections.Generic.Dictionary<ExceLint.Countable, string>>;
using Provenance = System.Collections.Generic.Dictionary<ExceLint.Countable, System.Collections.Generic.List<AST.Address>>;

namespace ExtractionLogic
{
    public class Group
    {
        private readonly Functions _functions;
        private readonly Variables _variables;
        private readonly Provenance _provenance;

        public Group(Functions f, Variables v, Provenance p)
        {
            _functions = f;
            _variables = v;
            _provenance = p;
        }

        public Functions Functions
        {
            get { return _functions; }
        }

        public Variables Variables
        {
            get { return _variables; }
        }

        public Provenance Provenance
        {
            get { return _provenance; }
        }
    }

    public interface ProblemReport { }
    public class NoReferencesForInvocation : ProblemReport
    {
        private readonly Fingerprint _f;
        private readonly string _w;
        public NoReferencesForInvocation(Fingerprint f, string worksheet)
        {
            _f = f;
            _w = worksheet;
        }
        public override string ToString()
        {
            return "No references for invocation of fingerprint " + _f.ToString() + " on worksheet " + _w;
        }
    }
    public class AliasingDetected : ProblemReport
    {
        private readonly Fingerprint _f;
        private readonly string _w;
        public AliasingDetected(Fingerprint f, string worksheet)
        {
            _f = f;
            _w = worksheet;
        }
        public override string ToString()
        {
            return "Aliasing detected for fingerprint " + _f.ToString() + " on worksheet " + _w;
        }
    }
    public class CannotConvertExpression : ProblemReport
    {
        private readonly Fingerprint _f;
        private readonly string _w;
        public CannotConvertExpression(Fingerprint f, string worksheet)
        {
            _f = f;
            _w = worksheet;
        }
        public override string ToString()
        {
            return "Cannot convert expression with fingerprint " + _f.ToString() + " on worksheet " + _w;
        }
    }

    public class Extract
    {
        private static ExpressionTools.EData inlineExpression(AST.Address addr, Depends.DAG graph, MemoDBOpt memodb)
        {
            // get top-level AST
            var ast = graph.getASTofFormulaAt(addr);

            // merge subtrees
            return ExpressionTools.flattenExpression(ast, graph, memodb);
        }

        public static string[] extractAll(Depends.DAG graph, Dictionary<AST.Address, string> formulas, List<ProblemReport> pr)
        {
            // init MemoDB
            var mdbo = MemoDBOpt.Some(new MemoDBO());

            // group formulas by 'function'
            var g = groupFunctions(graph, formulas);

            // allocate return array
            var fpcores = new List<string>();

            // produce a string for each fingerprinted 'function'
            foreach (var fkvp in g.Functions)
            {
                Fingerprint fingerprint = fkvp.Key; // the fingerprint that characterizes the 'function'
                var invocations = fkvp.Value;       // invocations of this function

                Dictionary<Source, ExpressionTools.EData> edatas = null;
                try
                {
                    // inline all formulas and index by invoking cell address
                    edatas = invocations.Select(references => {
                        var faddr = references[0].Tail;
                        return new Tuple<Source, ExpressionTools.EData>(faddr, inlineExpression(faddr, graph, mdbo));
                    }).ToDictionary(kvp => kvp.Item1, kvp => kvp.Item2);
                } catch (Exception)
                {
                    // TODO: fix occasional out-of-bounds error for invocations_for_addr[0]
                    pr.Add(new NoReferencesForInvocation(fingerprint, graph.getWorkbookName()));
                    continue;
                }

                PreList prelist = null;
                try
                {
                    // generate prelist
                    prelist = makePreList(g, fingerprint, edatas, graph);
                } catch (Exception)
                {
                    // sometimes aliasing bugs pop up; if that happens, move on
                    pr.Add(new AliasingDetected(fingerprint, graph.getWorkbookName()));
                    continue;
                }

                // convert an arbitrary instance of this 'function'
                try
                {
                    var fpcore = convertToFPCore(g, fingerprint, edatas, graph, prelist);
                    fpcores.Add(fpcore.ToExpr(0));
                }
                catch (Exception)
                {
                    // we can't convert everything; give up
                    pr.Add(new CannotConvertExpression(fingerprint, graph.getWorkbookName()));
                }
            }

            return fpcores.ToArray();
        }

        private static Group groupFunctions(Depends.DAG graph, Dictionary<AST.Address, string> formulas)
        {
            var functions = new Functions();
            var variables = new Variables();
            var provenance = new Provenance();

            // group formulas and invocations by resultant
            foreach (var f in formulas)
            {
                var addr = f.Key;

                // return all references
                var refs = ExpressionTools.transitiveRefs(addr, graph);

                // compute resultant
                var res = Vector.run(addr, graph);

                // save each 'function'
                if (!functions.ContainsKey(res))
                {
                    functions.Add(res, new List<ExpressionTools.Vector[]>());
                }
                functions[res].Add(refs);

                // assign a fresh variable to each unique reference
                if (!variables.ContainsKey(res))
                {
                    // never seen this function before
                    var vm = new VariableMaker();
                    variables.Add(res, new Dictionary<Countable, string>());

                    // we only want unique path resultants for data
                    var resultants =
                        refs
                            .Where(r => !graph.isFormula(r.Head))
                            .Select(r => ExceLint.Vector.RelativeVector(r.Head, r.Tail, graph))
                            .Distinct();

                    // assign variable
                    foreach (var r in resultants)
                    {
                        if (!variables[res].ContainsKey(r))
                        {
                            variables[res].Add(r, vm.nextVariable());
                        }
                    }
                }

                // track provenance
                if (!provenance.ContainsKey(res))
                {
                    var ps = new List<AST.Address>();
                    provenance.Add(res, ps);
                }
                provenance[res].Add(addr);
            }

            return new Group(functions, variables, provenance);
        }

        private static PreList makePreList(Group g, Fingerprint fingerprint, Dictionary<Source, ExpressionTools.EData> edatas, Depends.DAG graph)
        {
            // produce prelist
            var prelist = new PreList();
            foreach (var f_invocations in g.Functions[fingerprint])
            {
                var bindings = new Dictionary<string, double>();
                foreach (var vectref in f_invocations)
                {
                    var ref_resultant = ExceLint.Vector.RelativeVector(vectref.Head, vectref.Tail, graph);

                    // not every vector gets a variable (e.g., vectors that point to functions);
                    // if so, move on
                    if (g.Variables[fingerprint].ContainsKey(ref_resultant))
                    {
                        var variable = g.Variables[fingerprint][ref_resultant];

                        // only bind values to variables if a reference refers
                        // to data, not a formula
                        if (edatas[vectref.Tail].Data.ContainsKey(vectref.Head))
                        {
                            var value = edatas[vectref.Tail].Data[vectref.Head];

                            // have we already bound a variable to this value?
                            if (bindings.ContainsKey(variable))
                            {
                                // it had better be the same value, dude
                                if (bindings[variable] != value)
                                {
                                    throw new Exception("Same reference bound to different value!");
                                } else
                                {
                                    continue;
                                }
                            }

                            bindings.Add(variable, value);
                        }
                    }
                }
                prelist.Add(bindings);
            }
            return prelist;
        }

        private static FPCoreAST.FPCore convertToFPCore(
            Group g,
            Fingerprint fingerprint,
            Dictionary<Source, ExpressionTools.EData> edatas,
            Depends.DAG graph,
            PreList prelist)
        {
            // get invocations for this function
            var f_invocations = g.Functions[fingerprint];
            var invocation = f_invocations.First();
            var faddr = invocation.First().Tail;

            // get bindings for this invocation
            var bindings = new Dictionary<AST.Address, string>();
            foreach (var vectref in invocation)
            {
                var ref_resultant = ExceLint.Vector.RelativeVector(vectref.Head, vectref.Tail, graph);

                // not every vector gets a variable (e.g., references to functions)
                if (g.Variables[fingerprint].ContainsKey(ref_resultant))
                {
                    var variable = g.Variables[fingerprint][ref_resultant];
                    if (!bindings.ContainsKey(vectref.Head))
                    {
                        bindings.Add(vectref.Head, variable);
                    }
                }
            }

            return XL2FPCore.FormulaToFPCore(edatas[faddr].Expression, prelist, bindings, g.Provenance[fingerprint].ToArray());
        }
    }
}
