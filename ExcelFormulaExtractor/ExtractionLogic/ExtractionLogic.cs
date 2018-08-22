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

namespace ExtractionLogic
{
    public class Group
    {
        private readonly Functions _functions;
        private readonly Variables _variables;

        public Group(Functions f, Variables v)
        {
            _functions = f;
            _variables = v;
        }

        public Functions Functions
        {
            get { return _functions; }
        }

        public Variables Variables
        {
            get { return _variables; }
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

        public static string[] extractAll(Depends.DAG graph, Dictionary<AST.Address, string> formulas)
        {
            // init MemoDB
            var mdbo = MemoDBOpt.Some(new MemoDBO());

            // group formulas by 'function'
            var g = groupFunctions(graph, formulas);

            // allocate return array
            var fpcores = new string[g.Functions.Count];
            int i = 0;

            // produce a string for each fingerprinted 'function'
            foreach (var fkvp in g.Functions)
            {
                Fingerprint fingerprint = fkvp.Key;
                var vectors = fkvp.Value;  // vectors for this function's references

                // inline all formulas
                var edatas = vectors.Select(invocations_for_addr => {
                    var faddr = invocations_for_addr[0].Tail;
                    return new Tuple<Source, ExpressionTools.EData>(faddr, inlineExpression(faddr, graph, mdbo));
                }).ToDictionary(kvp => kvp.Item1, kvp => kvp.Item2);

                // generate prelist
                var prelist = makePreList(g, fingerprint, edatas, graph);

                // convert an arbitrary instance of this 'function'
                try
                {
                    var fpcore = convertToFPCore(g, fingerprint, edatas, graph, prelist);
                    fpcores[i] = fpcore.ToExpr(0);
                }
                catch (XL2FPCore.InvalidExpressionException)
                {
                    // we can't convert everything; give up
                }

                i++;

            }

            return fpcores;
        }


        private static Group groupFunctions(Depends.DAG graph, Dictionary<AST.Address, string> formulas)
        {
            var invocations = new Dictionary<Fingerprint, List<ExpressionTools.Vector[]>>();
            var refvars = new Dictionary<Fingerprint, Dictionary<Countable, string>>();

            

            // group formulas and invocations by resultant
            foreach (var f in formulas)
            {
                var addr = f.Key;

                // return all references
                var refs = ExpressionTools.transitiveRefs(addr, graph);

                // compute resultant
                var res = Vector.run(addr, graph);

                // save each 'invocation'
                if (!invocations.ContainsKey(res))
                {
                    invocations.Add(res, new List<ExpressionTools.Vector[]>());
                }
                invocations[res].Add(refs);

                // assign a fresh variable to each unique reference
                if (!refvars.ContainsKey(res))
                {
                    // never seen this function before
                    var vm = new VariableMaker();
                    refvars.Add(res, new Dictionary<Countable, string>());

                    foreach (var vectref in refs)
                    {
                        var ref_resultant = ExceLint.Vector.RelativeVector(vectref.Head, vectref.Tail, graph);
                        if (!refvars[res].ContainsKey(ref_resultant))
                        {
                            refvars[res].Add(ref_resultant, vm.nextVariable());
                        }
                    }
                }
            }

            return new Group(invocations, refvars);
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
                            System.Diagnostics.Debug.Assert(bindings[variable] == value);
                            continue;
                        }

                        bindings.Add(variable, value);
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
                var variable = g.Variables[fingerprint][ref_resultant];
                if (!bindings.ContainsKey(vectref.Head))
                {
                    bindings.Add(vectref.Head, variable);
                }
            }

            return XL2FPCore.FormulaToFPCore(edatas[faddr].Expression, prelist, bindings);
        }
    }
}
