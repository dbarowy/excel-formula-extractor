module ExpressionTools

open Depends
open System.Collections.Generic

exception FlattenOperationNotSupportedException of string

let join (ms: Map<'a,'b> list) : Map<'a,'b> =
    ms
    |> List.map Map.toSeq
    |> Seq.concat
    |> fun ms' -> Map(ms')

let adict(a: seq<('a*'b)>) = new Dictionary<'a,'b>(a |> dict)

type FExpression(expr: AST.Expression, data: Dictionary<AST.Address,double>) =
    member self.Expression = expr
    member self.Data = data

let rec private flattenedExpression(expr: AST.Expression)(graph: DAG) : AST.Expression*Map<AST.Address,double> =
    match expr with
    | AST.ReferenceExpr(r) -> flattenedRef r graph
    | AST.BinOpExpr(op, e1, e2) ->
        let e1flat,var1 = flattenedExpression e1 graph
        let e2flat,var2 = flattenedExpression e2 graph
        AST.BinOpExpr(op, e1flat, e2flat), join [var1; var2]
    | AST.UnaryOpExpr(op, e) ->
        let eflat,var = flattenedExpression e graph
        AST.UnaryOpExpr(op, eflat), var
    | AST.ParensExpr(e) ->
        let eflat, var = flattenedExpression e graph
        AST.ParensExpr(eflat), var

and private flattenedRef(ref: AST.Reference)(graph: DAG) : AST.Expression*Map<AST.Address,double> =
    match ref with
    | :? AST.ReferenceRange as r ->
        let addrs = r.Range.Addresses()
        let asts,varss = addrs |> Array.map (fun addr -> flattenedAddr addr graph) |> Array.toList |> List.unzip
        let env = new AST.Env(addrs.[0].Path, addrs.[0].WorkbookName, addrs.[0].WorksheetName)
        let union = AST.ReferenceUnion(env, asts)
        // recursively inline
        let expr,vars' = flattenedExpression (AST.ReferenceExpr(union)) graph
        // merge variable maps and return
        let vars = join varss
        expr, join [vars; vars']
    | :? AST.ReferenceAddress as r ->
        flattenedAddr r.Address graph
    | :? AST.ReferenceNamed -> raise (FlattenOperationNotSupportedException "Named references not yet supported.")
    | :? AST.ReferenceFunction as r ->
        let args = r.ArgumentList
        let argexprs, vars = args |> List.map (fun arg -> flattenedExpression arg graph) |> List.unzip
        let env = new AST.Env(r.Path, r.WorkbookName, r.WorksheetName)
        let fn = AST.ReferenceFunction(env, r.FunctionName, argexprs, r.Arity)
        // merge variable maps and return
        AST.ReferenceExpr(fn), join vars
    | :? AST.ReferenceConstant as c -> AST.ReferenceExpr(c), Map.empty
    | :? AST.ReferenceString as s -> AST.ReferenceExpr(s), Map.empty
    | :? AST.ReferenceBoolean as b -> AST.ReferenceExpr(b), Map.empty
    | :? AST.ReferenceUnion as ru ->
        let refs = ru.References
        let asts,vars = refs |> List.map (fun expr -> flattenedExpression expr graph) |> List.unzip
        let union = AST.ReferenceUnion(ru.Environment, asts)
        // merge variable maps and return
        AST.ReferenceExpr(union), join vars
    | _ -> raise (FlattenOperationNotSupportedException "Unknown reference type.")

// try to get AST; if that doesn't work, get the value
// try to parse the value as a double; if that doesn't work, return it as a string
and private flattenedAddr(addr: AST.Address)(graph: DAG) : AST.Expression*Map<AST.Address,double> =
    try
        // follow the reference
        let ast = graph.getASTofFormulaAt(addr)
        flattenedExpression ast graph
    with
    | :? KeyNotFoundException as e ->
        let value = graph.readCOMValueAtAddress addr
        let env = new AST.Env(addr.Path, addr.WorkbookName, addr.WorksheetName)
        try
            let d = System.Convert.ToDouble(value)
            let expr = AST.ReferenceExpr(AST.ReferenceAddress(env, addr))
            expr,Map([(addr,d)])
        with
        | _ ->
            AST.ReferenceExpr(AST.ReferenceString(env, value)),Map.empty

let flattenExpression(expr: AST.Expression)(graph: DAG) : FExpression =
    let expr, var = flattenedExpression expr graph
    let d = Map.toSeq var |> adict
    FExpression(expr, d)