module ExpressionTools

open Depends
open System.Collections.Generic

exception FlattenOperationNotSupportedException of string

let rec flattenedExpression(expr: AST.Expression)(graph: DAG) : AST.Expression =
    match expr with
    | AST.ReferenceExpr(r) -> flattenedRef r graph
    | AST.BinOpExpr(op, e1, e2) -> AST.BinOpExpr(op, flattenedExpression e1 graph, flattenedExpression e2 graph)
    | AST.UnaryOpExpr(op, e) -> AST.UnaryOpExpr(op, flattenedExpression e graph)
    | AST.ParensExpr(e) -> AST.ParensExpr(flattenedExpression e graph)

and flattenedRef(ref: AST.Reference)(graph: DAG) : AST.Expression =
    match ref with
    | :? AST.ReferenceRange as r ->
        let addrs = r.Range.Addresses()
        let asts = addrs |> Array.map (fun addr -> flattenedAddr addr graph) |> Array.toList
        let env = new AST.Env(addrs.[0].Path, addrs.[0].WorkbookName, addrs.[0].WorksheetName)
        let union = AST.ReferenceUnion(env, asts)
        // recursively inline
        flattenedExpression (AST.ReferenceExpr(union)) graph
    | :? AST.ReferenceAddress as r ->
        let faddr = flattenedAddr r.Address graph
        // recursively inline
        flattenedExpression faddr graph
    | :? AST.ReferenceNamed -> raise (FlattenOperationNotSupportedException "Named references not yet supported.")
    | :? AST.ReferenceFunction as r ->
        let args = r.ArgumentList
        let argexprs = args |> List.map (fun arg -> flattenedExpression arg graph)
        let env = new AST.Env(r.Path, r.WorkbookName, r.WorksheetName)
        let fn = AST.ReferenceFunction(env, r.FunctionName, argexprs, r.Arity)
        AST.ReferenceExpr(fn)
    | :? AST.ReferenceConstant as c -> AST.ReferenceExpr(c)
    | :? AST.ReferenceString as s -> AST.ReferenceExpr(s) 
    | :? AST.ReferenceBoolean as b -> AST.ReferenceExpr(b)
    | :? AST.ReferenceUnion as ru ->
        let refs = ru.References
        let asts = refs |> List.map (fun expr -> flattenedExpression expr graph)
        let union = AST.ReferenceUnion(ru.Environment, asts)
        AST.ReferenceExpr(union)
    | _ -> raise (FlattenOperationNotSupportedException "Unknown reference type.")

// try to get AST; if that doesn't work, get the value
// try to parse the value as a double; if that doesn't work, return it as a string
and flattenedAddr(addr: AST.Address)(graph: DAG) : AST.Expression =
    try
        graph.getASTofFormulaAt(addr)
    with
    | :? KeyNotFoundException as e ->
        let value = graph.readCOMValueAtAddress addr
        let env = new AST.Env(addr.Path, addr.WorkbookName, addr.WorksheetName)
        try
            let d = System.Convert.ToDouble(value)
            AST.ReferenceExpr(AST.ReferenceConstant(env, d))
        with
        | _ ->
            AST.ReferenceExpr(AST.ReferenceString(env, value))