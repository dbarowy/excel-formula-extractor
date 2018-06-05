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
        let asts = addrs |> Array.map (fun addr -> flattenedAddr addr graph)
        let union = asts |> Array.reduce (fun acc ast -> AST.BinOpExpr(",", acc, ast))
        union
    | :? AST.ReferenceAddress as r -> flattenedAddr r.Address graph
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