module ExpressionTools

open Depends
open System.Collections.Generic

exception FlattenOperationNotSupportedException of string

let AddrToKey(a: AST.Address) : AST.Address*bool*bool =
    a, a.ColMode = AST.AddressMode.Absolute, a.RowMode = AST.AddressMode.Absolute

let join (ms: Map<'a,'b> list) : Map<'a,'b> =
    ms
    |> List.map Map.toSeq
    |> Seq.concat
    |> fun ms' -> Map(ms')

let adict(a: seq<('a*'b)>) = new Dictionary<'a,'b>(a |> dict)

type FExpression(expr: AST.Expression, data: Dictionary<AST.Address,double>) =
    member self.Expression = expr
    member self.Data = data

type EData(fexpr: FExpression, cache_hits: int) =
    member self.Expression = fexpr.Expression
    member self.Data = fexpr.Data
    member self.CacheHits = cache_hits

type Vector(tail: AST.Address, head: AST.Address) =
    member self.Tail = tail
    member self.Head = head
    override self.ToString() =
        tail.ToString() + " -> " + head.ToString()

// first element: the converted expression
// second element: the set of references closed over in the expression
// third element: cache hits
type FormulaData = AST.Expression*Map<AST.Address,double>
type MemoDB = Dictionary<AST.Address*bool*bool,FormulaData>

// this method returns a tuple
let rec private flattenedExpression(expr: AST.Expression)(graph: DAG)(dbo: MemoDB option) : FormulaData*int =
    match expr with
    | AST.ReferenceExpr(r) -> flattenedRef r graph dbo
    | AST.BinOpExpr(op, e1, e2) ->
        let (e1flat,var1),ch1 = flattenedExpression e1 graph dbo
        let (e2flat,var2),ch2 = flattenedExpression e2 graph dbo
        (AST.BinOpExpr(op, e1flat, e2flat), join [var1; var2]), ch1+ch2
    | AST.UnaryOpExpr(op, e) ->
        let (eflat,var),ch = flattenedExpression e graph dbo
        (AST.UnaryOpExpr(op, eflat), var), ch
    | AST.ParensExpr(e) ->
        let (eflat, var), ch = flattenedExpression e graph dbo
        (AST.ParensExpr(eflat), var), ch

and private flattenedRef(ref: AST.Reference)(graph: DAG)(dbo: MemoDB option) : FormulaData*int =
    match ref with
    | :? AST.ReferenceRange as r ->
        // flatten the formulas referenced in the range
        let addrs = r.Range.Addresses()
        let fd,chs = addrs |> Array.map (fun addr -> flattenedAddr addr graph dbo) |> Array.toList |> List.unzip
        let (asts,varss) = List.unzip fd
        let env = new AST.Env(addrs.[0].Path, addrs.[0].WorkbookName, addrs.[0].WorksheetName)
        let union = AST.ReferenceUnion(env, asts)
        let ch = List.sum chs
        // recursively inline
        let (expr,vars'),ch' = flattenedExpression (AST.ReferenceExpr(union)) graph dbo
        // merge variable maps and return
        let vars = join varss
        (expr, join [vars; vars']),ch+ch'
    | :? AST.ReferenceAddress as r ->
        flattenedAddr r.Address graph dbo
    | :? AST.ReferenceNamed -> raise (FlattenOperationNotSupportedException "Named references not yet supported.")
    | :? AST.ReferenceFunction as r ->
        let args = r.ArgumentList
        let fd, chs = args |> List.map (fun arg -> flattenedExpression arg graph dbo) |> List.unzip
        let (argexprs, vars) = List.unzip fd
        let env = new AST.Env(r.Path, r.WorkbookName, r.WorksheetName)
        let fn = AST.ReferenceFunction(env, r.FunctionName, argexprs, r.Arity)
        // merge variable maps and return
        (AST.ReferenceExpr(fn), join vars), List.sum chs
    | :? AST.ReferenceConstant as c -> (AST.ReferenceExpr(c), Map.empty), 0
    | :? AST.ReferenceString as s -> (AST.ReferenceExpr(s), Map.empty), 0
    | :? AST.ReferenceBoolean as b -> (AST.ReferenceExpr(b), Map.empty), 0
    | :? AST.ReferenceUnion as ru ->
        let refs = ru.References
        let fd,chs = refs |> List.map (fun expr -> flattenedExpression expr graph dbo) |> List.unzip
        let (asts,vars) = List.unzip fd
        let union = AST.ReferenceUnion(ru.Environment, asts)
        // merge variable maps and return
        (AST.ReferenceExpr(union), join vars), List.sum chs
    | _ -> raise (FlattenOperationNotSupportedException "Unknown reference type.")

// try to get AST; if that doesn't work, get the value
// try to parse the value as a double; if that doesn't work, return it as a string
and private flattenedAddr(addr: AST.Address)(graph: DAG)(dbo: MemoDB option) : FormulaData*int =
    // follow the reference
    try
        // check the cache first
        match dbo with
        | Some db ->
            if db.ContainsKey (AddrToKey addr) then
                let res = db.[(AddrToKey addr)]
                res, 1
            else    // not in cache; save
                let ast = graph.getASTofFormulaAt(addr)
                let asti, ch = flattenedExpression ast graph dbo
                db.Add((AddrToKey addr), asti)
                asti, ch
          // no cache
        | None ->
            let ast = graph.getASTofFormulaAt(addr)
            flattenedExpression ast graph dbo
    with
    | :? KeyNotFoundException as e ->
        // values at not in graph; need to read them
        let value = graph.readCOMValueAtAddress addr
        let env = new AST.Env(addr.Path, addr.WorkbookName, addr.WorksheetName)
        let fd =
            try
                let d = System.Convert.ToDouble(value)
                let expr = AST.ReferenceExpr(AST.ReferenceAddress(env, addr))
                expr,Map([(addr,d)])
            with
            | _ ->
                AST.ReferenceExpr(AST.ReferenceString(env, value)),Map.empty
        // save in cache
        match dbo with
        | Some db ->
            db.Add((AddrToKey addr), fd)
        | None -> ()
        fd,0

let flattenExpression(expr: AST.Expression)(graph: DAG)(dbo: MemoDB option) : EData =
    let (expr, var), ch = flattenedExpression expr graph dbo
    let d = Map.toSeq var |> adict
    EData(FExpression(expr, d),ch)

let transitiveRefs(faddr: AST.Address)(graph: DAG) : Vector[] =
    let rec tr(addr: AST.Address) =
        let expr = graph.getASTofFormulaAt addr
        let addrs = Parcel.addrReferencesFromExpr expr |> Array.map (fun a -> Vector(faddr, a))
        let rngs = Parcel.rangeReferencesFromExpr expr
        let addrs' =
            Array.concat
                [
                    addrs;
                    (rngs
                     |> Array.map (fun rng ->
                        rng.Addresses()
                        |> Array.map (fun a ->
                            Vector(faddr, a)
                           )
                        )
                     |> Array.concat)
                ]
                |> Array.distinct
        let follow =
            addrs'
            |> Array.map (fun v ->
                if graph.isFormula v.Head then
                    Some (tr v.Head)
                else
                    None
               )
            |> Array.choose id
            |> Array.concat
        Array.concat [addrs'; follow]
    tr faddr
    