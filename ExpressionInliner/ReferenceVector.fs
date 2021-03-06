﻿namespace ExceLint
    open Depends
    open System
    open System.Collections.Generic
    open Utils

    module public Vector =
        type public Directory = string
        type public WorkbookName = string
        type public WorksheetName = string
        type public Path = Directory*WorkbookName*WorksheetName
        type public X = int    // i.e., column displacement
        type public Y = int    // i.e., row displacement
        type public Z = int    // i.e., worksheet displacement (0 if same sheet, 1 if different)
        type public C = double // i.e., a constant value

        type ArityZero() =
            static let idx: Dictionary<string,int> =
                let mutable i = 1
                Grammar.Arity0Names |>
                Array.sort |>
                Array.fold (fun (acc: Dictionary<string,int>)(n: string) ->
                    acc.Add(n, -i)
                    i <- i + 1
                    acc
                ) (new Dictionary<string,int>()) 
            static member isZeroArity n = idx.ContainsKey n
            static member hasIndex n = idx.[n]

        // components for mixed vectors
        type public VectorComponent =
        | Abs of int
        | Rel of int
            override self.ToString() : string =
                match self with
                | Abs(i) -> "Abs(" + i.ToString() + ")"
                | Rel(i) -> "Rel(" + i.ToString() + ")"

        // the vector, relative to an origin
        type public Coordinates = (X*Y*Path)
        type public RelativeVector =
        | NoConstant of X*Y*Z
        | NoConstantWithLoc of X*Y*Z*X*Y*Z
        | Constant of X*Y*Z*C
        | ConstantWithLoc of X*Y*Z*X*Y*Z*C
            member self.Zero =
                match self with
                | Constant(_,_,_,_) -> Constant(0,0,0,0.0)
                | ConstantWithLoc(x,y,z,_,_,_,_) -> ConstantWithLoc(x,y,z,0,0,0,0.0)
                | NoConstant(_,_,_) -> NoConstant(0,0,0)
                | NoConstantWithLoc(x,y,z,_,_,_) -> NoConstantWithLoc(x,y,z,0,0,0)
        type public MixedVector = (VectorComponent*VectorComponent*Path)
        type public MixedVectorWithConstant = (VectorComponent*VectorComponent*Path*C)
        type public SquareVector(dx: double, dy: double, x: double, y: double) =
            member self.dx = dx
            member self.dy = dy
            member self.x = x
            member self.y = y
            member self.AsArray = [| dx; dy; x; y |]
            override self.Equals(o: obj) : bool =
                match o with
                | :? SquareVector as o' ->
                    (dx, dy, x, y) = (o'.dx, o'.dy, o'.x, o'.y)
                | _ -> false
            override self.GetHashCode() = hash (dx, dy, x, y)
            override self.ToString() =
                "(" +
                dx.ToString() + "," +
                dy.ToString() + "," +
                x.ToString() + "," +
                y.ToString() +
                ")"

        // handy datastructures
        type public Edge = SquareVector*SquareVector
        type public DistDict = Dictionary<Edge,double>

        // the first component is the tail (start) and the second is the head (end)
        type public RichVector =
        | MixedFQVectorWithConstant of Coordinates*MixedVectorWithConstant
        | MixedFQVector of Coordinates*MixedVector
        | AbsoluteFQVector of Coordinates*Coordinates
            override self.ToString() : string =
                match self with
                | MixedFQVectorWithConstant(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | MixedFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | AbsoluteFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
        type private KeepConstantValue =
        | Yes
        | No

        type private VectorMaker = AST.Address -> AST.Address -> RichVector
        type private ConstantVectorMaker = AST.Address -> AST.Expression -> RichVector list
        type private Rebaser = RichVector -> DAG -> bool -> bool -> RelativeVector

        let private fullPath(addr: AST.Address) : string*string*string =
            // portably create full path from components
            (addr.Path, addr.WorkbookName, addr.WorksheetName)

        let private vector(tail: AST.Address)(head: AST.Address)(mixed: bool)(include_constant: bool) : RichVector =
            let tailXYP = (tail.X, tail.Y, fullPath tail)
            if mixed then
                let X = match head.XMode with
                        | AST.AddressMode.Absolute -> Abs(head.X)
                        | AST.AddressMode.Relative -> Rel(head.X)
                let Y = match head.YMode with
                        | AST.AddressMode.Absolute -> Abs(head.Y)
                        | AST.AddressMode.Relative -> Rel(head.Y)
                
                if include_constant then
                    let headXYP = (X, Y, fullPath head, 0.0)
                    MixedFQVectorWithConstant(tailXYP, headXYP)
                else
                    let headXYP = (X, Y, fullPath head)
                    MixedFQVector(tailXYP, headXYP)
            else
                let headXYP = (head.X, head.Y, fullPath head)
                AbsoluteFQVector(tailXYP, headXYP)

        let private originPath(dag: DAG) : Path =
            (dag.getWorkbookDirectory(), dag.getWorkbookName(), dag.getWorksheetNames().[0]);

        let private vectorPathDiff(p2: Path)(p1: Path)(graph: Depends.DAG) : int =
            let p2pci = graph.getPathClosureIndex(p2)
            let p1pci = graph.getPathClosureIndex(p1)
            p2pci - p1pci

        // represent the position of the head of the vector relative to the tail, (x1,y1,z1)
        // if the reference is off-sheet then optionally ignore X and Y vector components
        let private relativeToTail(absVect: RichVector)(dag: DAG)(offSheetInsensitive: bool)(includeLoc: bool) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
                if offSheetInsensitive && p1 <> p2 then
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p2))
                else
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x2-x1, y2-y1, vectorPathDiff p2 p1 dag)
                    else
                        NoConstant(x2-x1, y2-y1, vectorPathDiff p2 p1 dag)
            | MixedFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
                let x' = match x2 with
                            | Rel(x) -> x - x1
                            | Abs(x) -> x
                let y' = match y2 with
                            | Rel(y) -> y - y1
                            | Abs(y) -> y
                if offSheetInsensitive && p1 <> p2 then
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p2))
                else
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x', y', vectorPathDiff p2 p1 dag)
                    else
                        NoConstant(x', y', vectorPathDiff p2 p1 dag)
            | MixedFQVectorWithConstant(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2,c) = head

                let x' = match x2 with
                            | Rel(x) -> x - x1
                            | Abs(x) -> x
                let y' = match y2 with
                            | Rel(y) -> y - y1
                            | Abs(y) -> y
                if offSheetInsensitive && p1 <> p2 then
                    if includeLoc then
                        ConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2), c)
                    else
                        Constant(0, 0, dag.getPathClosureIndex(p2), c)
                else
                    if includeLoc then
                        ConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x', y', vectorPathDiff p2 p1 dag, c)
                    else
                        Constant(x', y', vectorPathDiff p2 p1 dag, c)

        let private RVSum(v1: RelativeVector)(v2: RelativeVector) : RelativeVector =
            match v1,v2 with
            | NoConstant(x1,y1,z1), NoConstant(x2,y2,z2) ->
                NoConstant(x1 + x2, y1 + y2, z1 + z2)
            | NoConstantWithLoc(x1,y1,z1,dx1,dy1,dz1), NoConstantWithLoc(x2,y2,z2,dx2,dy2,dz2) ->
                assert (x1 = x2 && y1 = y2 && z1 = z2)
                // we don't add reference sources, just reference destinations
                NoConstantWithLoc(x1, y1, z1, dx1 + dx2, dy1 + dy2, dz1 + dz2)
            | Constant(x1,y1,z1,c1), Constant(x2,y2,z2,c2) ->
                Constant(x1 + x2, y1 + y2, z1 + z2, c1 + c2)
            | ConstantWithLoc(x1,y1,z1,dx1,dy1,dz1,dc1), ConstantWithLoc(x2,y2,z2,dx2,dy2,dz2,dc2) ->
                assert (x1 = x2 && y1 = y2 && z1 = z2)
                // we don't add reference sources, just reference destinations
                ConstantWithLoc(x1, y1, z1, dx1 + dx2, dy1 + dy2, dz1 + dz2, dc1 + dc2)
            | _ -> failwith "Cannot sum RelativeVectors of different subtypes."

        let private Resultant(vs: RelativeVector[]) : RelativeVector =
            vs |>
            Array.fold (fun (acc: RelativeVector option)(v: RelativeVector) ->
                match acc with
                | None -> Some (RVSum v.Zero v)
                | Some a -> Some (RVSum a v)
            ) None |>
            (fun rvopt ->
                match rvopt with
                | Some rv -> rv
                | None -> failwith "Empty resultant!"
            )

        let private SquareMatrix(origin: X*Y)(vs: RelativeVector[]) : X*Y*X*Y =
            let (x,y) = origin
            let xyoff = vs |>
                        Array.fold (fun (xacc: X, yacc: Y)(rv: RelativeVector) ->
                            let (x',y') =
                                match rv with
                                | Constant(x,y,_,_) -> x,y
                                | NoConstant(x,y,_) -> x,y
                                | _ -> failwith "not supported"
                            xacc + x', yacc + y'
                        ) (0,0)
            (fst xyoff, snd xyoff, x, y)

        let zeroArityXYP(op: string)(tailPath: string*string*string)(cvc: C) : MixedVectorWithConstant option =
            if ArityZero.isZeroArity op then
                Some (VectorComponent.Abs (ArityZero.hasIndex op), VectorComponent.Abs (ArityZero.hasIndex op), tailPath, cvc)
            else
                None

        let refsForArityZeroOps(tail: AST.Address)(ops: string list) : RichVector list =
            if ops.Length = 0 then
                []
            else
                let tailPath = fullPath tail
                let tailXYP = tail.X, tail.Y, tailPath
                let c = 0.0    // no constant
                let heads = ops |> List.map (fun op -> zeroArityXYP op tailPath c) |> List.choose id
                heads |> List.map (fun head -> MixedFQVectorWithConstant(tailXYP, head))

        let transitiveInputVectors(fCell: AST.Address)(dag : DAG)(depth: int option)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker)(tail_is_fcell: bool) : RichVector[] =
            let rec tfVect(tailO: AST.Address option)(head: AST.Address)(depth: int option) : RichVector list =
                let vlist = match tailO with
                            | Some tail -> if tail_is_fcell then [vector_f fCell head]  else [vector_f tail head]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tfVect_b head (Some(d-1)) vlist
                | None -> tfVect_b head None vlist

            and tfVect_b(tail: AST.Address)(nextDepth: int option)(vlist: RichVector list) : RichVector list =
                let root = if tail_is_fcell then fCell else tail

                if (dag.isFormula tail) then
                    try
                        // parse again, because the DAG treats repeated
                        // references to the same cell but with different
                        // address modes as the same reference; they are not.

                        // Sometimes idiots denote comments with '='.
                        let fexpr = Parcel.parseFormulaAtAddress tail (dag.getFormulaAtAddress tail)

                        // find all of the inputs for source
                        let heads_single = Parcel.addrReferencesFromExpr fexpr |> List.ofArray
                        let heads_vector = Parcel.rangeReferencesFromExpr fexpr |>
                                                List.ofArray |>
                                                List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                                List.concat

                        // find all constant inputs for source
                        let cvects = cvector_f root fexpr

                        // Get references for zero-arity functions
                        let ops = Parcel.operatorNamesFromExpr fexpr
                        let zvects = refsForArityZeroOps root ops

                        let heads = heads_single @ heads_vector
                        // recursively call this function
                        vlist @ cvects @ zvects @ (List.map (fun head -> tfVect (Some tail) head nextDepth) heads |> List.concat)
                    with
                    | e -> vlist  // I guess we give up
                else
                    let value = dag.readCOMValueAtAddress(tail)
                    let mutable num = 0.0
                    num <- if Double.TryParse(value, &num) then
                               // a constant, i.e., references one thing
                               1.0
                           else if String.IsNullOrWhiteSpace(value) then
                               // it's blank, i.e., references nothing
                               0.0
                           else
                               // it's a string
                               -1.0  // pretty arbitrary... maybe we should have a "blank" dimension
                    let env = AST.Env(tail.Path, tail.WorkbookName, tail.WorksheetName)
                    let expr = AST.ReferenceExpr (AST.ReferenceConstant(env, num))
                    let dv = cvector_f root expr
                    dv @ vlist
    
            tfVect None fCell depth |> List.toArray

        let private makeVector(isMixed: bool)(includeConstant: bool): VectorMaker =
            (fun (source: AST.Address)(sink: AST.Address) ->
                vector source sink isMixed includeConstant
            )

        let private nopConstantVector : ConstantVectorMaker =
            (fun (a: AST.Address)(e: AST.Expression) -> [])

        let private makeConstantVectorsFromConstants(k: KeepConstantValue) : ConstantVectorMaker =
            (fun (tail: AST.Address)(e: AST.Expression) ->
                // convert into RichVector form
                let tailXYP = (tail.X, tail.Y, fullPath tail)

                // the path for the head is the same as the path for the tail for constants
                let path = fullPath tail
                let constants = Parcel.constantsFromExpr e

                // the vectorcomponents for constants are Abs(0)
                let cvc = Abs(0)

                // make vectors
                let cf = match k with
                         | Yes -> (fun (rc: AST.ReferenceConstant) -> rc.Value)
                         | No -> (fun (rc: AST.ReferenceConstant) -> if rc.Value = 0.0 then 0.0 else if rc.Value = -1.0 then -1.0 else 1.0)

                let vs = Array.map (fun (c: AST.ReferenceConstant) ->
                             RichVector.MixedFQVectorWithConstant(tailXYP, (cvc, cvc, path, cf c))
                         ) constants |> Array.toList

                vs
            )

        let getVectors(cell: AST.Address)(dag: DAG)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker)(transitive: bool)(isForm: bool)(tail_is_formula: bool) : RichVector[] =
            let depth = if transitive then None else (Some 1)
            let output = transitiveInputVectors cell dag depth vector_f cvector_f tail_is_formula
            output

        let ResultantMaker(cell: AST.Address)(dag: DAG)(isMixed: bool)(includeConstant: bool)(includeLoc: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(constant_f: ConstantVectorMaker)(rebase_f: Rebaser)(tail_is_formula: bool) : Countable =
            let vs = getVectors cell dag (makeVector isMixed includeConstant) constant_f isTransitive isFormula tail_is_formula
            let rebased_vs = vs |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive includeLoc)
            let resultant = rebased_vs |> Resultant
            let countable =
                resultant
                |> (fun rv ->
                        match rv with
                        | Constant(x,y,z,c) -> CVectorResultant(double x, double y, double z, double c)
                        | NoConstant(x,y,z) -> Vector(double x, double y, double z)
                        | ConstantWithLoc(x,y,z,dx,dy,dz,dc) -> FullCVectorResultant(double x, double y, double z, double dx, double dy, double dz, double dc)
                        | NoConstantWithLoc(x,y,z,dx,dy,dz) -> Countable.SquareVector(double dx, double dy, double dz, double x, double y, double z)
                   )
            countable

        let RelativeVector(head: AST.Address)(tail: AST.Address)(dag: DAG) : Countable =
            // vector params
            let isMixed = true
            let isOffSheetInsensitive = false
            let includeConstant = true
            let includeLoc = false
            // get vector
            let v = (makeVector isMixed includeConstant) tail head
            // make relative to tail
            let rv = relativeToTail v dag isOffSheetInsensitive includeLoc
            // return as Countable
            match rv with
            | Constant(x,y,z,c) -> CVectorResultant(double x, double y, double z, double c)
            | NoConstant(x,y,z) -> Vector(double x, double y, double z)
            | ConstantWithLoc(x,y,z,dx,dy,dz,dc) -> FullCVectorResultant(double x, double y, double z, double dx, double dy, double dz, double dc)
            | NoConstantWithLoc(x,y,z,dx,dy,dz) -> Countable.SquareVector(double dx, double dy, double dz, double x, double y, double z)

        type Vector() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = true
                let isFormula = true
                let isOffSheetInsensitive = false
                let tailIsAlwaysSourceFormula = true
                let includeConstant = true
                let includeLoc = false
                let keepConstantValues = KeepConstantValue.No
                let rebase_f = relativeToTail
                let constant_f = makeConstantVectorsFromConstants keepConstantValues
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f tailIsAlwaysSourceFormula
            static member capability : string*Capability =
                (typeof<Vector>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = Vector.run } )