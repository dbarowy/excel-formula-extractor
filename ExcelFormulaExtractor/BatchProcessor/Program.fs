open COMWrapper
open System
open System.IO
open System.Collections.Generic
open Depends
open ExtractionLogic

type Dict<'a,'b> = Dictionary<'a,'b>

let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)

let getAllFormulas (graph: DAG) : Dict<AST.Address,string> =
    let frms = graph.getAllFormulaAddrs()
    frms |> Array.map (fun addr -> (addr, graph.getFormulaAtAddress addr)) |> adict

[<EntryPoint>]
let main argv = 
    let dir = argv.[0]
    let output = argv.[1]
    let prlog = argv.[2]

    Console.CancelKeyPress.Add(
        (fun _ ->
            printfn "Ctrl-C received.  Cancelling..."
            System.Environment.Exit 1
        )
    )

    using(new Application()) (fun app ->
        let files = Directory.EnumerateFiles(dir, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray

        for file in files do
            let shortf = (System.IO.Path.GetFileName file)

            printfn "Opening: %A" shortf
            using(app.OpenWorkbook(file)) (fun wb ->
                let problems = new System.Collections.Generic.List<ExtractionLogic.ProblemReport>()

                printfn "Building dependence graph: %A" shortf
                let graph = wb.buildDependenceGraph()

                printfn "Getting all formulas: %A" shortf
                let formulas = getAllFormulas graph

                printfn "Converting to FPCores: %A" shortf
                let fpcores = ExtractionLogic.Extract.extractAll(graph, formulas, problems)

                printfn "Writing to output file: %A" output
                File.AppendAllLines(output, fpcores)

                printfn "Writing to problem report log: %A" prlog
                File.AppendAllLines(prlog, problems |> Seq.map (fun pr -> pr.ToString()))
            )
    )

    0