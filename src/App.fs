module App

open Elmish
open Elmish.React
open Feliz
open ExcelJS.Fable.GlobalBindings


let initializeAddIn () = Office.onReady ()



type State = { Count: int; Excelstate: string }

type Msg =
    | Increment
    | Decrement
    | OnPromiseSuccess of string * string
    | OnPromiseError of exn
    | UpdateMsg


let UpdateValue x =
    Excel.run (fun context ->

        let worksheet =
            context.workbook.worksheets.getActiveWorksheet ()

        let range = worksheet.getCell (1., 1.)

        context
            .sync()
            .``then`` (fun _ ->
                let customValue =
                    ResizeArray [| ResizeArray [| Some(x) |] |]

                range.values <- customValue

                if Office.context.requirements.isSetSupported ("ExcelApi", "1.2") then
                    range.format.autofitColumns ()))













let init () =
    let initialCmd =
        Cmd.OfPromise.perform Office.onReady () (fun x ->
            (x.host.ToString(), x.platform.ToString())
            |> OnPromiseSuccess)

    { Count = 0; Excelstate = "" }, initialCmd

let update (msg: Msg) (state: State) =
    match msg with
    | Increment -> { state with Count = state.Count + 3 }, Cmd.ofMsg UpdateMsg
    | Decrement -> { state with Count = state.Count - 1 }, Cmd.ofMsg UpdateMsg
    | OnPromiseSuccess (x, y) ->
        printfn "%A" x
        { state with Excelstate = x + " : " + y }, Cmd.none
    | OnPromiseError e ->
        printfn "%A" e
        state, Cmd.none
    | UpdateMsg ->
        let cmd =
            Cmd.OfPromise.perform UpdateValue () (fun x -> ("neuer Wert", "Test") |> OnPromiseSuccess)

        state, cmd

let render (state: State) (dispatch: Msg -> unit) =
    Html.div [ Html.button [ prop.onClick (fun _ -> dispatch Increment)
                             prop.text "Increment" ]

               Html.button [ prop.onClick (fun _ -> dispatch Decrement)
                             prop.text "Decrement" ]

               Html.h1 state.Count

               Html.p state.Excelstate ]

Program.mkProgram init update render
|> Program.withReactSynchronous "elmish-app"
|> Program.run
