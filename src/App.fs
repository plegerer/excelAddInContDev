module App

open System
open Elmish
open Elmish.React
open Feliz
open ExcelJS.Fable.GlobalBindings
open ExcelJS.Fable.Excel
open Thoth.Elmish
open System



let initializeAddIn () = Office.onReady ()



type State = { Count: int; Excelstate: string }

type Msg =
    | Increment
    | Decrement
    | OnPromiseSuccess of string * string
    | OnPromiseError of exn
    | UpdateMsg
    | OfficeInitialized of string * string
    | EventRegistered



let mapEvent = Event<Msg>()

let mapEventSubscription initial =
    let sub dispatch =
        let msgSender msg = 
            msg
            |> dispatch
            
        mapEvent.Publish.Add(msgSender)

    Cmd.ofSub sub






let handleSelectionChange (event:WorksheetSelectionChangedEventArgs) =
    
    Excel.run (fun context -> 
                    let address = event.address
                    context
                        .sync().``then``(fun x -> 
                                            let constext= "Event fired, address is: " + address
                                            Console.WriteLine constext
                                            mapEvent.Trigger (OnPromiseSuccess ("Address is: ", address) )
                                            Some x))
    

                        


let registerEvent() =
    Excel.run (fun context ->

        let worksheet =
            context.workbook.worksheets.getActiveWorksheet ()
        
        let eventResult = worksheet.onSelectionChanged.add(handleSelectionChange)
        context.sync())



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

                ))













let init () =
    let initialCmd =
        Cmd.OfPromise.perform initializeAddIn () 
            (fun x ->
                (x.host.ToString(), x.platform.ToString())|> OfficeInitialized)
    let registerEventCmd =
        Cmd.OfPromise.either registerEvent () (fun x ->
                ("eventhandler", " registered")
                |> OnPromiseSuccess) 
            (fun e -> OnPromiseError e)
    { Count = 0; Excelstate = "" }, Cmd.batch [initialCmd]

let update (msg: Msg) (state: State) =
    match msg with
    | OfficeInitialized (x,y) -> state , Cmd.ofMsg EventRegistered
    | EventRegistered -> 
        state, Cmd.OfPromise.either registerEvent () (fun x ->
                ("eventhandler", " registered")
                |> OnPromiseSuccess) 
                (fun e -> OnPromiseError e)


    | Increment -> { state with Count = state.Count + 3 }, Cmd.ofMsg UpdateMsg
    | Decrement -> { state with Count = state.Count - 1 }, Cmd.ofMsg UpdateMsg
    | OnPromiseSuccess (x, y) ->
       

        { state with Excelstate = x + " : " + y },
        Toast.message y
        |> Toast.position Toast.BottomCenter
        |> Toast.timeout (TimeSpan.FromSeconds(3.0))
        |> Toast.success
    | OnPromiseError e ->
        printfn "%A" e
        state, Cmd.none
    | UpdateMsg ->
        let cmd =
            Cmd.OfPromise.perform UpdateValue (state.Count) (fun x ->
                ("neuer Wert", string state.Count)
                |> OnPromiseSuccess)

        state, cmd



let render (state: State) (dispatch: Msg -> unit) =
    Html.div [ Html.button [ prop.onClick (fun _ -> dispatch Increment)
                             prop.text "Increment" ]

               Html.button [ prop.onClick (fun _ -> dispatch Decrement)
                             prop.text "Decrement" ]

               Html.h1 state.Count

               Html.p state.Excelstate ]

Program.mkProgram init update render
|> Program.withSubscription mapEventSubscription
|> Toast.Program.withToast Toast.render
|> Program.withReactSynchronous "elmish-app"
|> Program.run
