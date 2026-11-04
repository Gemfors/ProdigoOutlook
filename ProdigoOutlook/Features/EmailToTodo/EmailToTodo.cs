namespace ProdigoOutlook.Features.EmailToTodo
{
    /// <summary>
    /// PHASE 4 - planned, not built. Email to Todo: a task in Todo Kanban from the
    /// selected or open email - subject as the title, a way back to the email, an
    /// optional due date.
    ///
    /// The integration point is still to be decided with Olle. What is known (November
    /// 2026): Todo Kanban is an Electron app keeping its data in
    /// %APPDATA%\todo-kanban\kanban-data.json and in memory, so writing to that file
    /// directly would be overwritten by the running app. The options discussed were a
    /// drop folder the app imports from (%APPDATA%\todo-kanban\inbox\*.json, works when
    /// the app is closed) and a todokanban://add link the app registers. Either needs a
    /// change in Todo Kanban itself.
    /// </summary>
    internal static class EmailToTodo
    {
    }
}
