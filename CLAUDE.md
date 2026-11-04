# Prodigo for Outlook — working notes

VSTO Outlook (classic) add-in. C#, .NET Framework 4.8, C# 7.3, assembly name
`ProdigoOutlook`. Third member of the family: `..\Prodigo` (PowerPoint, VSTO) and
`..\ProdigoExcel` (Excel-DNA). Started 4 November 2026. Phase 1 (Send later) is built;
phases 2–8 are stubbed. See README for the roadmap.

## Who Olle is, and what that means

Olle is not a programmer. He builds these tools in his spare time and is Group
Strategy Manager at HiQ by day.

- Explain decisions in plain language, with what he'll actually see, and take
  decisions one at a time with a recommendation.
- Don't use unexplained API names in anything addressed to him.
- He asks sharp questions; treat him as capable.

## Build, install, test

- **Build:** MSBuild works in a Claude session on Olle's machine:
  `C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\MSBuild.exe ProdigoOutlook.sln`.
  Building **registers the add-in** in
  `HKCU\Software\Microsoft\Office\Outlook\Addins\ProdigoOutlook` pointing at
  `bin\Debug`. So the next Outlook start runs whatever was last built in Olle's real
  mailbox. That's the agreed install method, the same as the other two add-ins.
- **Claude's own shell is sandboxed, and its registry writes don't reach the real
  registry.** A build there compiles fine, and its registration shows up when read
  back from the same shell, but Outlook never sees it. This cost a debugging round on
  4 November 2026.
  - To register for real, run the build in Olle's Terminal panel, outside the
    sandbox, with `/t:Rebuild`.
  - Check with `reg query HKCU\Software\Microsoft\Office\Outlook\Addins` run there too.
  - The same applies to anything else that must reach the real registry.
- **Visual Studio skips registration when the build is up to date.** If the tab is
  missing after a build, use Build → *Rebuild Solution*.
- **Not an installer.** Olle tried one before and hit certificate problems. Don't
  push MSI or ClickOnce.
- **Outlook locks the DLL** while running, including when it's only in the tray. The
  build fails with a file-in-use error until Outlook is fully quit.
- **Signing:** the manifests are signed with the PowerPoint project's test
  certificate (thumbprint `FA385E7C…DE8B9`, in Olle's personal store, expires August
  2027). The `.pfx` copy is git-ignored.
- **Logic tests:** `dotnet run --project test\LogicTests`. These link
  `Core\TimeHelpers.cs` and `Features\SendLater\SendLaterTime.cs` directly. Put
  anything that can be wrong, and doesn't need Outlook, into such files.
- **Ribbon XML can be checked against the schema** at
  `…\Visual Studio\18\Insiders\Xml\Schemas\1033\customUI14.xsd`. Add it to an
  `XmlSchemaSet` with the target namespace
  `http://schemas.microsoft.com/office/2009/07/customui`; with a null namespace the
  add fails and everything "validates". One unknown attribute drops the whole tab
  silently.
- **No session can drive Outlook.** "Builds and tests pass" is not "works". Say so,
  and point to the manual test plan in the README.

## Layout

| Path | Contents |
|---|---|
| `ThisAddIn.cs` | Startup: hooks `ItemSend` and `NewInspector`. Nothing else, on purpose |
| `Ribbon.cs`, `Ribbon\*.xml` | One customUI file per Outlook window type; callbacks; `GetVisible` by feature tag |
| `AhkBridge.cs` | `Ping` / `Invoke` / `ListCommands`, as in the PowerPoint add-in |
| `Core\` | `Ol` (run wrapper, context item, owner window), `ComScope`, `Log`, `OutlookSettings`, `ItemProps` (hidden named properties), `TimeHelpers`, `Feature` (the roadmap list and on/off), `Defaults`, `FormStyle`, `Later`, `SettingsForm` |
| `OnSend\` | `IOnSendModule` and `OnSendDispatcher`, the only `ItemSend` handler |
| `Features\SendLater\` | Phase 1 |
| `Features\<Name>\` | Stubs for phases 2–8, each with its design notes in the class comment |

## Rules that hold everywhere

- **Never open an Outbox item just to look at it.** Opening one through the object
  model (`Items[i]`, `GetItemFromID`, `PropertyAccessor` on it) takes it out of the
  send queue, exactly like double-clicking it.
  - Found on 4 November 2026: merely opening the Scheduled window left Olle's
    scheduled email stuck in the Outbox for good.
  - Read the Outbox through `Folder.GetTable` instead (`SendLaterOutbox.ReadTable`,
    `FindInOutbox`, `FolderHasGen`). It lists properties without opening anything.
  - Open an item only to act on it, and then send or delete it in the same step.
  - In a table or a table filter, Prodigo's notes must use `ItemProps.TableSchema`,
    which adds the `/0x0000001F` type suffix. As a column the untyped name throws
    ("does not support this operation"); in a filter it silently matches nothing.
    This was verified against Olle's real Outbox on 4 November 2026.
  - Outlook can be queried read-only from Claude's shell
    (`[Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")`
    plus `GetTable`). Use that to test a table query on the real thing before
    asking Olle to restart Outlook. Never open items that way.

- **Don't release an inspector's own item** (`inspector.CurrentItem`), nor the
  appointment from `GetAssociatedAppointment` inside `ItemSend`. The window, or
  Prodigo's own button further up the call stack, is still using the same object.

- **Fail open.** Nothing in the send path may stop an item because of an error.
  `SendContext.StopSend` is only for an explicit user choice ("Go back" in phase 2).
  The dispatcher discards a stop request from a module that threw.
- **Startup stays empty.** Outlook disables slow add-ins. Load settings, forms and
  folders on first use. Phase 5's timer must start well after startup, not during it.
- **COM lifetime:**
  - Release what you fetched (`ComScope`, `ComScope.Release`).
  - Never release the `ItemSend` item parameter or `Application`.
  - Keep event sources in static fields, or their events stop arriving once the GC
    collects them.
- **Don't save the appointment inside `ItemSend`.** Outlook saves it itself during
  the send. Use `Later.Run` to do it a couple of seconds afterwards.
- **Hidden notes use `ItemProps`** (named properties in Prodigo's own property set
  `{AE5A16EB-…}`), never `UserProperties`. Those cause `winmail.dat` on external
  email. Never change the GUID, or existing stamps become invisible.
- **The class is `Feature`, not `Features`:** `ProdigoOutlook.Features` is the
  namespace of the feature folders, and a class with the same name made
  `Features.SendLater` resolve to the namespace. The ribbon attribute is
  `[ForFeature]` for the same reason.
- **New .cs files must be added to `ProdigoOutlook.csproj` by hand.**
- **Screentips are the contract with AutoHotkey.** `ListCommands` reports them.
  Keep them stable and unique; labels are free to change.

## How Send later works

**Delivery is local, not server-held, on Olle's account** (cached Exchange mode):
- The delayed item waits in the local Outbox, and Outlook sends it when the time
  comes. On 4 November 2026 an email due at 17:23 went at 19:32, when the sleeping
  laptop woke.
- Olle confirmed this is fine: it's how Outlook's own Delay Delivery behaves for him.
  The requirement is "behave like Outlook's Delay Delivery", not "deliver while
  Outlook is closed".
- Don't add workarounds such as server-side sending or wake timers unless he asks.

Stamps on the organizer's appointment:
- `SendLater.Id` links the appointment to its messages.
- `SendLater.Until` is the delivery time.
- `SendLater.Queued` is `1` once the first invite is in the Outbox.

Each outgoing meeting message is tagged with `Id`, `Gen` (unique per message), `Kind`
and `Until`.

1. The Send Later button (`SendLaterCommands`) writes `Until`, records a
   `PendingCommand` in memory, saves, and calls `Send()`.
   - Commands expire after 30 seconds, and a failed `Send()` removes the command and
     restores the old `Until`.
   - Otherwise a send that never happened (an unresolved name, a phase 2 "Go back")
     would leave a command behind to decide some later, unrelated update.
2. `SendLaterOnSend` (in `ItemSend`) finds the appointment through
   `GetAssociatedAppointment(false)`. It decides the delivery time with
   `SendLaterTime.Decide` and sets it with `SendLaterOutbox.ApplyDeferred`. That
   reads the time back and falls back to `PR_DEFERRED_SEND_TIME` if it didn't stick,
   logging which route was used.
   - The initial invite gets the stamped time.
   - Anything sent while it waits gets the stamp + 1 minute, except a reschedule from
     the form, which gets the exact time.
3. When a send replaces something, `ScheduleCleanup` runs 3 seconds later and retries
   for about 18 seconds.
   - Only once the new message is *seen* in the Outbox (or in Sent Items, for Send
     now) does it delete the older waiting messages (`SendLaterTime.PlanCleanup`).
   - **An old invite is deleted only if this send accounts for every one of its
     recipients**: each either gets the new invite or was removed in the same send.
     Recipients are compared by address, from the `To` tag. Without this rule, "send
     updates only to added attendees" or forwarding the meeting would delete the
     invite everyone else was waiting for. An untagged old invite is never deleted.
   - A cancellation to removed attendees is deleted only if none of them is on an
     invite still going out.
   - A whole-meeting cancellation deletes everything for that meeting.
   - If nothing is ever confirmed, nothing is deleted. The +1 minute ordering is then
     the fallback.
4. `OutboxGuard` handles the Outbox trap: opening an Outbox item unsubmits it. When
   such a window closes unsubmitted (`PR_MESSAGE_FLAGS` without `MSGFLAG_SUBMIT`):
   - an email asks whether to put it back;
   - a meeting message goes back without asking.

   Resubmits are flagged so `ItemSend` doesn't retime them.

**Unverified Outlook behaviour that the design leans on.** Each item is in the
README's test plan.
- Custom named properties set on a `MeetingItem` during `ItemSend` persist into the
  Outbox copy. If they don't, cleanup finds nothing and the +1 minute fallback
  applies.
- `Delete()` works on a submitted, deferred Outbox item.
- `MailItem.Send()` on an item already in the Outbox resubmits it, keeping its
  deferred time.
- `AppointmentItem.Delete()` on a sent meeting sends no cancellation. If it does,
  cleanup removes it.
- In a whole-meeting cancellation's `ItemSend`, `MeetingStatus` is already
  `olMeetingCanceled`. If it isn't, the cancellation is treated as one to removed
  attendees, and every waiting invite is still deleted.

## Phase 4: Todo Kanban

Not designed yet; the brief says to ask Olle first. What's known:
- The app is Electron, in `..\TodoKanban\todo-kanban`.
- Its data lives in `%APPDATA%\todo-kanban\kanban-data.json` and in memory, so a
  direct write to the file would be overwritten by the running app.
- Proposed integration points: a drop folder the app imports from, or a
  `todokanban://` link. Both need a Todo Kanban change.
