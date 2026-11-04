# Prodigo for Outlook

The Outlook (classic) member of the Prodigo family, alongside Prodigo for PowerPoint
(`..\Prodigo`) and Prodigo for Excel (`..\ProdigoExcel`). It's a VSTO add-in in C# on
.NET Framework 4.8. It adds a **Prodigo** tab to the main window, the email compose
and read windows, and the appointment/meeting window, each showing only what makes
sense there.

It's built in phases. **Phase 1, Send later, is built.** Phases 2–8 are designed,
stubbed and listed below; their buttons stay hidden until they exist.

---

## Install and update

The add-in installs the same way as the PowerPoint and Excel ones: Outlook loads it
straight from the Visual Studio build folder. There's no installer and no certificate
to buy.

**First time:**

1. Open `ProdigoOutlook.sln` in Visual Studio 2026. You need the *Office/SharePoint
   development* workload, the same as for Prodigo for PowerPoint.
2. **Quit Outlook completely.** Use File → Exit, or right-click the Outlook icon in
   the system tray → Exit. Closing the window isn't enough: Outlook keeps running in
   the tray and keeps the add-in's files locked.
3. Build (Ctrl+Shift+B), or press F5 to start Outlook with the debugger attached.
4. Start Outlook. The **Prodigo** tab appears.

Building registers the add-in under
`HKCU\Software\Microsoft\Office\Outlook\Addins\ProdigoOutlook`, pointing at
`ProdigoOutlook\bin\Debug\ProdigoOutlook.vsto`. That path doesn't change between
builds, so it's written once.

**Updating:** quit Outlook completely, build, start Outlook. Your everyday Outlook
always runs whatever you last built. That's the same deliberate trade-off as the other
two add-ins.

**Signing:** it's signed with the same test certificate as Prodigo for PowerPoint
(`CN=HIQ\OlleGe`, expires August 2027). When that expires, use Project Properties →
Signing → *Create Test Certificate*, the same fix as in the PowerPoint project.

### If the Prodigo tab is missing

1. **Turn on ribbon error messages:** File → Options → Advanced → Developers → *Show
   add-in user interface errors*. One wrong attribute in a ribbon XML file makes
   Office drop the whole tab silently.
2. **Check whether Outlook disabled it:** File → Options → Add-ins → Manage: *COM
   Add-ins* → Go. *Prodigo for Outlook* should be ticked. Also check Manage:
   *Disabled Items*.
3. **Slow-start warnings:** Outlook disables add-ins it considers slow to start and
   shows a yellow bar about it. Prodigo does almost nothing at startup to avoid this.
   If it happens anyway, choose **Always enable this add-in** in that bar (or in
   File → Manage COM Add-ins).

### If a build misbehaves

Prodigo → Settings → untick the feature. Its buttons disappear and it stops taking
part in sending, immediately, with no rebuild. Prodigo also "fails open": if its
own code hits an error while you send, the email or invite goes out normally and the
error goes to the log.

---

## Features and roadmap

| Phase | Feature | Status | Where |
|---|---|---|---|
| 1 | **Send later**: emails and meeting invites, editable and reschedulable, Scheduled list, room warning, Outbox protection | ✅ Built (untested in Outlook) | Compose, Appointment, Main window |
| 2 | Pre-send checks: missing attachment, empty subject, big reply-all, internal and external recipients mixed, wrong account, weekday/date mismatch. Rules in a JSON file | Planned | On send |
| 3 | Insert availability as text ("Jag kan tis 10–12, ons efter 14…") | Planned | Compose |
| 4 | Email to Todo Kanban | Planned; integration point to be decided | Main window, Read |
| 5 | Follow-up tracker: "remind me if no reply in X days" | Planned | Compose, Main window |
| 6 | Meeting presets: length, agenda, attendees from an email, Teams | Planned | Main window, Read, Appointment |
| 7 | File email (.msg and attachments) to a project folder | Planned | Main window, Read |
| 8 | Copy recipients, clean subject prefixes, email attendees | Planned | All windows |

---

## Send later

### Emails

In an email window, go to **Prodigo → Send Later** (Alt, Z, P, S, L). A small window
opens:

- **Preset buttons**, each showing its real date: *In 1 hour*, *Tomorrow 08:00*,
  *Next workday 08:00*, *Monday 08:00*. **One click, or press 1–4, and the email is
  sent**, to be delivered at that time.
- **Other time:** pick a date and time, then click **Send**. Times in the past, or
  less than a minute away, are refused with a note under the picker.

The email waits in your **Outbox** and Exchange delivers it at that time, even with
Outlook closed. To cancel it, delete it from the Outbox or use **Scheduled** (below).

**Editing a scheduled email:** open it from the Outbox, change it, and press
**Send**. The original time is kept. If the email has lost its delivery time by
then, Prodigo asks *"Keep it scheduled for …?"*, since it can't tell whether Outlook
dropped the time or you removed it.

To change the time, press **Send Later** in that window: the form shows *"Currently
scheduled: …"* and offers a new time or **Send now**. A time set with Outlook's own
Delay Delivery is respected.

### Meetings

In a meeting you organise, go to **Prodigo → Send Later**. It's hidden on meetings
you were invited to. You get the same window, plus a **warning if a meeting room is
among the attendees**: the room isn't reserved until the invite is delivered.

What happens:

- The meeting appears in your calendar straight away. Teams links are created now.
- The invite waits in your Outbox. Exchange holds it and delivers it at the chosen
  time, even with Outlook closed.
- Attendees get nothing until then. Their calendars and free/busy don't know about
  the meeting yet.
- At the chosen time it's delivered, moves to Sent Items, and replies start coming in.

**Editing a scheduled meeting:** open it **in your calendar** (not the Outbox),
change it and press **Send Update**. Prodigo replaces the waiting invite with the new
version, so attendees get one up-to-date invite. It's delivered 1 minute after the
original time.

The minute is a safety margin. The new version is held until just after the original
time, and only once it's safely in the Outbox is the old one deleted. If deleting the
old one fails for any reason, attendees get the invite and then the update a minute
later, never the other way round.

**Changing the time:** open the meeting in your calendar and press **Send Later**.
Pick a new time, or **Send now**. The waiting invite is replaced.

**Removing attendees before delivery:** the people you removed get nothing at all.
Outlook's cancellation to them is deleted together with the old invite.

**Updates to only some attendees** (Outlook's *"send updates only to added or deleted
attendees"*, or forwarding the meeting) never replace the waiting invite. Prodigo
only replaces it when the new version goes to everyone the old one did, so nobody can
end up with nothing.

**Cancelling the meeting before delivery:** use **Cancel send** in the Scheduled
window. If you use Outlook's own *Cancel Meeting*, Prodigo removes both the waiting
invite and the cancellation, so nobody hears anything either way.

**If you open the waiting invite from the Outbox,** Prodigo tells you to edit the
meeting in the calendar instead and offers to open it there.

### Scheduled (main window)

**Prodigo → Scheduled** (Alt, Z, P, S, C) lists everything waiting in the Outbox to
be sent later, including emails delayed with Outlook's own Delay Delivery. It has
buttons for **Change time…**, **Send now**, **Cancel send**, **Open** and **Refresh**.
For a meeting, the row acts on the meeting itself.

### The Outbox trap, and how Prodigo handles it

In classic Outlook, **opening an item in the Outbox takes it out of the send queue**.
If you close it without pressing Send, it sits there and is never delivered. Prodigo
watches for this:

- **Email:** when you close it, Prodigo asks *"Keep it scheduled for …?"*. **Yes**
  puts it back in the queue.
- **Meeting invite:** it's put back automatically.

---

## Keytips and AutoHotkey

The Prodigo tab's keytip is **ZP** in every window. Each feature keeps the same
keytip everywhere.

| Keys | Command | Windows | Status |
|---|---|---|---|
| Alt, Z, P, S, L | Send Later | Compose, Appointment | Built |
| Alt, Z, P, S, C | Scheduled Sends | Main window | Built |
| Alt, Z, P, S, T | Prodigo Settings | All | Built |
| Alt, Z, P, A, V | Insert Availability | Compose | Phase 3 |
| Alt, Z, P, T, D | Email to Todo | Main window, Read | Phase 4 |
| Alt, Z, P, F, U | Follow-ups / Remind If No Reply | Main window / Compose | Phase 5 |
| Alt, Z, P, M, P | Meeting Preset | Main window, Read, Appointment | Phase 6 |
| Alt, Z, P, F, E | File Email | Main window, Read | Phase 7 |
| Alt, Z, P, C, R | Copy Recipients | All | Phase 8 |
| Alt, Z, P, E, A | Email Attendees | Main window, Appointment | Phase 8 |
| Alt, Z, P, C, S | Clean Subject | Compose | Phase 8 |

In the Send Later window, **1–9** pick a preset, **Alt+S** sends at the picked time,
**Alt+N** is Send now, and **Esc** cancels.

**For hotkeys, use the bridge rather than keystrokes.** Keytip sequences depend on
timing, and the AHK notes record one losing its Alt. The bridge is the same one the
PowerPoint add-in has:

```autohotkey
; AutoHotkey v2
ProdigoOl(cmd) {
    static svc := ""
    if !IsObject(svc) {
        try {
            ai := ComObjActive("Outlook.Application").COMAddIns.Item("ProdigoOutlook")
            if !ai.Connect
                ai.Connect := true
            svc := ai.Object
        } catch {
            return "ERROR: no bridge"
        }
    }
    try
        return svc.Invoke(cmd)
    catch as e {
        svc := ""
        return "ERROR: " e.Message
    }
}

#HotIf WinActive("ahk_exe OUTLOOK.EXE")
^!l::ProdigoOl("Send Later")
#HotIf
```

- `Ping` tells you the bridge is alive.
- `ListCommands` lists what's available. Only commands whose feature is built and
  switched on are listed.
- `Invoke("Send Later")` runs a command in the Outlook window in front, exactly as if
  you'd clicked the button. It returns `OK` once the command has been dispatched.

The names are the buttons' screentips: *Send Later*, *Scheduled Sends*, *Prodigo
Settings*. Case and spaces don't matter.

---

## Settings and files

| What | Where |
|---|---|
| Settings (feature switches, presets, room warning) | `%APPDATA%\Prodigo\outlook-settings.txt`, plain `key = value` |
| Log, one file per day, 14 days kept | `%LOCALAPPDATA%\Prodigo\Logs\outlook-yyyy-MM-dd.log` |
| Phase 2 rules (future) | `%APPDATA%\Prodigo\outlook-rules.json` |

**Prodigo → Settings** edits the settings file and has buttons to open both folders.

Presets are one per line, top to bottom. They can be:
- `+1h`, `+30m`, `+1h30m`
- `today 13:00`, `tomorrow 08:00`
- `nextworkday 08:00`
- `monday 08:00` (also Swedish: `måndag 08:00`)

Prodigo's notes on items ("stamps") are hidden properties in Prodigo's own property
set. They aren't Outlook "user properties", which on email can cause `winmail.dat`
for external recipients.

---

## Manual test plan

### Phase 1: Send later

You need a colleague (or a second mailbox) as recipient. Keep today's log open
(Settings → Open log folder): it records which route set each delivery time and
every invite that was replaced.

**Setup**
- [ ] After a build and an Outlook restart, the Prodigo tab is in the main window,
  a new email, a received email, and a new meeting.
- [ ] Send Later is **not** shown on a meeting you were invited to.
- [ ] Prodigo → Settings → untick Send later → OK. The Send Later and Scheduled
  buttons disappear. Tick it again and they come back.

**Emails**
- [ ] Send Later → *In 1 hour* sends in one click. The email is in the Outbox, and
  the log says `Deferred email … (DeferredDeliveryTime)`. It arrives about an hour
  later.
- [ ] A time in the past, or 30 seconds from now, is refused with a message.
- [ ] Keys 1–4 pick presets; Esc cancels and nothing is sent.
- [ ] Schedule an email, open it from the Outbox, change the text, press Send. It
  still arrives at the original time, with the new text.
- [ ] Open a scheduled email from the Outbox and close it **without** sending.
  Prodigo asks *Keep it scheduled?*. Answer Yes: it's delivered on time. Repeat and
  answer No: it stays unsent in the Outbox.
- [ ] Scheduled → Change time… moves an email; Send now sends it immediately;
  Cancel send removes it.
- [ ] An email scheduled with Outlook's own Delay Delivery also shows in Scheduled.

**Meetings**
- [ ] **Arrives at the right time:** schedule an invite for 15 minutes from now.
  Your calendar shows it straight away. The colleague gets nothing until the time,
  then gets the invite. The log says which route set the time (`DeferredDeliveryTime`
  or the `PR_DEFERRED_SEND_TIME` fallback).
- [ ] **Outlook closed:** schedule an invite, quit Outlook completely, and confirm
  it still arrives on time.
- [ ] **Update before delivery:** schedule, then change the time or location in the
  calendar and Send Update. The colleague gets **one** invite, with the change, about
  1 minute after the original time. The log says `replaced 1 waiting message(s)`.
- [ ] **Remove an attendee before delivery:** the removed person gets nothing; the
  others get the updated invite.
- [ ] **Add an attendee and choose "send updates only to added attendees":** the new
  person gets the invite, *and* the original attendees still get theirs at the
  original time.
- [ ] **Forward the waiting meeting** from the calendar to someone: the original
  attendees still get their invite.
- [ ] **Meeting from a second mailbox** (if you have one): it shows in Scheduled,
  and updates replace the waiting invite.
- [ ] **Cancel before delivery, via Scheduled → Cancel send:** nobody gets anything,
  and the meeting leaves your calendar.
- [ ] **Cancel before delivery, via Outlook's Cancel Meeting:** nobody gets anything.
  If the log says *never showed up in the Outbox*, the fallback applied instead:
  invite, then cancellation a minute later.
- [ ] **Reschedule:** Send Later on a waiting meeting shows *Currently scheduled* and
  moves it. The colleague gets one invite at the new time.
- [ ] **Send now** on a waiting meeting: delivered immediately, and the old waiting
  invite is gone.
- [ ] **Open the waiting invite from the Outbox:** Prodigo points to the calendar.
  Close the invite anyway: it's still delivered on time.
- [ ] **Deleting the invite from the Outbox** cancels the delayed send.
- [ ] **Recurring meeting** scheduled for later: arrives on time, with the series
  intact.
- [ ] **Teams meeting** scheduled for later: the Teams link is in the delivered
  invite and works.
- [ ] **Room warning:** add a meeting room as an attendee. The warning shows in the
  Send Later window. Turn it off in Settings and it's gone.
- [ ] **No delay:** an ordinary meeting sent with Outlook's Send button goes out
  immediately, as before.
- [ ] Repeat the first three meeting tests in **cached mode** and in **online mode**.
  To switch: File → Account Settings → your account → Change → *Use Cached Exchange
  Mode*.

**Fail-open check**
- [ ] With Send later on, ordinary emails and meetings without a delay go out
  exactly as before. The log has no errors for them.

### Phases 2–8

These get written when each phase is built. At minimum, each will cover: the
feature's happy path, the switch in Settings, its keytip and bridge name, and (for
anything that runs on send) that an error still lets the item go.

---

## Differences from the other Prodigo add-ins, and why

| | PowerPoint | Excel | Outlook | Why |
|---|---|---|---|---|
| Project type | VSTO | Excel-DNA | VSTO | Excel-DNA is Excel-only. Outlook needs VSTO's events (sending) and per-window ribbons. |
| .NET | 4.7.2 | 4.8 | 4.8 | Newer of the two, as in Excel. |
| Ribbon | One XML | Inline XML | **Four XML files** | Outlook asks for a ribbon per kind of window. |
| Settings | .NET settings (reset on version bump) | Text file | Text file (Excel's pattern) | Readable and portable, and doesn't reset on updates. |
| Logging | None | None | **Daily log** | Code that runs while sending must fail silently; the log is the only trace. |
| Keytips | None | None | **On every button** | Asked for. The bridge is still the reliable route for hotkeys. |
| COM cleanup | Not needed | Not needed | **ComScope helper** | Outlook items held open cause "item was changed" errors and Exchange limits. |
| Startup | Loads guides | Queues shortcuts | **Hooks two events only** | Outlook disables add-ins that start slowly. |
| Tests | `TestDecks` | `test\AuditCheck` | `test\LogicTests` | Same idea as Excel: logic without Outlook types, tested for real. |

---

## For development

- **Build:** Visual Studio, or MSBuild from the command line (`vswhere -find
  MSBuild\**\Bin\MSBuild.exe`). A command-line build also registers the add-in.
- **Logic tests:** `dotnet run --project test\LogicTests` from the repository
  folder. The exit code is the number of failures.
- **New source file:** add it to `ProdigoOutlook.csproj` by hand. VSTO projects list
  every file, and an unlisted file is silently not compiled.
- **Project notes for Claude** are in `CLAUDE.md`.
