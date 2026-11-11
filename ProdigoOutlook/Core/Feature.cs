using System.Collections.Generic;
using System.Linq;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// Every feature on the roadmap, built or not, in one list.
    ///
    /// A feature's ribbon buttons are shown only when it is built *and* switched on in
    /// Settings (Ribbon.GetVisible asks <see cref="IsOn"/>), and the on-send dispatcher
    /// only calls modules whose feature is on. So the off switch in Settings really
    /// does take a feature out of the way - the escape hatch if a build misbehaves.
    ///
    /// When a planned feature gets built: set Built to true here, give its buttons a
    /// real onAction in the ribbon XML, and that is all the wiring there is.
    /// </summary>
    internal static class Feature
    {
        public const string SendLater = "sendlater";
        public const string PreSendChecks = "presend";
        public const string Availability = "availability";
        public const string EmailToTodo = "todo";
        public const string FollowUp = "followup";
        public const string MeetingPresets = "meetingpresets";
        public const string FileEmail = "fileemail";
        public const string Conveniences = "conveniences";

        public static readonly IReadOnlyList<FeatureInfo> All = new[]
        {
            new FeatureInfo(SendLater,      1, true,  "Send later", "Schedule emails and meeting invites, and edit or reschedule them before they go"),
            new FeatureInfo(PreSendChecks,  2, true,  "Pre-send checks", "Before an email goes: missing attachment, reply to many, inside and outside HiQ mixed, wrong weekday for a date"),
            new FeatureInfo(Availability,   3, true,  "Insert availability", "Insert your free times from the calendar as text"),
            new FeatureInfo(EmailToTodo,    4, false, "Email to todo", "Create a task in Todo Kanban from an email"),
            new FeatureInfo(FollowUp,       5, false, "Follow-up tracker", "Remind you when a sent email gets no reply"),
            new FeatureInfo(MeetingPresets, 6, false, "Meeting presets", "New meeting with preset length, agenda and attendees"),
            new FeatureInfo(FileEmail,      7, false, "File email", "Save an email and its attachments to a project folder"),
            new FeatureInfo(Conveniences,   8, false, "Small conveniences", "Copy recipients, email attendees, clean subject prefixes"),
        };

        public static FeatureInfo Find(string id) => All.FirstOrDefault(f => f.Id == id);

        /// <summary>Built, and not switched off in Settings.</summary>
        public static bool IsOn(string id)
        {
            FeatureInfo info = Find(id);
            return info != null && info.Built && OutlookSettings.Current.IsFeatureOn(id);
        }
    }

    internal sealed class FeatureInfo
    {
        public FeatureInfo(string id, int phase, bool built, string name, string description)
        {
            Id = id;
            Phase = phase;
            Built = built;
            Name = name;
            Description = description;
        }

        public string Id { get; }
        public int Phase { get; }
        public bool Built { get; }
        public string Name { get; }
        public string Description { get; }
    }
}
