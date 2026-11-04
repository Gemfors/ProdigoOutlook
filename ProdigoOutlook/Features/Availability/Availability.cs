namespace ProdigoOutlook.Features.Availability
{
    /// <summary>
    /// PHASE 3 - planned, not built. Insert Availability in the compose window.
    ///
    /// Asks for a meeting length and how many days ahead to look, reads Olle's own
    /// calendar through the object model (no Graph needed), respects working hours
    /// from Settings, and inserts the free slots as readable text at the cursor:
    /// "Jag kan tis 10–12, ons efter 14 eller tors förmiddag", or the English version.
    ///
    /// Split it the usual way: reading busy times from Outlook in one class, and turning
    /// busy times into free slots and slots into a sentence in another with no Outlook
    /// types, so the sentence-building can be tested.
    /// </summary>
    internal static class Availability
    {
    }
}
