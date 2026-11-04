namespace ProdigoOutlook.OnSend
{
    /// <summary>
    /// A feature that wants a say when an item is sent. The dispatcher calls every
    /// module whose feature is switched on, lowest <see cref="Order"/> first.
    ///
    /// Orders in use, with room between them:
    ///   100  Pre-send checks (phase 2) - may ask the user, and may stop the send
    ///   200  Follow-up tracker (phase 5) - tags the outgoing item
    ///   300  Send later (phase 1) - last, so it schedules the item as it finally is
    /// </summary>
    internal interface IOnSendModule
    {
        /// <summary>The Features id that switches this module on and off.</summary>
        string FeatureId { get; }

        int Order { get; }

        /// <summary>
        /// Called with the item being sent (MailItem, MeetingItem, ...). Throwing is
        /// safe: the dispatcher logs it and lets the item go. The only way to stop the
        /// send is <see cref="SendContext.StopSend"/>, for an explicit user choice.
        /// </summary>
        void OnSend(object item, SendContext context);
    }

    /// <summary>What a module can tell the dispatcher about the send in progress.</summary>
    internal sealed class SendContext
    {
        internal bool StopRequested { get; private set; }

        internal string StopReason { get; private set; }

        /// <summary>
        /// Stops the item from being sent. Only for a decision the user made just now,
        /// such as "Go back" in a pre-send warning - never as a reaction to an error.
        /// If the module throws after calling this, the request is discarded and the
        /// item is sent after all.
        /// </summary>
        public void StopSend(string reason)
        {
            StopRequested = true;
            StopReason = reason;
        }

        internal void Withdraw()
        {
            StopRequested = false;
            StopReason = null;
        }
    }
}
