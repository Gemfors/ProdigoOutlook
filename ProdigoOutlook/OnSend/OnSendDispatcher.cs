using System;
using System.Collections.Generic;
using System.Linq;
using ProdigoOutlook.Core;
using ProdigoOutlook.Features.FollowUp;
using ProdigoOutlook.Features.PreSendChecks;
using ProdigoOutlook.Features.SendLater;

namespace ProdigoOutlook.OnSend
{
    /// <summary>
    /// The one handler on Application.ItemSend. Every feature that cares about sending
    /// plugs in here as an <see cref="IOnSendModule"/>, so adding one never means
    /// touching another.
    ///
    /// Fails open, always. Each module runs inside its own try/catch; an exception is
    /// logged and the next module still runs. The send is stopped only when a module
    /// asked for it on the user's behalf *and* finished without throwing. An error can
    /// never be the reason an email or invite does not go.
    /// </summary>
    internal static class OnSendDispatcher
    {
        /// <summary>
        /// Every module, built or not. Not-yet-built features are listed so their order
        /// is fixed now; Feature.IsOn keeps them out until they exist.
        /// </summary>
        private static readonly IOnSendModule[] Modules =
        {
            new PreSendChecksOnSend(),
            new FollowUpOnSend(),
            new SendLaterOnSend(),
        };

        public static void OnItemSend(object item, ref bool cancel)
        {
            var context = new SendContext();

            List<IOnSendModule> active;
            try
            {
                active = Modules.Where(m => Feature.IsOn(m.FeatureId)).OrderBy(m => m.Order).ToList();
            }
            catch (Exception ex)
            {
                Log.Error("ItemSend: could not work out which modules to run; sending as normal", ex);
                return;
            }

            foreach (IOnSendModule module in active)
            {
                bool stoppedBefore = context.StopRequested;
                try
                {
                    module.OnSend(item, context);
                }
                catch (Exception ex)
                {
                    Log.Error("ItemSend: module " + module.FeatureId + " failed; item is sent as normal", ex);
                    if (!stoppedBefore) context.Withdraw();
                }

                // Once the user has said "Go back", later modules (send later above all)
                // must not schedule an item that is not going anywhere.
                if (context.StopRequested) break;
            }

            if (context.StopRequested)
            {
                Log.Info("ItemSend: send stopped by user choice - " + context.StopReason);
                cancel = true;
            }
        }
    }
}
