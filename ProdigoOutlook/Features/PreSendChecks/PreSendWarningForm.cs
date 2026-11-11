using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using ProdigoOutlook.Core;

namespace ProdigoOutlook.Features.PreSendChecks
{
    /// <summary>
    /// "Check before sending": everything the checks found, in one list, with Go back
    /// and Send anyway. Go back is the default, so Enter pressed out of habit does not
    /// send a mistake; Esc and closing the window also mean Go back.
    /// </summary>
    internal sealed class PreSendWarningForm : Form
    {
        private bool _sendAnyway;

        private PreSendWarningForm(IList<PreSendFinding> findings)
        {
            FormStyle.ApplyDialog(this, "Prodigo - Check before sending");
            Build(findings);
        }

        /// <summary>True for Send anyway; false for Go back.</summary>
        public static bool Ask(IWin32Window owner, IList<PreSendFinding> findings)
        {
            using (var form = new PreSendWarningForm(findings))
            {
                form.ShowDialog(owner);
                return form._sendAnyway;
            }
        }

        private void Build(IList<PreSendFinding> findings)
        {
            const int left = 16, width = 460;
            int top = 14;

            Controls.Add(new Label
            {
                Text = "Check before sending",
                Left = left, Top = top, Width = width, Height = 22,
                Font = new Font(Font.FontFamily, 11F, FontStyle.Bold),
            });
            top += 30;

            foreach (PreSendFinding f in findings)
            {
                var label = new Label
                {
                    Text = "•  " + f.Message,
                    Left = left, Top = top, Width = width,
                    MaximumSize = new Size(width, 0), AutoSize = true,
                };
                Controls.Add(label);
                top += label.GetPreferredSize(new Size(width, 0)).Height + 8;
            }
            top += 8;

            var goBack = new Button { Text = "&Go back", Left = left + width - 230, Top = top, Width = 110, Height = 30 };
            goBack.Click += (s, e) => { _sendAnyway = false; Close(); };
            var send = new Button { Text = "&Send anyway", Left = left + width - 110, Top = top, Width = 110, Height = 30 };
            send.Click += (s, e) => { _sendAnyway = true; Close(); };
            Controls.AddRange(new Control[] { goBack, send });
            AcceptButton = goBack;
            CancelButton = goBack;
            ActiveControl = goBack;

            ClientSize = new Size(left * 2 + width, top + 30 + 14);
            TopMost = true;
        }
    }
}
