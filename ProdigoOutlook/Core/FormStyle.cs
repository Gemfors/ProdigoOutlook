using System.Drawing;
using System.Windows.Forms;

namespace ProdigoOutlook.Core
{
    /// <summary>
    /// The look every Prodigo dialog shares, applied in one call.
    ///
    /// AutoScaleDimensions matters as much as AutoScaleMode, a lesson from ProdigoExcel:
    /// with only AutoScaleMode.Font set the scale factor stays 1:1, so a form laid out
    /// in pixels clips its own labels at 150% display scaling while looking perfect at
    /// 100%. SizeF(7F, 15F) is Segoe UI 9pt at 96 DPI, the font set here. Call this
    /// before adding any controls, so their pixel positions are what gets scaled.
    /// </summary>
    internal static class FormStyle
    {
        public static void ApplyDialog(Form form, string title)
        {
            form.Text = title;
            form.Font = new Font("Segoe UI", 9F);
            form.AutoScaleDimensions = new SizeF(7F, 15F);
            form.AutoScaleMode = AutoScaleMode.Font;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterParent;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.ShowInTaskbar = false;
        }
    }
}
