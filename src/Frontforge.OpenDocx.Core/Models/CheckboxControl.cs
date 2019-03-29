using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class CheckboxControl
        : IContent
    {
        #region instance fields

        private readonly CheckboxControlConfig _config;

        #endregion

        #region constructors

        internal CheckboxControl(CheckboxControlConfig config)
        {
            _config = config;
        }

        #endregion

        #region properties

        internal bool IsChecked => _config?.Checked ?? false;

        internal string Label => _config?.Label;

        #endregion
    }
}