using Frontforge.OpenDocx.Core.ModelConfiguration;

namespace Frontforge.OpenDocx.Core.Models
{
    public class CheckboxControl
        : IContent
    {
        private readonly CheckboxControlConfig _config;

        internal CheckboxControl(CheckboxControlConfig config)
        {
            _config = config;
        }

        internal bool IsChecked => _config?.Checked ?? false;

        internal string Label => _config?.Label;
    }
}