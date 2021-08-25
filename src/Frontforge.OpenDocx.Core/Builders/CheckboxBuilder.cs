using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders
{
    public class CheckboxBuilder
    {
        private readonly CheckboxControlConfig _config = new CheckboxControlConfig();

        internal CheckboxBuilder() { }

        public static implicit operator CheckboxControl(CheckboxBuilder builder)
        {
            return new CheckboxControl(builder._config);
        }

        public CheckboxBuilder Check(bool? value = true)
        {
            return this.Chain(p => p._config.Checked = value);
        }

        public CheckboxBuilder Label(string value)
        {
            return this.Chain(p => p._config.Label = value);
        }
    }
}
