using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders {
    public class ImageContentBuilder
    {
        private readonly ImageContentConfig _config = new ImageContentConfig();

        internal ImageContentBuilder(string name)
        {
            _config.Name = name;
        }

        public static implicit operator ImageContent(ImageContentBuilder builder)
        {
            return new ImageContent(builder._config);
        }
    }
}
