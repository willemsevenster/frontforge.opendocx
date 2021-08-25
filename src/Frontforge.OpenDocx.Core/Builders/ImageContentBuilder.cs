using Frontforge.OpenDocx.Core.ModelConfiguration;
using Frontforge.OpenDocx.Core.Models;

namespace Frontforge.OpenDocx.Core.Builders {
    public class ImageContentBuilder
    {
        #region instance fields

        private readonly ImageContentConfig _config = new ImageContentConfig();

        #endregion

        #region constructors

        internal ImageContentBuilder(string name)
        {
            _config.Name = name;
        }

        #endregion

        #region implementation

        public static implicit operator ImageContent(ImageContentBuilder builder)
        {
            return new ImageContent(builder._config);
        }

        #endregion
    }
}
