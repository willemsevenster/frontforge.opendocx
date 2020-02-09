using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using Frontforge.OpenDocx.Core.ModelConfiguration;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Frontforge.OpenDocx.Core.Models {
    public class ImageContent : IContent
    {
        private readonly ImageContentConfig _config;

        internal ImageContent(ImageContentConfig config)
        {
            _config = config ?? new ImageContentConfig();
        }

        internal ImageContent(string name)
        {
            _config = new ImageContentConfig {Name = name};
        }


        public Run GetRun(RunProperties runProperties)
        {
            var run = new Run {RunProperties = runProperties.CloneNode()};

            run.RunProperties.AppendChild(new NoProof());

            var drawing = new Drawing();
            var inline = new Inline
            {
                DistanceFromTop = (UInt32Value)0U, 
                DistanceFromBottom = (UInt32Value)0U, 
                DistanceFromLeft = (UInt32Value)0U, 
                DistanceFromRight = (UInt32Value)0U
            };

            var nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            var graphicFrameLocks1 = new A.GraphicFrameLocks(){ NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);


            var graphic1 = new Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var graphicData1 = new GraphicData(){ Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            var picture1 = new DocumentFormat.OpenXml.Drawing.Pictures.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            var nonVisualPictureProperties1 = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties();
            var nonVisualDrawingProperties1 = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = _config.Name };
            var nonVisualPictureDrawingProperties1 = new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            
            var blipFill1 = new Pic.BlipFill();
            var blip1 = new Blip(){ Embed = _config.Name };
            var blipExtensionList1 = new BlipExtensionList();

            var blipExtension1 = new BlipExtension(){ Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            var useLocalDpi1 = new A14.UseLocalDpi(){ Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            var stretch1 = new Stretch();
            var fillRectangle1 = new FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            var shapeProperties1 = new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties();

            var transform2D1 = new Transform2D();
            var offset1 = new Offset(){ X = 0L, Y = 0L };

            transform2D1.Append(offset1);

            var presetGeometry1 = new PresetGeometry(){ Preset = ShapeTypeValues.Rectangle };
            var adjustValueList1 = new AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline.Append(nonVisualGraphicFrameDrawingProperties1);
            inline.Append(graphic1);
            drawing.Append(inline);
            run.AppendChild(drawing);
            return run;
        }
    }
}