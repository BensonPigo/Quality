using Sci;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace ImageCompress
{
    public class ImageHelper
    {
        public static byte[] ImageCompress(byte[] InputImageBytes, Bitmap bitmap = null)
        {
            if (InputImageBytes == null)
            {
                return InputImageBytes;
            }

            byte[] resultBytes = null;
            if (bitmap == null)
            {
                resultBytes = InputImageBytes;
            }
            else
            {
                EncoderParameter encoderParameter = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, 100L);
                EncoderParameters encoderParameters = new EncoderParameters();
                encoderParameters.Param[0] = encoderParameter;
                ImageCodecInfo imageCodecInfo = ImageCodecInfo.GetImageEncoders().Where(s => s.FormatID == ImageFormat.Jpeg.Guid).First();

                using (MemoryStream stream = new MemoryStream())
                {
                    bitmap.Save(stream, imageCodecInfo, encoderParameters);
                    resultBytes = stream.ToArray();
                }

            }

            MemoryStream oMemoryStream = new MemoryStream(InputImageBytes);
            Image oImage = System.Drawing.Image.FromStream(oMemoryStream);
            Bitmap oBitmap = bitmap == null ? new Bitmap(oImage) : bitmap;

            if (resultBytes.Length <= 524288)
            {
                return resultBytes;
            }
            else
            {
                // 每次超過500kb就將照片縮20%
                return ImageCompress(InputImageBytes, ResizeBitmap(oBitmap, 0.8));
            }

        }
        private static Bitmap ResizeBitmap(Bitmap bmp, double zoomRatio)
        {
            int newWidth = MyUtility.Convert.GetInt(bmp.Width * zoomRatio);
            int newHeight = MyUtility.Convert.GetInt(bmp.Height * zoomRatio);
            Bitmap result = new Bitmap(newWidth, newHeight);
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(bmp, 0, 0, newWidth, newHeight);
            }

            return result;
        }
    }
}
