﻿using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Pulsar.Common.Video.Compression
{
    public class JpgCompression : IDisposable
    {
        private readonly ImageCodecInfo _encoderInfo;
        private readonly EncoderParameters _encoderParams;

        public JpgCompression(long quality)
        {
            EncoderParameter parameter = new EncoderParameter(Encoder.Quality, quality);
            this._encoderInfo = GetEncoderInfo("image/jpeg");
            this._encoderParams = new EncoderParameters(1); // Only one parameter now
            this._encoderParams.Param[0] = parameter;
            // Removed: this._encoderParams.Param[1] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionRle);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_encoderParams != null)
                {
                    _encoderParams.Dispose();
                }
            }
        }

        public byte[] Compress(Bitmap bmp)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                bmp.Save(stream, _encoderInfo, _encoderParams);
                return stream.ToArray();
            }
        }

        public void Compress(Bitmap bmp, ref Stream targetStream)
        {
            bmp.Save(targetStream, _encoderInfo, _encoderParams);
        }

        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] imageEncoders = ImageCodecInfo.GetImageEncoders();
            int num2 = imageEncoders.Length - 1;
            for (int i = 0; i <= num2; i++)
            {
                if (imageEncoders[i].MimeType == mimeType)
                {
                    return imageEncoders[i];
                }
            }
            return null;
        }
    }
}