//-----------------------------------------------------------------------
// <copyright file="Extensions.cs" company="Andy Young">
//     Copyright (c) Andy Young. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace Magispec
{
    using System;
    using System.Drawing;
    using System.Drawing.Drawing2D;

    /// <summary>
    /// Extension methods
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Resizes an image
        /// </summary>
        /// <param name="original">The image to resize</param>
        /// <param name="width">The new width in pixels</param>
        /// <param name="height">The new height in pixels</param>
        /// <returns>A resized version of the original image</returns>
        public static Image Resize(this Image original, int width, int height)
        {
            Image newImage = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(newImage))
            {
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(original, 0, 0, width, height);
            }

            return newImage;
        }

        /// <summary>
        /// Compares one color to another
        /// </summary>
        /// <param name="a">First color</param>
        /// <param name="b">Second color</param>
        /// <returns>A number close to 100 if similar</returns>
        public static double CompareTo(this Color a, Color b)
        {
            return 100.0 * (1.0 - ((double)(Math.Abs(a.R - b.R) + Math.Abs(a.G - b.G) + Math.Abs(a.B - b.B)) / (765.0))); // 255 * 3
        }
    }
}