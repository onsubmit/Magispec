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
                g.SmoothingMode = SmoothingMode.None;
                g.InterpolationMode = InterpolationMode.NearestNeighbor;
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
            var diffR = Math.Abs(a.R - b.R);
            var diffG = Math.Abs(a.G - b.G);
            var diffB = Math.Abs(a.B - b.B);

            return Math.Max(Math.Max(diffR, diffG), diffB);
        }
    }
}