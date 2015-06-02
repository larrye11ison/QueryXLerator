using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace QueryXLerator
{
    internal class ExcelTableStyleSampleImages
    {
        private static List<ImageSource> DarkStyles = null;
        private static List<ImageSource> LightStyles = null;
        private static List<ImageSource> MediumStyles = null;

        public static ImageSource GetImageForStyle(string styleName)
        {
            if (DarkStyles == null)
            {
                GetImages();
            }
            List<ImageSource> theListToUse = null;

            if (styleName.IndexOf("Dark") > -1)
            {
                theListToUse = DarkStyles;
            }
            if (styleName.IndexOf("Light") > -1)
            {
                theListToUse = LightStyles;
            }
            if (styleName.IndexOf("Medium") > -1)
            {
                theListToUse = MediumStyles;
            }

            var digits = System.Text.RegularExpressions.Regex.Match(styleName, @"\d+$").Value;

            int index;
            if (int.TryParse(digits, out index))
            {
                return theListToUse[index - 1];
            }
            return null;
        }

        internal static void GetImages()
        {
            DarkStyles = new List<ImageSource>();
            LightStyles = new List<ImageSource>();
            MediumStyles = new List<ImageSource>();
            var executingAssembly = Assembly.GetExecutingAssembly();
            var streamsIWant = executingAssembly.GetManifestResourceNames()
                .Where(rn => rn.IndexOf("ExcelTableStyles", StringComparison.CurrentCultureIgnoreCase) >= 0);
            foreach (var streamName in streamsIWant)
            {
                using (var resourceStream = executingAssembly.GetManifestResourceStream(streamName))
                {
                    BitmapImage bmi = new BitmapImage();

                    MemoryStream memStream = new MemoryStream();
                    resourceStream.CopyTo(memStream);

                    List<ImageSource> theListToUse = null;

                    if (streamName.IndexOf("Dark") > -1)
                    {
                        theListToUse = DarkStyles;
                    }
                    if (streamName.IndexOf("Light") > -1)
                    {
                        theListToUse = LightStyles;
                    }
                    if (streamName.IndexOf("Medium") > -1)
                    {
                        theListToUse = MediumStyles;
                    }
                    var bm = new BitmapImage();

                    bm.BeginInit();
                    bm.StreamSource = memStream;
                    bm.EndInit();
                    bm.Freeze();

                    int numImages = (int)bm.Width / 61;

                    for (int i = 0; i < numImages; i++)
                    {
                        CroppedBitmap cbm = new CroppedBitmap(bm, new Int32Rect(61 * i, 0, 61, (int)bm.Height));

                        theListToUse.Add(cbm);
                    }
                }
            }
        }
    }
}