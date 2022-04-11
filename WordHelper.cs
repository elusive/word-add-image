namespace word_add_diagram
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    using Word = Microsoft.Office.Interop.Word;

    public class WordHelper
    {
        public static Word.Window GetOpenDocument()
        {
            try
            {
                Word.Application wordObject;
                wordObject = (Word.Application)NativeMethods.GetActiveObject("Word.Application");
                var window = wordObject.ActiveWindow; // get first open word document.
                return window;

                // returns list of strings for each document path
                //var openDocs = new List<string>();
                //for (var i=0; i < wordObject.Windows.Count; i++)
                //{
                //    object idx = i + 1;
                //    var windowObject = wordObject.Windows[idx];
                //    openDocs.Add(windowObject.Document.FullName);
                //}

                //return openDocs;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Adds the image at cursor position in the provided 
        /// <see cref="Word.Window"/> instance.
        /// </summary>
        /// <param name="windowObject">The window object.</param>
        /// <param name="imagePath">The image path.</param>
        /// <param name="altText">The alt text.</param>
        /// <returns><c>True</c> if image added without error, otherwise <c>false</c>.</returns>
        public static bool AddImageAtCursor(Word.Window windowObject, string imagePath, string altText)
        {
            try
            {
                if (windowObject == null)
                {
                    throw new ArgumentNullException("Word window instance is null");
                }

                if (!File.Exists(imagePath))
                {
                    throw new FileNotFoundException("Image file not found.");
                }

                var addedImage = windowObject.Application.Selection.InlineShapes.AddPicture(imagePath);
                addedImage.AlternativeText = altText;

                return true;
            }
            catch
            {
                return false;
            }
        }
    }

}
