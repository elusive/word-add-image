namespace word_add_diagram
{
    using System;
    using System.IO;

    class Program
    {
        static int Main(string[] args)
        {
            var path = args[0];
            var altText = args[1];

            if (string.IsNullOrEmpty(path)) return Result.EmptyImagePath;
            if (!File.Exists(path)) return Result.ImagePathNotFound;
            if (string.IsNullOrEmpty(altText)) return Result.EmptyAltText;

            var doc = WordHelper.GetOpenDocument();
            if (doc == null) return Result.NoOpenWordDocument;

            var success = WordHelper.AddImageAtCursor(doc, path, altText);
            if (!success) return Result.FailedToAddImage;

            // done
            EndAndWait();
            return Result.Success;
        }

        static void EndAndWait(string msg = "Press Enter to exit...")
        {
            Console.WriteLine(msg);
            Console.ReadLine();
        }
    }

    /// <summary>
    /// Result values for the command line application.
    /// </summary>
    public class Result
    {
        public const int Success = 0;
        public const int EmptyImagePath = 1;
        public const int ImagePathNotFound = 2;
        public const int EmptyAltText = 3;
        public const int NoOpenWordDocument = 4;
        public const int FailedToAddImage = 5;
    }
}
