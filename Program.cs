namespace word_add_diagram
{
    using System;
    using System.CommandLine;
    using System.CommandLine.NamingConventionBinder;
    using System.CommandLine.Parsing;
    using System.IO;

    class Program
    {
        static int Main(string path = "~/downloads/test.png", string altText = "")
        {
            var pathArgument = new Argument<string>("path", "The path to the image to add to our document.");
            var altTextArgument = new Argument<string>("alt-text", "The alt text for the image we will add to our document.");

            var cmd = new RootCommand {
                    pathArgument,
                    altTextArgument
                };
            cmd.Description = "The base command for this utility, adding an image with alt text to the open word document.";
            cmd.Handler = CommandHandler.Create<string, string>(AddWordImageCommandHandler);

            return cmd.Invoke(new [] {path, altText});
        }

        static void EndAndWait(string msg = "Press Enter to exit...")
        {
            Console.WriteLine(msg);
            Console.ReadLine();
        }

        private static int AddWordImageCommandHandler(string imagePath, string altText)
        {
            if (string.IsNullOrEmpty(imagePath)) return Result.EmptyImagePath;
            if (!File.Exists(imagePath)) return Result.ImagePathNotFound;
            if (string.IsNullOrEmpty(altText)) return Result.EmptyAltText;

            var doc = WordHelper.GetOpenDocument();
            if (doc == null) return Result.NoOpenWordDocument;

            var success = WordHelper.AddImageAtCursor(doc, imagePath, altText);
            if (!success) return Result.FailedToAddImage;

            return Result.Success;
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
