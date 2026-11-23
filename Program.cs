using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

internal static class Program
{
    private static void Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.WriteLine("Usage: PptxMarkdown <presentation.pptx>");
            return;
        }

        string inputPath = Path.GetFullPath(args[0]);
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"File not found: {inputPath}");
            return;
        }

        string directory = Path.GetDirectoryName(inputPath)!;
        string name = Path.GetFileNameWithoutExtension(inputPath);
        string markdownPath = Path.Combine(directory, name + ".md");
        string outputPath = Path.Combine(directory, name + "-edited.pptx");

        try
        {
            ExtractText(inputPath, markdownPath);

            Console.WriteLine($"Extracted text to:\n{markdownPath}");
            Console.WriteLine();
            Console.WriteLine("Edit and save the Markdown file, then press any key here to import it.");
            Console.ReadKey(intercept: true);

            File.Copy(inputPath, outputPath, overwrite: true);
            int updated = ImportText(outputPath, markdownPath);

            Console.WriteLine($"\nUpdated {updated} paragraph(s).");
            Console.WriteLine($"Created:\n{outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"\nError: {ex.Message}");
        }
    }

    private static void ExtractText(string pptxPath, string markdownPath)
    {
        using var doc = PresentationDocument.Open(pptxPath, false);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slideIds = presentationPart.Presentation.SlideIdList?.ChildElements
            ?? throw new InvalidOperationException("Slide list not found.");

        using var writer = new StreamWriter(markdownPath, false, new UTF8Encoding(false));

        for (int slideIndex = 0; slideIndex < slideIds.Count; slideIndex++)
        {
            if (slideIndex > 0)
                writer.WriteLine("\n---\n");

            var slideId = (P.SlideId)slideIds[slideIndex];
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;

            if (shapeTree == null)
                continue;

            foreach (var shape in shapeTree.Descendants<P.Shape>())
            {
                uint? shapeId = shape.NonVisualShapeProperties?
                    .NonVisualDrawingProperties?.Id?.Value;

                var textBody = shape.TextBody;
                if (shapeId == null || textBody == null)
                    continue;

                var paragraphs = textBody.Elements<A.Paragraph>().ToList();

                for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
                {
                    string text = string.Concat(
                        paragraphs[paragraphIndex].Descendants<A.Text>()
                            .Select(t => t.Text));

                    writer.WriteLine(
                        $"<!-- slide:{slideIndex + 1};shape:{shapeId};paragraph:{paragraphIndex} -->");
                    writer.WriteLine(text);
                    writer.WriteLine();
                }
            }
        }
    }

    private static int ImportText(string pptxPath, string markdownPath)
{
    string markdown = File.ReadAllText(markdownPath, Encoding.UTF8);
    var entries = new Dictionary<(int Slide, uint Shape, int Paragraph), string>();
    string[] lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

    var markerPattern = new Regex(
        @"^<!--\s*slide:(\d+);shape:(\d+);paragraph:(\d+)\s*-->$");

    for (int i = 0; i < lines.Length; i++)
    {
        var match = markerPattern.Match(lines[i].Trim());
        if (!match.Success)
            continue;

        var key = (
            int.Parse(match.Groups[1].Value),
            uint.Parse(match.Groups[2].Value),
            int.Parse(match.Groups[3].Value));

        entries[key] = i + 1 < lines.Length ? lines[i + 1] : "";
    }

    int updated = 0;

    using (var doc = PresentationDocument.Open(pptxPath, true))
    {
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slideIds = presentationPart.Presentation.SlideIdList?.ChildElements
            ?? throw new InvalidOperationException("Slide list not found.");

        foreach (var entry in entries)
        {
            int slideNumber = entry.Key.Slide;
            uint shapeId = entry.Key.Shape;
            int paragraphIndex = entry.Key.Paragraph;

            if (slideNumber < 1 || slideNumber > slideIds.Count)
                continue;

            var slideId = (P.SlideId)slideIds[slideNumber - 1];
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;

            if (shapeTree == null)
                continue;

            P.Shape? targetShape = null;

            foreach (var candidate in shapeTree.Descendants<P.Shape>())
            {
                if (candidate.NonVisualShapeProperties?
                        .NonVisualDrawingProperties?.Id?.Value == shapeId)
                {
                    targetShape = candidate;
                    break;
                }
            }

            var paragraphs = targetShape?.TextBody?.Elements<A.Paragraph>().ToList();

            if (paragraphs == null ||
                paragraphIndex < 0 ||
                paragraphIndex >= paragraphs.Count)
                continue;

            var textElements = paragraphs[paragraphIndex]
                .Descendants<A.Text>()
                .ToList();

            // Don't add new run structures to empty paragraphs in this first version.
            if (textElements.Count == 0)
                continue;

            // Change the first existing text node; clear any remaining text nodes.
            textElements[0].Text = entry.Value;

            for (int i = 1; i < textElements.Count; i++)
                textElements[i].Text = "";

            slidePart.Slide.Save();
            updated++;
        }
    }

    // Validate the saved file. A nonzero count means it needs investigation
    // before opening it in PowerPoint.
    using (var check = PresentationDocument.Open(pptxPath, false))
    {
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        var errors = validator.Validate(check).Take(10).ToList();

        if (errors.Count > 0)
        {
            string details = string.Join(
                Environment.NewLine,
                errors.Select(e => $"{e.Description} ({e.Path?.XPath})"));

            throw new InvalidOperationException(
                $"The output failed Open XML validation:{Environment.NewLine}{details}");
        }
    }

    return updated;
}


    private static void ReplaceParagraphText(A.Paragraph paragraph, string newText)
    {
        var textElements = paragraph.Descendants<A.Text>().ToList();

        if (textElements.Count == 0)
        {
            paragraph.Append(new A.Run(new A.Text(newText)));
            return;
        }

        // Keep the first text run (and its formatting); clear the remaining
        // runs so old text is not left behind.
        textElements[0].Text = newText;

        for (int i = 1; i < textElements.Count; i++)
            textElements[i].Text = "";
    }
}
