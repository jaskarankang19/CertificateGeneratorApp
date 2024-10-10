using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System.IO;
using Xceed.Words.NET;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using DocumentFormat.OpenXml;
//using System.Reflection.Metadata;

public class GenerateRoCertificateModel : PageModel
{
    [BindProperty]
    public IFormFile UploadedFile { get; set; }
    [BindProperty]
    public string CustomFolderPath { get; set; } // This will store the user's folder input
    [BindProperty]
    public string RoNumber { get; set; } // This will store the user's folder input


    public string Message { get; set; }

    public IActionResult OnPost()
    {

        if (UploadedFile != null && !string.IsNullOrWhiteSpace(CustomFolderPath))
        {

            if (!Directory.Exists(CustomFolderPath))
            {
                Directory.CreateDirectory(CustomFolderPath);
            }

            // Generate a unique file name with the same extension as the uploaded file
            var uploadsFolder = Path.Combine(Path.GetTempPath(), "ExcelUploads");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            // Use the original file extension and generate a unique name
            var fileExtension = Path.GetExtension(UploadedFile.FileName);
            var uniqueFileName = Guid.NewGuid().ToString() + fileExtension;
            var filePath = Path.Combine(uploadsFolder, uniqueFileName);

            // Save the uploaded file
            // Save the uploaded file in the custom folder
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                UploadedFile.CopyTo(stream);
            }

            // Process the Excel file and generate certificates
            GenerateCertificates(filePath, CustomFolderPath);

            
        }
        else
        {
            Message = "Please upload a valid Excel file.";
        }

        return Page();
    }


    private void GenerateCertificates(string excelFilePath, string uploadsFolder)
    {
        //string certificatesOutputFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Certificates");
        string certificatesOutputFolder = uploadsFolder;
        //string roNumber = RoNumber;

        if (!Directory.Exists(certificatesOutputFolder))
        {
            Directory.CreateDirectory(certificatesOutputFolder);
        }

        // Read the Excel file
        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().Where(x => x.Cell(11).GetString() == RoNumber);
            if (rows.Count() == 0)
            {

                Message = "This RO does not have any data.";
                return;
            }
            foreach (var row in rows)
            {
                string roName = row.Cell(9).GetString();
                string village = row.Cell(3).GetString();
                string name = row.Cell(5).GetString();
                string fatherName = row.Cell(6).GetString();
                string block = row.Cell(2).GetString();
                string type = row.Cell(7).GetString();
                string wardNumber = row.Cell(10).GetString();
                string roNumber = row.Cell(11).GetString();
                string district = row.Cell(12).GetString();

                GenerateCertificate(roName, village, name, fatherName, block, type, wardNumber, roNumber, district, certificatesOutputFolder);

                Message = "Certificates generated successfully!";
            }
        }
    }

    // The same GenerateCertificate method from your console app
    private void GenerateCertificate(string roName, string village, string name, string fatherName, string block, string type, string wardNumber, string roNumber, string district, string outputFolder)
    {
        string filePath = Path.Combine(outputFolder, $"{village}_{type}_{name}_Certificate.docx");

        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();

            // Add the header (as per your image)
            body.Append(CreateParagraph("ਗ੍ਰਾਮ ਪੰਚਾਇਤ ਚੋਣ-2024", true, "AnmolLipi", 16, justification: "center"));
            body.Append(CreateParagraph("(ਬਿਨ੍ਹਾਂ ਮੁਕਾਬਲਾ ਚੋਣ ਦਾ ਨਤੀਜਾ ਘੋਸ਼ਿਤ ਕਰਨ ਬਾਰੇ)", true, "AnmolLipi", 16, justification: "center"));
            body.Append(CreateParagraph("ਪ੍ਰਮਾਣ ਪੱਤਰ", true, "AnmolLipi", 14, justification: "center", underline: true));
            if (type == "Sarpanch")
            {
                body.Append(CreateParagraph("(ਸਰਪੰਚ)", true, "AnmolLipi", 14, justification: "center"));
            }
            if (type == "Panch")
            {
                body.Append(CreateParagraph("(ਪੰਚ)", true, "AnmolLipi", 14, justification: "center"));
            }


            // Add space
            body.Append(new Paragraph(new Run(new Text(" "))));

            // Create the dynamic certificate content with underlined and bold dynamic values
            Paragraph certificateParagraph = new Paragraph();
            Paragraph signatureParagraph = new Paragraph();

            // Append the static part first
            certificateParagraph.Append(CreateRun("ਮੈਂ ", false, "AnmolLipi", 12));
            certificateParagraph.Append(CreateRun(roName, true, "AnmolLipi", 12, underline: true)); // roName - Underlined and Bold
            //certificateParagraph.Append(CreateRun(" ਸਹਾਇਕ ਰਿਟਰਨਿੰਗ ਅਫ਼ਸਰ ਇਹ ਤਸਦੀਕ ਕਰਦਾ ਹਾਂ ਕਿ ਮਿਤੀ 15.10.2024 ਨੂੰ ਨਿਸ਼ਚਿਤ ਗ੍ਰਾਮ ਪੰਚਾਇਤ ਚੋਣਾਂ ਸਬੰਧੀ ਬਿਨ੍ਹਾ ਮੁਕਾਬਲਾ ਸ਼੍ਰੀ ", false, "AnmolLipi", 12));
            certificateParagraph.Append(CreateRun(" ਰਿਟਰਨਿੰਗ ਅਫ਼ਸਰ / ਸਹਾਇਕ ਰਿਟਰਨਿੰਗ ਅਫ਼ਸਰ ਇਹ ਤਸਦੀਕ ਕਰਦਾ ਹਾਂ ਕਿ ਮਿਤੀ", false, "AnmolLipi", 12));
            certificateParagraph.Append(CreateRun(" 15-10-2024 ", true, "AnmolLipi", 12, underline: true));
            certificateParagraph.Append(CreateRun(" ਨੂੰ ਨਿਸ਼ਚਿਤ ਗ੍ਰਾਮ ਪੰਚਾਇਤ ਚੋਣਾਂ ਸਬੰਧੀ ਬਿਨ੍ਹਾ ਮੁਕਾਬਲਾ ਸ਼੍ਰੀ / ਸ਼੍ਰੀਮਤੀ ", false, "AnmolLipi", 12));
            certificateParagraph.Append(CreateRun(name, true, "AnmolLipi", 12, underline: true)); // name - Underlined and Bold
            certificateParagraph.Append(CreateRun(", ਪਿਤਾ ਸ਼੍ਰੀ ", false, "AnmolLipi", 12));

            if (type == "Sarpanch")
            {
                certificateParagraph.Append(CreateRun(fatherName, true, "AnmolLipi", 12, underline: true)); // fatherName - Underlined and Bold
                certificateParagraph.Append(CreateRun(" ਨੂੰ ਗ੍ਰਾਮ ਪੰਚਾਇਤ ", false, "AnmolLipi", 12));
                certificateParagraph.Append(CreateRun(village, true, "AnmolLipi", 12, underline: true)); // village - Underlined and Bold
                certificateParagraph.Append(CreateRun(" ਦਾ ਬਤੌਰ ਸਰਪੰਚ ਜੇਤੂ ਘੋਸ਼ਿਤ ਕਰਦਾ ਹਾਂ।", false, "AnmolLipi", 12));
                // Append the certificate content to the body
                body.Append(certificateParagraph.CloneNode(true));

                // Add dynamic content
                body.Append(CreateParagraph("ਇਸ ਨੂੰ ਬਤੌਰ ਸਰਪੰਚ ਘੋਸ਼ਿਤ ਕਰਨ ਉਪਰੰਤ ਇਹ ਪ੍ਰਮਾਣ ਪੱਤਰ ਜਾਰੀ ਕੀਤਾ ਜਾਂਦਾ ਹੈ।\n", false, "AnmolLipi", 12, justification: "center"));
            }
            if (type == "Panch")
            {
                certificateParagraph.Append(CreateRun(fatherName, true, "AnmolLipi", 12, underline: true)); // fatherName - Underlined and Bold
                certificateParagraph.Append(CreateRun(" ਨੂੰ ਵਾਰਡ ਨੰ: ", true, "AnmolLipi", 12));
                certificateParagraph.Append(CreateRun(wardNumber, true, "AnmolLipi", 12, underline: true));// fatherName - Underlined and Bold
                certificateParagraph.Append(CreateRun("    ਗ੍ਰਾਮ ਪੰਚਾਇਤ ", false, "AnmolLipi", 12));
                certificateParagraph.Append(CreateRun(village, true, "AnmolLipi", 12, underline: true)); // village - Underlined and Bold
                certificateParagraph.Append(CreateRun(" ਦਾ ਬਤੌਰ ਪੰਚ ਜੇਤੂ ਘੋਸ਼ਿਤ ਕਰਦਾ ਹਾਂ।", false, "AnmolLipi", 12));
                // Append the certificate content to the body
                body.Append(certificateParagraph.CloneNode(true));

                // Add dynamic content
                body.Append(CreateParagraph("ਇਸ ਨੂੰ ਬਤੌਰ ਪੰਚ ਘੋਸ਼ਿਤ ਕਰਨ ਉਪਰੰਤ ਇਹ ਪ੍ਰਮਾਣ ਪੱਤਰ ਜਾਰੀ ਕੀਤਾ ਜਾਂਦਾ ਹੈ।\n", false, "AnmolLipi", 12, justification: "center"));
            }




            // Add empty lines
            body.Append(CreateEmptyLine());
            body.Append(CreateEmptyLine());
            body.Append(CreateEmptyLine());

            // Add signature line

            body.Append(CreateParagraph("ਸਥਾਨ....................\t\t\t\t\t\tਹਸਤਾਖਰ..............................", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph("ਸਮਾਂ....................\t\t\t\t\t\t\t\tਰਿਟਰਨਿੰਗ ਅਫ਼ਸਰ / ", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph("\t\t\t\t\t\t\t\t\tਸਹਾਇਕ ਰਿਟਰਨਿੰਗ ਅਫ਼ਸਰ -,", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph("\t\t\t\t\t\t\t\t\tਕਮ -ਪ੍ਰੀਜਾਇਡਿੰਗ ਅਫ਼ਸਰ", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph($"\t\t\t\t\t\t\t\t\tਗ੍ਰਾਮ ਪੰਚਾਇਤ {village}", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph($"\t\t\t\t\t\t\t\t\tਬਲਾਕ {block}", false, "AnmolLipi", 12, justification: "left"));
            body.Append(CreateNoSpaceParagraph($"\t\t\t\t\t\t\t\t\tਜਿਲ੍ਹਾ {district}", false, "AnmolLipi", 12, justification: "left"));


            // Append body to the document
            mainPart.Document.Append(body.CloneNode(true));
            mainPart.Document.Save();

            // Create footer and add footer number to the bottom left
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            Footer footer = new Footer();

            // Add the footer number to the left corner
            Paragraph footerParagraph = CreateParagraph(roNumber.ToString(), false, "AnmolLipi", 12, justification: "left");

            footer.Append(footerParagraph);

            // Attach the footer part to the document
            footerPart.Footer = footer;
            SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<SectionProperties>();
            if (sectionProperties == null)
            {
                sectionProperties = new SectionProperties();
                mainPart.Document.Body.Append(sectionProperties.CloneNode(true));
            }

            FooterReference footerReference = new FooterReference() { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) };
            sectionProperties.Append(footerReference);
            //mainPart.Document.Append(body);
            mainPart.Document.Save();
        }
    }
    static Paragraph CreateEmptyLine()
    {
        return new Paragraph(new Run(new Text(" ")));
    }
    static Paragraph CreateNoSpaceParagraph(string text, bool bold, string fontName = "AnmolLipi", int fontSize = 12, string justification = "left", bool underline = false)
    {
        Run run = CreateRun(text, bold, fontName, fontSize, underline);

        Paragraph paragraph = new Paragraph();
        ParagraphProperties paragraphProperties = new ParagraphProperties();

        if (justification == "center")
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });
        }
        else if (justification == "right")
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
        }
        else
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Left });
        }
        // Set spacing between lines (remove space after the paragraph)
        SpacingBetweenLines spacingBetweenLines = new SpacingBetweenLines()
        {
            After = "0", // Remove space after the paragraph
            Before = "0"  // (Optional) Remove space before the paragraph, if needed
        };

        paragraphProperties.Append(spacingBetweenLines);
        paragraph.Append(paragraphProperties);
        paragraph.Append(run);

        return paragraph;
    }
    // Helper method to create a run with optional bold and underline
    static Run CreateRun(string text, bool bold, string fontName = "AnmolLipi", int fontSize = 12, bool underline = false)
    {
        RunProperties runProperties = new RunProperties();

        if (bold)
        {
            runProperties.Append(new Bold());
        }

        if (underline)
        {
            runProperties.Append(new Underline() { Val = UnderlineValues.Single });
        }

        runProperties.Append(new RunFonts() { Ascii = fontName, HighAnsi = fontName, ComplexScript = fontName });
        runProperties.Append(new FontSize() { Val = (fontSize * 2).ToString() });

        Run run = new Run();
        run.Append(runProperties);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        return run;
    }

    // Helper method to create formatted paragraphs
    static Paragraph CreateParagraph(string text, bool bold, string fontName = "AnmolLipi", int fontSize = 12, string justification = "left", bool underline = false)
    {
        Run run = CreateRun(text, bold, fontName, fontSize, underline);

        Paragraph paragraph = new Paragraph();
        ParagraphProperties paragraphProperties = new ParagraphProperties();

        if (justification == "center")
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });
        }
        else if (justification == "right")
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
        }
        else
        {
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Both });
        }

        paragraph.Append(paragraphProperties);
        paragraph.Append(run);

        return paragraph;
    }
}
