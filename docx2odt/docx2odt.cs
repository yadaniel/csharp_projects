using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

class App {
    static void Main(string[] args) {

        // foreach(string arg in args.Skip(1)) {
        // foreach(string arg in args) {
        //     Console.WriteLine($"{arg}");
        // }
        // Console.WriteLine();

        if(args.Length == 0) {
            // TODO: maybe add option FORCE to overwrite existing outzipPath
            Console.WriteLine(@"usage: docx2odt.exe <docx> [<docx>]");
        } else {
            foreach(string infile in args) {
                if(infile.EndsWith(".docx")) {

                    string cwd = Directory.GetCurrentDirectory();

                    // string outfile = infile.Replace(".docx", ".pdf");
                    // Convert($"{cwd}\\{infile}", $"{cwd}\\{outfile}", WdSaveFormat.wdFormatPDF);

                    string zipfile = infile.Replace(".docx", ".zip");
                    string outfile = infile.Replace(".docx", ".odt");
                    string infilePath = $"{cwd}\\{infile}";
                    string outfilePath = $"{cwd}\\{outfile}";
                    string outzipPath = $"{cwd}\\{zipfile}";
                    string unzippedDocumentFolder = outzipPath.Replace(".zip", "");

                    // TODO: maybe this should be VERBOSE or DEBUG option
                    // Console.WriteLine($"zipfile => {zipfile}");
                    // Console.WriteLine($"outfile => {outfile}");
                    // Console.WriteLine($"infilePath => {infilePath}");
                    // Console.WriteLine($"outfilePath => {outfilePath}");
                    // Console.WriteLine($"outzipPath => {outzipPath}");
                    // Console.WriteLine($"unzippedDocumentFolder => {unzippedDocumentFolder}");
                    // Environment.Exit(0);

                    if(Directory.Exists(unzippedDocumentFolder)) {
                        Console.WriteLine($"output folder for document already exists => {unzippedDocumentFolder}");
                    } else {

                        Convert(infilePath, outfilePath, WdSaveFormat.wdFormatOpenDocumentText);
                        File.Copy(outfilePath, outzipPath, overwrite: true);
                        ZipFile.ExtractToDirectory(outzipPath, unzippedDocumentFolder);

                        try {
                            File.Delete(outfilePath);
                            File.Delete(outzipPath);
                        } catch(Exception e) {
                            Console.WriteLine($"File.Delete failed with {e.Message}");
                        }

                        Console.WriteLine($"output folder for document created => {unzippedDocumentFolder}");
                    }
                }
            }
        }
        //Console.ReadKey();
    }

    // Convert method
    public static void Convert(string input, string output, WdSaveFormat format) {
        // Create an instance of Word.exe
        _Application oWord = new Word.Application {

            // Make this instance of word invisible (Can still see it in the taskmgr).
            Visible = false
        };

        // Interop requires objects.
        object oMissing = System.Reflection.Missing.Value;
        object isVisible = true;
        object readOnly = true;     // Does not cause any word dialog to show up
        //object readOnly = false;  // Causes a word object dialog to show at the end of the conversion
        object oInput = input;
        object oOutput = output;
        object oFormat = format;

        // Load a document into our instance of word.exe
        _Document oDoc = oWord.Documents.Open(
                             ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                             ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                             ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                         );

        // Make this document the active document.
        oDoc.Activate();

        // Save this document using Word
        oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                   );

        // Always close Word.exe.
        oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
    }
}

