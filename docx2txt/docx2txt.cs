using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

class App {
    static void Main(string[] args) {

        if(args.Length != 1) {
            Console.WriteLine(@"usage: docx2txt.exe <docx>");
        } else {
            string infile = args[0];
            if(infile.EndsWith(".docx")) {

                string cwd = Directory.GetCurrentDirectory();

                string outfile = infile.Replace(".docx", ".txt");
                string infilePath = $"{cwd}\\{infile}";
                string outfilePath = $"{cwd}\\{outfile}";

                // TODO: maybe this should be VERBOSE or DEBUG option
                // Console.WriteLine($"outfile => {outfile}");
                // Console.WriteLine($"infilePath => {infilePath}");
                // Console.WriteLine($"outfilePath => {outfilePath}");
                // Environment.Exit(0);

                if(File.Exists(outfilePath)) {
                    Console.WriteLine($"output file already exists => {outfile}");
                } else {

                    try {
                        Convert(infilePath, outfilePath, WdSaveFormat.wdFormatText);
                        // Convert(infilePath, outfilePath, WdSaveFormat.wdFormatFlatXML);
                        // Convert(infilePath, outfilePath, WdSaveFormat.wdFormatOpenDocumentText);
                        // Convert(infilePath, outfilePath, WdSaveFormat.wdFormatXML);
                    } catch(Exception e) {
                        Console.WriteLine($"{e.Message}");
                    }

                    Console.WriteLine($"output file created => {outfile}");
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

