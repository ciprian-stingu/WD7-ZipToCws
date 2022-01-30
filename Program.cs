using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;

namespace WD7_ZipToCws
{
    class Program
    {
        private const string RUN_GCODE_FILE_NAME = "run.gcode";
        private const string MANIFEST_XML_FILE_NAME = "manifest.xml";
        private const string GCODE_FILE_NAME = "1.gcode";

        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Please provide a zip file.");
                Environment.Exit(0);
            }

            var zipFileName = args[0];

            if (Path.GetExtension(zipFileName).ToLower() != ".zip")
            {
                Console.WriteLine("Not a zip file.");
                Environment.Exit(0);
            }

            Console.Write($"Opening {zipFileName} for reading...");
            ZipArchive zipFile = null;
            try
            {
                zipFile = ZipFile.OpenRead(zipFileName);
            } 
            catch (Exception e)
            {
                Console.WriteLine("Fail to load the zip file.");
                Console.WriteLine(e.Message);
                Environment.Exit(0);
            }
            Console.WriteLine("done");

            Console.Write($"Opening {RUN_GCODE_FILE_NAME}...");
            ZipArchiveEntry runGcodeFile = null;
            try
            {
                runGcodeFile = zipFile.GetEntry(RUN_GCODE_FILE_NAME);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Fail to read {RUN_GCODE_FILE_NAME} from zip.");
                Console.WriteLine(e.Message);
                Environment.Exit(0);
            }

            if (runGcodeFile == null)
            {
                zipFile.Dispose();
                Console.WriteLine("Not a valid zip file.");
                Environment.Exit(0);
            }
            Console.WriteLine("done");

            Console.Write($"Reading {RUN_GCODE_FILE_NAME}...");
            Stream runGCodeFileStream = null;
            try
            {
                runGCodeFileStream = runGcodeFile.Open();
            }
            catch (Exception e)
            {
                zipFile.Dispose();
                Console.WriteLine($"Fail to open {RUN_GCODE_FILE_NAME} from zip.");
                Console.WriteLine(e.Message);
                Environment.Exit(0);
            }

            List<string> runGCodeLines = new List<string>();
            using (StreamReader reader = new StreamReader(runGCodeFileStream))
            {
                while (!reader.EndOfStream)
                {
                    runGCodeLines.Add(reader.ReadLine());
                }
            }
            runGCodeFileStream.Dispose();
            if (runGCodeLines.Count == 0)
            {
                zipFile.Dispose();
                Console.WriteLine($"Empty {RUN_GCODE_FILE_NAME} in zip.");
                Environment.Exit(0);
            }
            Console.WriteLine("done");

            int totalLayers = GetIntValueFromGCode(runGCodeLines, ";totalLayer:");
            if (totalLayers == 0)
            {
                zipFile.Dispose();
                Console.WriteLine($"No layers found in {RUN_GCODE_FILE_NAME}.");
                Environment.Exit(0);
            }

            Console.Write($"Creating {MANIFEST_XML_FILE_NAME} data...");
            string manifestXmlData = CreateManifestFileData(totalLayers);
            if (manifestXmlData.Length == 0)
            {
                zipFile.Dispose();
                Console.WriteLine($"Can't create data for {MANIFEST_XML_FILE_NAME}.");
                Environment.Exit(0);
            }
            Console.WriteLine("done");

            Console.Write($"Creating {GCODE_FILE_NAME} data...");
            string gCodeFileData = CreateGcodeFileData(runGCodeLines, totalLayers);
            if (gCodeFileData.Length == 0)
            {
                zipFile.Dispose();
                Console.WriteLine($"Can't create data for {GCODE_FILE_NAME}.");
                Environment.Exit(0);
            }
            Console.WriteLine("done");

            var cwsFile = Path.GetFileNameWithoutExtension(zipFileName) + ".cws";
            Console.WriteLine($"Preparing to create {cwsFile}...");
            using (var memoryStream = new MemoryStream())
            {
                using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
                {
                    Console.WriteLine($"Adding {GCODE_FILE_NAME}...");
                    AddTextFileInArchive(archive, GCODE_FILE_NAME, gCodeFileData);
                    Console.WriteLine($"Adding {MANIFEST_XML_FILE_NAME}...");
                    AddTextFileInArchive(archive, MANIFEST_XML_FILE_NAME, manifestXmlData);

                    for(var i = 1; i <= totalLayers; i++)
                    {
                        Console.WriteLine($"Adding {i}.png...");
                        CopyPngFileFromArchive(zipFile, archive, $"{i}.png", "slice" + (i - 1).ToString("0000") + ".png");
                    }
                    //Console.WriteLine("Adding preview.png...");
                    //CopyPngFileFromArchive(zipFile, archive, "preview.png", "preview.png");
                    //Console.WriteLine("Adding preview_mini.png...");
                    //CopyPngFileFromArchive(zipFile, archive, "preview_cropping.png", "preview_mini.png");
                }

                Console.Write($"Writing {cwsFile} to disk...");
                using (var fileStream = new FileStream(Path.GetFileNameWithoutExtension(zipFileName) + ".cws", FileMode.Create))
                {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(fileStream);
                }
                Console.WriteLine("done");
            }

            zipFile.Dispose();

        }

        private static bool AddTextFileInArchive(ZipArchive zipArchive, string fileName, string fileData)
        {
            var utf8Encoding = new UTF8Encoding(false);

            var file = zipArchive.CreateEntry(fileName, CompressionLevel.Optimal);
            file.ExternalAttributes = file.ExternalAttributes | (Convert.ToInt32("664", 8) << 16);
            using (var fileStream = file.Open())
            using (var fileToCompressStream = new MemoryStream(utf8Encoding.GetBytes(fileData)))
            {
                fileToCompressStream.CopyTo(fileStream);
            }

            return true;
        }

        private static bool CopyPngFileFromArchive(ZipArchive srcArchive, ZipArchive dstArchive, string srcFileName, string dstFileName)
        {
            ZipArchiveEntry pngFile = null;
            try
            {
                pngFile = srcArchive.GetEntry(srcFileName);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Fail to read {srcFileName} from zip.");
                Console.WriteLine(e.Message);
                return false;
            }

            if (pngFile == null)
            {
                Console.WriteLine($"{srcFileName} not found in zip.");
                return false;
            }

            Stream pngFileStream = null;
            try
            {
                pngFileStream = pngFile.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine($"Fail to open {srcFileName} from zip.");
                Console.WriteLine(e.Message);
                return false;
            }

            var file = dstArchive.CreateEntry(dstFileName, CompressionLevel.Optimal);
            file.ExternalAttributes = file.ExternalAttributes | (Convert.ToInt32("664", 8) << 16);
            using (var fileStream = file.Open())
            ConvertPngTo32Bit(pngFileStream, fileStream);

            return true;
        }


        private static void ConvertPngTo32Bit(Stream sourceFile, Stream destinationFile)
        {
            using (var image = new Bitmap(sourceFile))
            {
                Bitmap bmp = new Bitmap(image.Width, image.Height, PixelFormat.Format32bppArgb);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.DrawImage(image, new Rectangle(new Point(), image.Size), new Rectangle(new Point(), image.Size), GraphicsUnit.Pixel);
                    bmp.Save(destinationFile, ImageFormat.Png);
                }
            }
        }

        private static string CreateGcodeFileData(List<string> gCodeLines, int totalLayers)
        {
            double layerHeight = GetDoubleValueFromGCode(gCodeLines, ";layerHeight:");

            int bottomLayerExposureTime = Convert.ToInt32(GetDoubleValueFromGCode(gCodeLines, ";bottomLayerExposureTime:") * 1000);
            int bottomLightOffTime = Convert.ToInt32(GetDoubleValueFromGCode(gCodeLines, ";bottomLightOffTime:") * 1000);
            double bottomLayerLiftSpeed = GetDoubleValueFromGCode(gCodeLines, ";bottomLayerLiftSpeed:");
            double bottomLayerLiftHeight = GetDoubleValueFromGCode(gCodeLines, ";bottomLayerLiftHeight:");
            double bottomDropHeight = bottomLayerLiftHeight - layerHeight;

            int normalExposureTime = Convert.ToInt32(GetDoubleValueFromGCode(gCodeLines, ";normalExposureTime:") * 1000);
            int lightOffTime = Convert.ToInt32(GetDoubleValueFromGCode(gCodeLines, ";lightOffTime:") * 1000);
            double normalLayerLiftSpeed = GetDoubleValueFromGCode(gCodeLines, ";normalLayerLiftSpeed:");
            double normalLayerLiftHeight = GetDoubleValueFromGCode(gCodeLines, ";normalLayerLiftHeight:");
            double normalDropHeight = normalLayerLiftHeight - layerHeight;

            double normalDropSpeed = GetDoubleValueFromGCode(gCodeLines, ";normalDropSpeed:");
            int bottomLayers = GetIntValueFromGCode(gCodeLines, ";bottomLayerCount:");

            StringBuilder gCodeData = new StringBuilder();

            gCodeData.Append(";****Build and Slicing Parameters****\n");
            gCodeData.Append(GetOriginalCommentSection(gCodeLines));
            gCodeData.Append($";Number of Layers = {totalLayers}\n");
            gCodeData.Append($";Number of Slices = {totalLayers}\n");
            gCodeData.Append(";********** Header Start ********\n");
            gCodeData.Append("G28 ; Home\n");
            gCodeData.Append("G21 ; Set units to be mm\n");
            gCodeData.Append("G91 ; Relative Positioning\n");
            gCodeData.Append("M17 ; Enable motors\n");

            for(var currentLayer = 0; currentLayer < totalLayers; currentLayer++)
            {
                gCodeData.Append($";********** Pre-Slice {currentLayer} ********\n");
                gCodeData.Append("G4 P0 ; Make sure any previous relative moves are complete\n");
                gCodeData.Append($";********** Layer {currentLayer} ******\n");
                gCodeData.Append($";<Slice> {currentLayer}\n");
                gCodeData.Append("M106 S255 ; UV on\n");
                gCodeData.Append(";<Delay> " + (currentLayer < bottomLayers ? bottomLayerExposureTime : normalExposureTime) + "\n");
                gCodeData.Append("M106 S0 ; UV off\n");
                gCodeData.Append(";<Slice> Blank\n");
                gCodeData.Append(";<Delay> " + (currentLayer < bottomLayers ? bottomLightOffTime : lightOffTime)  + "\n");
                gCodeData.Append($";********** Lift Sequence {currentLayer} ******\n");
                gCodeData.Append("G1 Z" + (currentLayer < bottomLayers ? bottomLayerLiftHeight : normalLayerLiftHeight) + " F" + (currentLayer < bottomLayers ? bottomLayerLiftSpeed : normalLayerLiftSpeed) + "\n");
                gCodeData.Append("G4 P0 ; Wait for lift rise to complete\n");
                gCodeData.Append("G1 Z-" + (currentLayer < bottomLayers ? bottomDropHeight : normalDropHeight) + " F" + normalDropSpeed + "\n");
            }

            gCodeData.Append(";********** Footer ******\n");
            gCodeData.Append("M106 S0          ; UV off\n");
            gCodeData.Append("G4 P0            ; wait for last lift to complete\n");
            gCodeData.Append("G1 Z40.0 F150.0  ; lift model clear of resin\n");
            gCodeData.Append("G4 P0            ; sync\n");
            gCodeData.Append("M18              ;Disable Motors\n");
            gCodeData.Append(";<Completed>\n");

            return gCodeData.ToString();
        }

        private static string GetOriginalCommentSection(List<string> gCodeLines)
        {
            StringBuilder comments = new StringBuilder();
            foreach (var line in gCodeLines)
            {
                if (line.Contains(";START_GCODE_BEGIN"))
                {
                    break;
                }
                comments.Append(";*");
                comments.Append(line.Substring(1));
                comments.Append("\n");
            }
            return comments.ToString();
        }

        private static int GetIntValueFromGCode(List<string> gCodeLines, string key)
        {
            int value = 0;
            foreach (var line in gCodeLines)
            {
                if (line.Contains(key))
                {
                    value = int.Parse(line.Substring(line.IndexOf(":") + 1));
                    break;
                }
            }
            return value;
        }

        private static double GetDoubleValueFromGCode(List<string> gCodeLines, string key)
        {
            double value = 0.0;
            foreach (var line in gCodeLines)
            {
                if (line.Contains(key))
                {
                    value = double.Parse(line.Substring(line.IndexOf(":") + 1));
                    break;
                }
            }
            return value;
        }

        private static string CreateManifestFileData(int totalLayers)
        {
            StringBuilder xmlData = new StringBuilder();

            xmlData.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xmlData.Append("<manifest FileVersion=\"1\">\n");
            xmlData.Append("  <GCode>\n");
            xmlData.Append($"    <name>{GCODE_FILE_NAME}</name>\n");
            xmlData.Append("  </GCode>\n");
            //xmlData.Append("  <ScenePreview>\n");
            //xmlData.Append("    <Default>preview.png</Default>\n");
            //xmlData.Append("  </ScenePreview>\n");
            xmlData.Append("  <Slices>\n");
            for (var i = 0; i < totalLayers; i++)
            {
                xmlData.Append("    <Slice>\n");
                xmlData.Append("      <name>slice" + i.ToString("0000") + ".png</name>\n");
                xmlData.Append("    </Slice>\n");
            }
            xmlData.Append("  </Slices>\n");
            xmlData.Append("</manifest>\n");

            return xmlData.ToString();
        }
    }
}
