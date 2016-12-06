using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace image_to_spreadsheet {
    class Program {
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        static void Main(string[] args) {
            DateTime timeStart = DateTime.Now;

            if (args.Length == 0) {
                DisplayUsage();
                return;
            }

            if (!File.Exists(args[0])) {
                Console.WriteLine("Error: invalid path given for source image");
                return;
            }

            string strOutputPath = (args.Length > 1) ? args[1] : args[0].Substring(0, args[0].LastIndexOf('.')) + ".xlsx";
            if (!strOutputPath.EndsWith(".xlsx")) { strOutputPath += ".xlsx"; }

            Application oXL;
            _Workbook oWB;
            _Worksheet oSheet;

            Console.WriteLine("Initializing Excel...");
            oXL = new Application();
            Workbooks oWBs = oXL.Workbooks;
            oWB = oWBs.Add(Missing.Value);
            oSheet = oWB.ActiveSheet;

            try {
                Console.WriteLine("Converting image...");
                Bitmap b = new Bitmap(args[0]);

                string strTotalPixels = (b.Width * b.Height).ToString();
                int nCurrentPixelCount = 0;

                bool bNewColor = true;
                Color c = Color.FromArgb(0, 0, 0, 0);
                Color colorToWrite = Color.FromArgb(0, 0, 0, 0);
                int startX = 0;
                int startY = 0;
                for (int x = 0; x < b.Width; x++) {
                    oSheet.Range[oSheet.Cells[1, x + 1], oSheet.Cells[1, x + 1]].ColumnWidth = 0.33;
                    for (int y = 0; y < b.Height; y++) {
                        if (x == 0) {
                            oSheet.Range[oSheet.Cells[y + 1, 1], oSheet.Cells[y + 1, 1]].RowHeight = 3;
                        }

                        c = b.GetPixel(x, y);

                        if (bNewColor) {
                            colorToWrite = c;
                            startX = x + 1;
                            startY = y + 1;
                            bNewColor = false;
                        }
                        else {
                            if (c != colorToWrite) {
                                oSheet.Range[oSheet.Cells[startY, startX], oSheet.Cells[y + 1, x + 1]].Interior.Color = ColorTranslator.ToOle(colorToWrite);
                                startY = y + 1;
                                startX = x + 1;
                                colorToWrite = c;
                            }
                        }

                        Console.Write("\r{0}/{1}", (++nCurrentPixelCount).ToString(), strTotalPixels);
                    }

                    // reset here, too
                    if (!bNewColor) {
                        oSheet.Range[oSheet.Cells[startY, startX], oSheet.Cells[b.Height + 1, x + 1]].Interior.Color = ColorTranslator.ToOle(colorToWrite);
                        bNewColor = true;
                    }
                }                

                DateTime timeEnd = DateTime.Now;
                TimeSpan timeDelta = timeEnd - timeStart;

                Console.WriteLine();
                Console.WriteLine("Saving spreadsheet...");

                oWB.SaveAs(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, strOutputPath));
                Console.WriteLine("Output saved as " + strOutputPath);
                Console.WriteLine("Done!");
                Console.WriteLine("Conversion time: " + timeDelta.TotalMilliseconds.ToString() + " milliseconds");
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
            finally {
                uint id;
                GetWindowThreadProcessId((IntPtr)oXL.Hwnd, out id);

                oWB.Close();
                oXL.Quit();

                Process[] processes = Process.GetProcessesByName("Excel");
                foreach (Process p in processes) {
                    if (p.Id == id) {
                        try {
                            p.Kill();
                        }
                        catch (Exception e) {
                            Console.WriteLine(e.Message);
                        }
                    }
                }
            }
        }

        static void DisplayUsage() {
            Console.WriteLine("Usage: image_to_spreadsheet image_path [output_path]");
            Console.WriteLine("Example: image_to_spreadsheet input.png output.xlsx");
            Console.WriteLine("Output parameter is optional; if no output parameter is provided, input parameter is used, and file extension replaced with .xlsx (input.png -> input.xlsx)");
        }
    }
}
