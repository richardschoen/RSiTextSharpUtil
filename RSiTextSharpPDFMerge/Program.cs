using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RSiTextSharpUtil;

namespace RSiTextSharpUtilTester
{
    class Program
    {
        static void Main(string[] args)
        {

            // Expected number of command line parms
            int expectedParms = 3;

            try
            {

                // Get command line parms and bail out if less than expected number of parms passed. 
                // Parms should be passed with double quotes around values and single space between parms.
                // Example Windows command line call: pgm1.exe "parm1" "parm2" "parm3"
                if (args.Length < expectedParms)
                {
                    throw new Exception(expectedParms + " required parms: [filelisttomergemerge] [outputpdffile] [replacepdf=true|false]");
                }

                // Extract parms from command line
                string parmfilelist = args[0];
                string parmoutputpdffile = args[1];
                bool parmreplace = Convert.ToBoolean(args[2]);

                // Output any log info to console. This info will get returned in STDOUT which 
                // also gets pipelined back to IBMi job if you use the MONO command to call this from 
                // your IBMi system jobs. 
                Console.WriteLine("Starting Process Now - " + System.DateTime.Now.ToString());
                Console.WriteLine("--------------------------------------------------");
                Console.WriteLine("Parameters:");
                Console.WriteLine("--------------------------------------------------");
                Console.WriteLine("Parm file list:" + parmfilelist);
                Console.WriteLine("Parm output PDF file:" + parmoutputpdffile);
                Console.WriteLine("Parm replace:" + parmreplace);

                // Make sure output PDF file does not exist
                if (File.Exists(parmoutputpdffile))
                {
                    if (parmreplace)
                    {
                        File.Delete(parmoutputpdffile);
                    }
                    else
                    {
                        throw new Exception(String.Format("Output PDF file {0} already exists and replace not selected. PDF merge process cancelled.", parmoutputpdffile));
                    }
                }

                // Instantiate PDF helper object
                PDFHelper _pdf = new PDFHelper();

                // Burst file list on semicolon
                string[] arrFileList = parmfilelist.Split(';');

                // Allocate ByteInfoArray for each image file
                ByteArrayInfo[] infoarray = new ByteArrayInfo[arrFileList.Length];

                // Load each file in to memory as a ByteArrayInfo object
                // TODO - Alter logic to load file to memory one at a time
                // This could cause issues with LARGE files
                for (int x = 0; x < arrFileList.Length; x++)
                {
                    string curfile = arrFileList[x];
                    infoarray[x] = new ByteArrayInfo(File.ReadAllBytes(curfile), curfile);
                }

                // Merge all selected files in to single PDF output file
                var rtn = _pdf.MergeIntoPDF(parmoutputpdffile, parmreplace, infoarray);
                
                // Bail if errors occurred during merge
                if (!rtn)
                {
                    throw new Exception(String.Format("PDF merge process failed. Error: {0}", _pdf.GetLastError()));
                }
                
                // Successful query. Exit program with a success message and 0 error.
                Console.WriteLine(String.Format("Merge to PDF file {0} was successful", parmoutputpdffile));
                Environment.ExitCode = 0;

            }
            catch (Exception ex)
            {
                // Write error message to console/STDOUT log and set exit code 
                // to 99 to indicate an error to the Operating System. 
                Console.WriteLine("Error: " + ex.Message);
                Environment.ExitCode = 99;
            }
            finally
            {
                // Exit the program
                Console.WriteLine("ExitCode: " + Environment.ExitCode);
                Console.WriteLine("Ending Process Now - " + System.DateTime.Now.ToString());
                Environment.Exit(Environment.ExitCode);
            }
        }
    }
}
