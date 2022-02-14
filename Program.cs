using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CreateExcelCombobox
{
    class Program
    {
        static eexam_utestEntities _context = new eexam_utestEntities();
        static void Main(string[] args)
        {
            EditExel( CopyFile());

        }

        

        // copy file
        static string CopyFile()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"files");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filePath = Path.Combine(path, "Bulk_Questions_Template.xlsx");
            string newFileName = "Bulk_Questions_Template" + DateTime.Now.Ticks.ToString() + ".xlsx";
            string newFilePath = Path.Combine(path, newFileName);
            File.Copy(filePath, newFilePath);
            return newFileName;
        }

        static void EditExel(string fileName)
        {
            int index = 2;
            List<string> questionComplexities = _context.QuestionComplexities.Select(n => n.Name).Distinct().ToList();
            string ddl = "";
            foreach (var item in questionComplexities)
            {
                ddl += item+",";
            }
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "files", fileName);
            //Create an instance for word app  
            Application application = new Application();

            //Set status for word application is to be visible or not.  
            application.Visible = false;
            application.Interactive = false;

            //Create a missing variable for missing value  
            object missing = System.Reflection.Missing.Value;
            object ReadOnlyRecommended = false;

            Workbook Book = application.Workbooks.Open(path, missing, ReadOnlyRecommended);

            Worksheet Sheet = (Worksheet)Book.Worksheets[1];

            bool failed = false;
            do
            {
                try
                {
                    for (int i = index; i < 100; i++)
                    {
                        Range Range = Sheet.get_Range("K" + i);

                        Range.Validation.Add(XlDVType.xlValidateList
                            , XlDVAlertStyle.xlValidAlertStop
                            , XlFormatConditionOperator.xlBetween
                            , ddl
                            , Type.Missing
                            );
                        Range.Validation.InCellDropdown = true;

                        Book.Save();
                        // Sheet.EnableSelection = XlEnableSelection.xlNoSelection;
                    }
                    failed = false;
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    failed = true;
                }
                System.Threading.Thread.Sleep(10);
            } while (failed);

            var x = Book.Saved;

            //CLEAN UP
            GC.Collect();
            GC.WaitForPendingFinalizers();


            if (Book != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Book);
            application.Quit();
            if (application != null)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(application);



        }

    }
}
