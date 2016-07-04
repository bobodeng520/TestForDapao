using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace ConsoleApplication1
{
    class Program
    {
        public class Account
        {
            public int ID { get; set; }
            public double Balance { get; set; }
        }

        static void DisplayInExcel(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
            var wordApp = new Word.Application();

            // Open the specific excel
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"D:\TestForDapao\ConsoleApplication1\攀枝花川港-检定（新）.xlsx");
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);

            // Open the specific word
            Word.Document wordDocument = (Word.Document)wordApp.Documents.Open(@"D:\TestForDapao\ConsoleApplication1\攀枝花川港-LNG加气机-校准.doc");

            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned  
            // by property Workbooks. The new workbook becomes the active workbook. 
            // Add has an optional parameter for specifying a praticular template.  
            // Because no argument is sent in this example, Add creates a new workbook. 
            //excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is 
            // removed in a later procedure.
            //Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            //workSheet.Cells[1, "A"] = "ID Number";
            //workSheet.Cells[1, "B"] = "Current Balance";

            //var row = 1;
            //foreach (var acct in accounts)
            //{
                //row++;
                //workSheet.Cells[row, "A"] = acct.ID;
                //workSheet.Cells[row, "B"] = acct.Balance;
            //}

            //workSheet.Columns[1].AutoFit();
            //workSheet.Columns[2].AutoFit();

            //((Excel.Range)workSheet.Columns[1]).AutoFit();
            //((Excel.Range)workSheet.Columns[2]).AutoFit();

            // Put the spreadsheet contents on the clipboard. The Copy method has one 
            // optional parameter for specifying a destination. Because no argument   
            // is sent, the destination is the Clipboard.
            excelWorksheet.Range["A11:K21"].Copy();

            
            wordApp.Visible = true;

            // The Add method has four reference parameters, all of which are  
            // optional. Visual C# 2010 allows you to omit arguments for them if 
            // the default values are what you want.
            //wordApp.Documents.Add(@"D:\TestForDapao\ConsoleApplication1\攀枝花川港-LNG加气机-校准.doc");

            // PasteSpecial has seven reference parameters, all of which are  
            // optional. This example uses named arguments to specify values  
            // for two of the parameters. Although these are reference  
            // parameters, you do not need to use the ref keyword, or to create  
            // variables to send in as arguments. You can send the values directly.
            object restr1 = "基本误差";
            if (wordDocument.Bookmarks.Equals("基本误差"))
            //if (wordApp.Selection.Find.Equals("基本误差"))
            {
                //object BookMarkName = "基本误差";
                //object what = Word.WdGoToItem.wdGoToBookmark;
                //object Nothing = System.Reflection.Missing.Value;
                //wordDocument.ActiveWindow.Selection.GoTo(ref what, ref Nothing, ref Nothing, ref BookMarkName);
                //object dummy = System.Reflection.Missing.Value;
                //object what = Word.WdGoToItem.wdGoToLine;
                //object which = Word.WdGoToDirection.wdGoToNext;
               // object count = 1;
                //wordApp.Selection.GoTo(ref what, ref which, ref count, ref dummy);

                //object unit = Word.WdUnits.wdParagraph;
                //object count = 1;
                //object extend = Word.WdMovementType.wdExtend;
                //wordApp.Selection.MoveDown(ref unit, ref count, ref extend);
                //wordApp.Selection.TypeParagraph();
                //Console.WriteLine(123);
                System.Console.WriteLine("Hello World!");
                wordApp.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
            }
            

        }

        static void CreateIconInWordDoc()
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;

            // The Add method has four reference parameters, all of which are  
            // optional. Visual C# 2010 allows you to omit arguments for them if 
            // the default values are what you want.
            wordApp.Documents.Add();

            // PasteSpecial has seven reference parameters, all of which are  
            // optional. This example uses named arguments to specify values  
            // for two of the parameters. Although these are reference  
            // parameters, you do not need to use the ref keyword, or to create  
            // variables to send in as arguments. You can send the values directly.
            //wordApp.Selection.PasteAndFormat();
        }

        static void Main(string[] args)
        {
            // Create a list of accounts. 
            var bankAccounts = new List<Account> {
                new Account { 
                  ID = 345678,
                  Balance = 541.27
                },
                new Account {
                  ID = 1230221,
                  Balance = -127.44
                }
            };

            // Display the list in an Excel spreadsheet.
            DisplayInExcel(bankAccounts);

            // Create a Word document that contains an icon that links to 
            // the spreadsheet.
            //CreateIconInWordDoc();
        }
    }
}
