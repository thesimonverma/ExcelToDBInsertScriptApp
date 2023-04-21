// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

Console.WriteLine("Program started!");

try
{
    //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
    using (SpreadsheetDocument doc = SpreadsheetDocument.Open("C:\\Users\\xxxxx\\Downloads\\b.xlsx", false))
    {
        //create the object for workbook part  
        WorkbookPart workbookPart = doc.WorkbookPart;
        Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
        StringBuilder excelResult = new StringBuilder();
        List<string> list = new List<string>();

        //using for each loop to get the sheet from the sheetcollection  
        foreach (Sheet thesheet in thesheetcollection)
        {
            //excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
            //excelResult.AppendLine("----------------------------------------------- ");
            //statement to get the worksheet object by using the sheet id  
            Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

            SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
            foreach (Row thecurrentrow in thesheetdata)
            {
                list = new List<string>();
                foreach (Cell thecurrentcell in thecurrentrow)
                {
                    //statement to take the integer value  
                    string currentcellvalue = string.Empty;
                    if (thecurrentcell.DataType != null)
                    {
                        if (thecurrentcell.DataType == CellValues.SharedString)
                        {
                            int id;
                            if (Int32.TryParse(thecurrentcell.InnerText, out id))
                            {
                                SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                if (item.Text != null)
                                {
                                    //code to take the string value  
                                    list.Add(item.Text.Text);
                                }
                                else if (item.InnerText != null)
                                {
                                    currentcellvalue = item.InnerText;
                                }
                                else if (item.InnerXml != null)
                                {
                                    currentcellvalue = item.InnerXml;
                                }
                            }
                        }
                    }
                    else
                    {
                        excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                    }
                }
                //append here
                list[1] = list[1].Replace("'", "''");
                excelResult.AppendLine(String.Format("\tIF NOT EXISTS (SELECT * FROM [dbo].[xxx]\n\t\t\t\tWHERE a like '{0}' \n\t\t\t\tAND b like '{1}')\n\tBEGIN\n\t\tINSERT INTO [dbo].[doctype] \n\t\tVALUES ('{0}','{1}') \n\tEND", list[0], list[1]));
                excelResult.AppendLine("\t-------------------------------------------------------------------------");
            }
            excelResult.Append("");
            Console.WriteLine(excelResult.ToString());
            using (StreamWriter writetext = new StreamWriter("C:\\Users\\xxxx\\Downloads\\sqlquery.txt"))
            {
                writetext.WriteLine("BEGIN");
                writetext.WriteLine(excelResult.ToString());
                writetext.WriteLine("END");
            }
            Console.ReadLine();
        }
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}
