using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace A6_Volodymyr
{

    /// <summary>
    /// At first all of this code was implemented into the main form.
    /// 
    /// However, the code became to messy, and I thought it would be a good idea to let a separate class handle the export process
    /// 
    /// The job of this class is to create a file with set parameters recieved from the Save File Dialogue created by the Main Form
    /// </summary>

    public class ExportManager
    {

        public ExportManager()
        {

        }

        //this class will trigger different methods for export, depending on which targer file format has been chosen
        public void Export(bool excel, int filterIndex, string fileName, TaskManager taskManager, int selectedIndex)
        {

            if (excel)
            {
                switch (filterIndex)
                {

                    case 1:
                        CreateTxtFile(fileName, taskManager);

                        break;

                    case 2:
                        CreateExcelFile(fileName, taskManager);

                        break;

                    case 3:
                        CreateCalendarFile(fileName, taskManager, selectedIndex);

                        break;
                }
            }
            else //if user did not want excel support, it will be ommitted
            {
                switch (filterIndex)
                {
                    case 1:

                        CreateTxtFile(fileName, taskManager);

                        break;

                    case 2:
                        CreateCalendarFile(fileName, taskManager, selectedIndex);

                        break;


                }

            }

        }
        
        //The simplest of them all, just copies everything from the task manager and pastes it into a .txt using the name provided by the user
        private void CreateTxtFile(string fileName, TaskManager taskManager)
        {
            string[] strList = taskManager.ListToStringArray();

            System.IO.File.WriteAllLines(fileName, strList);
        }
        //this one was the most difficult to implement
        //Excel didn't want to close properly and caused memory leaks
        //This sholud be fixed now
        //See line 133 for more info
        private void CreateExcelFile(string fileName, TaskManager taskManager)
        {
            //initiate the Excel application and create an instance of it
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            //make it invisible
            excelApp.Visible = false;
            //this represents all the workbooks that are currently open
            var workbooks = excelApp.Workbooks;
            //by adding a new workbook we open one and can start editing it
            var workbook = workbooks.Add();
            //we are opening a new sheet in the workbook
            var sheet = workbook.Sheets[1];
            //this represents a cell or a range of cells, usefull to specify a specific point
            Microsoft.Office.Interop.Excel.Range range;
            //we set the range to a single cell A1 by default
            range = sheet.Range["A1", "A1"];

            //this string will help us move around between the columns of the spreadsheet
            //coupled with the counter that will move to the next row once a task has been recorded in spreadsheet
            string cellName;
            int counter = 1;
            foreach (var item in taskManager.list)
            {
                //records the date in the cell A. The row, defined by int counter depends on which task is currently being recorded
                cellName = "A" + counter.ToString();
                range = sheet.Range(cellName, cellName);
                range.Value2 = item.Date.ToShortDateString();
                //recodes the time in cell B
                cellName = "B" + counter.ToString();
                range = sheet.Range(cellName, cellName);
                range.Value2 = item.GetTimeToString();
                //records the priority in cell C
                cellName = "C" + counter.ToString();
                range = sheet.Range(cellName, cellName);
                range.Value2 = item.GetPriorityToString();
                //records the description in cell D
                cellName = "D" + counter.ToString();
                range = sheet.Range(cellName, cellName);
                range.Value2 = item.Description;

                ++counter;

            }
            //saving the workbook under the name the user has provided
            workbook.SaveAs(fileName);
            //closing the said workbook
            workbook.Close(fileName);
            //closing excel
            excelApp.Quit();
            //I noticed there were spikes in memory usage whenever I exported something into excel
            //I opened the task manager and found that there were about ten EXCEL.EXE processes running
            //I googled the problem and realised it was a memory leak
            //This prompted me to google how to solve it
            //The problem was that once I was done exporting, the program still held references to some excel objects
            //There were three main rules I found that a lot of people recommended
            /*
             * 1) No double dots (Explained earlier)
             * 2) I should release all the ComObjects that I use in reverse order of initiating them
             * 3) I can try and call a garbage collector manually


             */

            /*
             * I tested different configurations and the program seemes to work if I put a garabage collector after I call the export manager in MainForm
             * I kept the code for releasing the values here just in case I doesn't work properly
             * 
             * If the program has crashed when you were doing something with excel, I suggest opening the task managaer going to the "Details" window and checking
             * that ther is no EXCEL.EXE running. Otherwise, the garbage collection should do its job

             */
            //release to avoid memory leaks

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excelApp);
        }
        //calendar files (.ics) follow a very strict structure which is avaliable online
        //any .txt file can be turned into .ics if the structure has been followed
        private void CreateCalendarFile(string fileName, TaskManager taskManager, int selectedIndex)
        {
            //It was difficult for me to figure out how to create multiple .ics files
            //I think it makes more sence to have one event at a time
            //Therefore, the export will only happen if the user has selected a listbox item first
            if (selectedIndex != -1)
            {
                //I think one can say that StringBuilder class is something in between string and string[]
                //Basically it allows to fill itself with multiple string characters
                //It is kind of like a .txt file, except it is not a file and all the data is saved in the process memory
                StringBuilder stringBuilder = new StringBuilder();
                //gets the date from a selevted task
                DateTime dateStart = taskManager.list[selectedIndex].Date;
                //adds 60 minutes to the start date, so it can be used as the end date
                DateTime dateEnd = dateStart.AddMinutes(60);
                //will be used to assess priority
                int priority = 0;
                //.ics has three priority levels: LOW (6-9) MEDIUM (5) HIGH (1-4)
                switch(taskManager.list[selectedIndex].Priority)
                {
                    case PriorityType.Less_Important:
                        priority = 6;

                        break;

                    case PriorityType.Normal:
                        priority = 5;

                        break;

                    case PriorityType.Important:
                        priority = 3;

                        break;

                    case PriorityType.Very_Important:
                        priority = 1;

                        break;
                }
                string summary = taskManager.list[selectedIndex].GetPriorityToString();
                string description = taskManager.list[selectedIndex].Description;
                //start building our .ics file
                //this is all taken from the typical structure of .ics File
                stringBuilder.AppendLine("BEGIN:VCALENDAR");
                stringBuilder.AppendLine("VERSION:2.0");
                stringBuilder.AppendLine("PRODID:A6_Volodymyr");
                stringBuilder.AppendLine("CALSCALE:GREGORIAN");
                stringBuilder.AppendLine("METHOD:PUBLISH");
                stringBuilder.AppendLine("BEGIN:VTIMEZONE");
                stringBuilder.AppendLine("TZID:Europe/Stockholm");
                stringBuilder.AppendLine("BEGIN:STANDARD");
                stringBuilder.AppendLine("TZOFFSETTO:+0100");
                stringBuilder.AppendLine("TZOFFSETFROM:+0100");
                stringBuilder.AppendLine("END:STANDARD");
                stringBuilder.AppendLine("END:VTIMEZONE");
                stringBuilder.AppendLine("BEGIN:VEVENT");
                //start adding our own parameters
                stringBuilder.AppendLine("DTSTART:" + dateStart.ToString("yyyyMMddTHHmm00"));
                stringBuilder.AppendLine("DTEND:" + dateEnd.ToString("yyyyMMddTHHmm00"));
                //priority is recorded two times: as summary and as an actual priority
                //I felt like priority type was a good brief summary, plus the user can easily edit it by opening the file outside the program
                stringBuilder.AppendLine("SUMMARY:" + summary + "");
                stringBuilder.AppendLine("LOCATION:" + "");
                stringBuilder.AppendLine("DESCRIPTION:" + description + "");
                stringBuilder.AppendLine("PRIORITY:" + priority.ToString());
                stringBuilder.AppendLine("END:VEVENT");

                //end calendar item
                stringBuilder.AppendLine("END:VCALENDAR");
                //save everything into a single string;
                string calendarItem = stringBuilder.ToString();
                //export
                System.IO.File.WriteAllText(fileName, calendarItem);
            }
            else
            {
                //if the user did not select anything, the program will just go back to the MainForm
                return;
            }
        }
    
    
    }
}
