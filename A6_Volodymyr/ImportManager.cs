using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace A6_Volodymyr
{
    public class ImportManager
    {
        string test;
        public string[] arrayTest;

        public string strTest
        {
            get { return test; }
        }

        public bool Import(bool excel, int filterIndex, string fileName, TaskManager taskManager)
        {
            bool success = false;

            if (excel)
            {
                switch (filterIndex)
                {

                    case 1:
                        success = ImportTxtFile(fileName, taskManager);

                        break;
                        
                    case 2:
                        success = ImportExcelFile(fileName, taskManager);

                        break;

                    case 3:
                        success = ImportCalcFile(fileName, taskManager);

                        break;
                }
            }
            else
            {
                switch (filterIndex)
                {
                    case 1:

                        success = ImportTxtFile(fileName, taskManager);

                        break;

                    case 2:
                        success = ImportCalcFile(fileName, taskManager);

                        break;


                }

            }

            return success;
        }

        private bool ImportTxtFile(string fileName, TaskManager taskManager)
        {

            bool success = false;
            //This array will be used to store contents of the txt file
            string[] strFileContent;
            //reading every single line in the .txt and storing everything as an array
            strFileContent = System.IO.File.ReadAllLines(fileName);
            //this loop will go through every line and convert it into a task, which will then be added to the task manager
            foreach(string item in strFileContent)
            {
                //creating an instance of a task
                Task task = new Task();

                string strDate;
                string strTime;
                string strDateTime;
                string strPriority;
                string strDescription;

                int intEnd = item.Length; //gets the length of a current line in the .txt document

                //since we expect to have the length of at least 65 (based on how my listbox and export features work) 
                //if a .txt has anything less than that, the program will, most probably, not be able to read it

                if (intEnd < 66
                    ) 
                {
                    success = false;
                    return success;
                }

                //the date will always have the same length and format(assuming the user doesn't edit the file of course)
                //therefore, we can use constant values

                //Substring allows to read a specified portion of the string
                //It is neccesarry for us in order to be able to store the data in the corresponding cells, such as DateTime, Priority and description
                //Otherwise, the program will import the entire line as one array item
                //The program will work just fine in this case, but I thought it would be more interesting to write a code that will
                //pick apart the info in the .txt and sends it to corresponding properties of the Task.cs
                //It allows for more flexibility
                //Who knows, maybe some other class might want to only receive one type of data (for example only priority) from the text file
                //It would allow for more flexibility

                //Therefore, I wrote this code the reads all the info in the .txt and sends it to corresponding properties of the Task.cs
                //Then, task manager will add this task to the list

                //It is worth mentioning that this code will work ONLY if the file is in exactly the same state as it was after the export
                //In other words, if the user edits a file after the export from this program, the program might not be able to read it
                //Thus, it is important to take care of this by checking the validity of the file

                //String.Substring(starting point, length (how long until the program should stop reading))


                strDate = item.Substring(0, 11);
                strTime = item.Substring(20, 5);
                strDateTime = strDate + strTime;

                //This will read all the data starting from the line where priority is expected to be
                //Considering there are four differnt options, with each and avery one having a different string length, it is impossible to simply set an int
                //represening the length that we want to read
                //This means that this string will also include anything after the priority, including the description
                //it is neccesarry to edit it, so that there are no errors
                //Also, the problem with this is that everything after the priority type will be included
                //It means the description and the empty space before that will be included in the string
                //Of course, the program will not compile, because the enum will be rendered unreadable
                //Thus, we need to cut out everything that does not represent priority type

                strPriority = item.Substring(43);

                //reads description
                strDescription = item.Substring(65, intEnd - 65);

                //gets the first instance of an empty space (" ") in the description string
                int priorityEnd = strPriority.IndexOf(" ");
                //one of the ways to make sure we are doing everything right
                //if there are no empty spaces the user might have done something wrog by editing the .txt file
                if (priorityEnd == -1)
                {
                    success = false;
                    return success;
                }
                //The priority types are saved in the file with spaces (eg. Less Important)
                //The enums, on the other hand don't have those (eg. Less_Important)
                //Therefore we have to convert the spaces in between to underscores
                //However, simply doing so will turn make the one-word enums unreadble
                //For example, the program will turn "Important" into "Important_"
                //Thereofore, we want to check if there is a character after the empty space
                //if there is one, the program will replace it with an underscore "_"
                //if there is an empty space, this part will be ommitted
                //This is also neccesarry because in the next step the program will start cutting out everything that is not the priority type
                //It will do it after the first empty space
                //Without the underscore, errors are possible
                //For example, it will turn "Less Important" into "Less"
                //Obviously the program will not compile
                if (strPriority.Substring(priorityEnd + 1, 1) != " ")
                {
                    strPriority = strPriority.Remove(priorityEnd, 1);
                    strPriority = strPriority.Insert(priorityEnd, "_");
                    priorityEnd = strPriority.IndexOf(" ");
                }

                //gets the entire length of the strPriority
                //it is neccesatty because I didn't find a better way to get the priori
                int priorityLength = strPriority.Length;
                strPriority = strPriority.Remove(priorityEnd, priorityLength - priorityEnd);
                test = strPriority;

                //I've read up on try-catch-finally
                //It seems like a good way of dealing with unhandled exceptions
                //The user might edit the file so that it is unreadable
                //It would be nice if instead of crashing the program displayed an error message and returned to normal
                //Therefore, I'm going to try and catch all the errors that might be reasonable to excpect
                //However, I am just a beginner and it is entire possible that I missed some exceptions
                //
                try
                {
                    task.Date = DateTime.Parse(strDateTime);
                }
                catch(System.FormatException) //this will happen if the user edited the datetime to an unreadble format
                {
                    success = false;
                    return success;
                }

                try
                {
                    task.Priority = (PriorityType)Enum.Parse(typeof(PriorityType), strPriority);
                }
                //if the user has edited an priority level, so that it no longer part of the PriorityType.cs, the program will return an error message
                catch (System.ArgumentException) 
                {
                    success = false;
                    return success;
                }

                task.Description = strDescription;

                taskManager.Add(task);

                success = true;

            }

            return success;

        }

        private bool ImportExcelFile(string fileName, TaskManager taskManager)
        {
            bool success = true;

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            var workbooks = excelApp.Workbooks;

            var workbook = workbooks.Open(fileName);

            var sheet = workbook.Sheets[1];

            Microsoft.Office.Interop.Excel.Range range;

            range = sheet.Range["A1", "A1"];

            int usedColumns = sheet.UsedRange.Columns.Count;
            int usedRows = sheet.UsedRange.Rows.Count;
            string cellName;
            int counter = 1;


            for (int index = 1; index <= usedRows; index++)
            {
                int priorityEnd = -1;
                string strDate;
                double dblDate;
                string strTime;
                string strDateTime;
                string strPriority;
                string strDescription;

                Task task = new Task();

                cellName = "A" + index.ToString();
                range = sheet.Range(cellName, cellName);
                try
                {
                    strDate = range.Value2;
                }
                //this was one of the exceptions I encountered when exporting a file that was not created in this programm
                //and was filled with random characters or when trying to import an epmty file
                //I tried doing "if(!String.IsNullOrEmpty(range.Value2.ToString()))", but, since range2 is dynamic
                //it failed to convert the empy cell to string, since it didn't have any info to go off
                //thus, I had to resort to try-catch method
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    success = false;
                    //Most of the comments for this are done in ExportManager, line 132

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);
                    return success;
                }


                cellName = "B" + index.ToString();
                range = sheet.Range(cellName, cellName);
                //due to the way date is written excel percieves it as a bool
                //thus, it was neccesarry to record it this way first and convert it to string later
                try
                {
                    dblDate = range.Value2;
                    DateTime conv = DateTime.FromOADate(dblDate);
                    strTime = conv.ToString();
                }
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    success = false;
                    //Most of the comments for this are done in ExportManager, line 132

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);
                    return success;
                }

                strDateTime = strDate + strTime;
                strDateTime = strDateTime.Replace("30.12.1899", "");

                cellName = "C" + index.ToString();
                range = sheet.Range(cellName, cellName);
                strPriority = range.Value2;
                if(String.IsNullOrWhiteSpace(strPriority))
                {
                    success = false;

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);

                    return success;
                }


                cellName = "D" + index.ToString();
                range = sheet.Range(cellName, cellName);
                strDescription = range.Value2;

                if (String.IsNullOrWhiteSpace(strDescription))
                {
                    success = false;

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);

                    return success;
                }

                priorityEnd = strPriority.IndexOf(" ");


                if (priorityEnd != -1)
                {
                    strPriority = strPriority.Remove(priorityEnd, 1);
                    strPriority = strPriority.Insert(priorityEnd, "_");
                    priorityEnd = strPriority.IndexOf("");
                }



                try
                {
                    task.Date = DateTime.Parse(strDateTime);
                }
                catch (System.FormatException) //this will happen if the user edited the datetime to an unreadble format
                {
                    success = false;

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);

                    return success;
                }

                try
                {
                    task.Priority = (PriorityType)Enum.Parse(typeof(PriorityType), strPriority);
                }
                //if the user has edited an priority level, so that it no longer part of the PriorityType.cs, the program will return an error message
                catch (System.ArgumentException)
                {
                    success = false;


                    test = strPriority;

                    workbook.Close(fileName);

                    excelApp.Quit();

                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(excelApp);

                    return success;
                }

                task.Description = strDescription;

                taskManager.Add(task);

                success = true;

                test = strPriority;

                ++counter;
            }



            workbook.Close(fileName);

            excelApp.Quit();

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excelApp);

            success = true;

            return success;

        }

        private bool ImportCalcFile(string fileName, TaskManager taskManager)
        {
            bool success = false;
            //before i created an istance of Task.cs inside the loop which scans all the file contents
            //This was neccessarry because I was expecting to extract multiple tasks from .txt and/or .xlsx
            //Therefore, with each new iteration of the loop a new task would have to be created and added
            //
            //When it comes to .ics it really is only possible to store/extract one task
            //Therefore, it is not neccessary to create and destroy instances with every new iteration of the loop
            Task task = new Task();

            string strDateTime;
            string strPriority;
            string strDescription;

            string[] strFileContent;

            //reading every single line in the .txt and storing everything as an array
            strFileContent = System.IO.File.ReadAllLines(fileName);

            arrayTest = strFileContent;

            int indexDateTime = -1;
            int indexPriority = -1;
            int indexDescription = -1;

            //.ics has a very specific format
            //Thus, in order to extract the neccesarry information we will scan the file for the lines that our programm can read, and ignore the rest
            //By doing this I assume that the .ics will have the same logic as I programmed in ExportManager.cs
            //That is: Date and time stored at "DTSTART:"
            //Priority stored at "PRIORITY:"
            //I am not going to read priority off of "SUMMARRY:"
            //Reading it from "PRIORITY:" allows me to import any .ics, not just the one's created by my program
            //Description stored at "DESCRIPTION:"
            for (int i = 0; i < strFileContent.Length; i++)
            {
                //The designated start date is the date for when the task will be due
                if(strFileContent[i].Contains("DTSTART:"))
                {
                    indexDateTime = i;
                }

                if (strFileContent[i].Contains("PRIORITY:"))
                {
                    indexPriority = i;
                }

                if (strFileContent[i].Contains("DESCRIPTION:"))
                {
                    indexDescription = i;
                }
            }

            strDateTime = strFileContent[indexDateTime];
            //.ics stores dates in the following format: DTSTART:YYYYMMDDTHHMMSS
            //We have to get rid of "DTSTART:"
            //Otherwise, the program will not read it as DateTime
            //We also need to add "-" between dates and ":" between times
            strDateTime = strDateTime.Remove(0, 8);
            strDateTime = strDateTime.Insert(4, "-");
            strDateTime = strDateTime.Insert(7, "-");
            strDateTime = strDateTime.Insert(13, ":");
            strDateTime = strDateTime.Insert(16, ":");

            strPriority = strFileContent[indexPriority];
            //We have to get rid of "PRIORITY:"
            strPriority = strPriority.Remove(0, 9);

            strDescription = strFileContent[indexDescription];
            //We have to get rid of "DESCRIPTION:"
            strDescription = strDescription.Remove(0, 12);

            int intPriority = Int32.Parse(strPriority);

            try
            {
                task.Date = DateTime.Parse(strDateTime);
            }
            catch (System.FormatException) //this will happen if the user edited the datetime to an unreadble format
            {
                success = false;
                return success;
            }
            //based on the priority set in the file, the program will pick the corresponging options from PriorityType.cs
            if(intPriority == 6 || intPriority == 7 || intPriority == 8 || intPriority == 9 || intPriority == 0)
            {
                task.Priority = PriorityType.Less_Important;
            }
            else if(intPriority == 5)
            {
                task.Priority = PriorityType.Normal;
            }
            else if(intPriority == 4 || intPriority == 3)
            {
                task.Priority = PriorityType.Important;
            }
            else if (intPriority == 2 || intPriority == 1)
            {
                task.Priority = PriorityType.Very_Important;
            }
            else if(intPriority == -1)
            {
                success = false;
                return success;
            }

            if(String.IsNullOrWhiteSpace(strDescription))
            {
                success = false;
                return success;
            }
            else
            {
                task.Description = strDescription;
            }

            taskManager.Add(task);

            success = true;

            return success;
        }
    }
}
