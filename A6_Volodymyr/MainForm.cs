using System;
using System.Text;
using System.Windows.Forms;

namespace A6_Volodymyr
{
    public partial class MainForm : Form
    {
        TaskManager taskManager = new TaskManager();
        bool excelSupport;

        public MainForm()
        {
            InitializeComponent();
            InitializeGUI();
        }

        private void InitializeGUI() //initialising user interface
        {

            /*
             I thought it would be very nice if the user could export the ToDo list to Microsoft Excel
            However, from my exprience and excperience of other people online, it seems like using C# to create Microsoft Excel spreadsheets could cause problems
            and/or make the program ustable

            Therefore I introduced this little part, which asks the user if the Excel the user wants the Excel support
            If not, the Excel part will be entirely ommited from the program
            I tested it on multiple computers, and it worked without problems

            If, for some reason the program still crashes, do the following: 

            1) comment out the code on lines 184-195 (the whole of (if(excelsupport)))
            2) then comment out the code on lines  331-417 (the whole of private void exportToText_ItemToolStripMenuItem_Click(object sender, EventArgs e))
             */

            ExcelSupport();


            //Initializes the menu strip
            InititalizeMenuStrip();

            //when a combox set to DropDownList, the user can't edit it and thus can't add new values, which in turn 
            //might result in unhandled exceptions and errors that could crash the program
            cmbPriority.DropDownStyle = ComboBoxStyle.DropDownList; 

            cmbPriority.Items.Clear(); //clearing the combobox

            //converting the items contained in the PriorityType enum into a string array
            //this is neccessarry so that we remove underscores ("_") from the priority types
            //It makes the user interface look nicer
            string[] priority = Enum.GetNames(typeof(PriorityType));

            //Going through every item in the array and adding it to the combobox, 
            //while replacing the underscores with empty spaces ("_", " ")
            foreach (var item in priority) cmbPriority.Items.Add(item.Replace("_", " "));

            cmbPriority.SelectedItem = "Normal"; //setting the default comboBox choice to "Normal"

            //clearing out all the input and output items
            txtToDo.Text = String.Empty;
            lstResults.Items.Clear();

            //writign a tooltip for every input and output item on the form
            toolTip.SetToolTip(dateTimePicker, "Click to open calendar for the date and type in the time here");
            toolTip.SetToolTip(cmbPriority, "Click to choose the priority level for your task");
            toolTip.SetToolTip(txtToDo, "Type a brief description of your task");
            toolTip.SetToolTip(lstResults, "This is where all your tasks are displayed. You can select a task by clicking on it");
            toolTip.SetToolTip(btnAdd, "Click to add a new entry to the list");
            toolTip.SetToolTip(btnEdit, "Click to edit selected task");
            toolTip.SetToolTip(btnDelete, "Click to delete selected task");

            //setting dateTimePicker to a custom format
            dateTimePicker.Format = DateTimePickerFormat.Custom;
            dateTimePicker.CustomFormat = "yyyy-MM-dd HH:mm";

            btnDelete.Enabled = false;
            btnEdit.Enabled = false;

            this.Icon = new System.Drawing.Icon("icon.ico");
        }



        private void InititalizeMenuStrip() //this is a separate intialisation class for the menuStrip
        {
            /// <summary>
            /// For this assignment I chose to try and add menuStrip items dynamically, rather than pre-define them in
            /// the MainForm.cs [Design]. This would also mean I would have to programm my own event handlers
            /// I know they would normally be created automatically in the MainForm.Designer.cs once I double click on an existing item
            /// But since I chose to add the Items dynamically, I had to manually program them in this class.
            /// I hope this is ok, since it would allow me to better understand the EventHandlers
            /// Plus, it was neccessary to do in order to allow the user to choose between Excel support or not
            /// In case user doesn't want Excel support, the corresponding sub-item will not be created
            /// If not, I am ready to re-submit the task with automatically created EventHandlers
            /// </summary>
            ///
            menuStrip.Items.Clear(); //clears all the items from the menu strip
            //useful when we will need to call InitializeGUI() (and, consequently InititalizeMenuStrip() method) again

            //declaring a new variable of the ToolStripMenuItem() type
            //This will be a menuStrip item called "File"
            //I chose this path because I couldn't figure out how to add sub items using menuStrip.Add.Item()
            var fileItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "File",
                Text = "File"
            };

            //this will later become a sub-Item "New"
            var newItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "New",
                Text = "New"
            };

            //assignining a Ctrl+N shortcut to the sub-item "New"
            newItem.ShortcutKeys = Keys.Control | Keys.N;

            //this will later become a sub-Item "Export"
            var exportItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "Export",
                Text = "Export"
            };

            exportItem.ShortcutKeys = Keys.Control | Keys.S;

            //sub-item for import
            var importItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "Import",
                Text = "Import"
            };

            importItem.ShortcutKeys = Keys.Control | Keys.O;

            //this will later become a sub-Item "Exit"
            var exitItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "Exit",
                Text = "Exit"
            };

            //assignining a Alt+F4 shortcut to the sub-item "Exit"
            exitItem.ShortcutKeys = Keys.Alt | Keys.F4;

            //this will later become an item "Help"
            var helpItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "Help",
                Text = "Help"
            };
            //this will later become a sub-Item "About"
            var aboutItem = new System.Windows.Forms.ToolStripMenuItem()
            {
                Name = "About",
                Text = "About"
            };

            //initializing separators for the menuStrip
            var separator1 = new System.Windows.Forms.ToolStripSeparator();
            var separator2 = new System.Windows.Forms.ToolStripSeparator();


            //adding sub-items to the "File" item
            fileItem.DropDownItems.Add(newItem);
            fileItem.DropDownItems.Add(separator1);
            fileItem.DropDownItems.Add(exportItem);
            fileItem.DropDownItems.Add(importItem);
            fileItem.DropDownItems.Add(separator2);
            fileItem.DropDownItems.Add(exitItem);
            //by adding a "File" item we will also add all the sub-items
            menuStrip.Items.Add(fileItem);

            helpItem.DropDownItems.Add(aboutItem);
            menuStrip.Items.Add(helpItem);

            //creating a new EventHandler for when the user clicks on the "File" sub-items
            newItem.Click += new EventHandler(newItemToolStripMenuItem_Click);
            exitItem.Click += new EventHandler(exitItemToolStripMenuItem_Click);
            exportItem.Click += new EventHandler(export_ItemToolStripMenuItem_Click);
            importItem.Click += new EventHandler(import_ItemToolStripMenuItem_Click);
            aboutItem.Click += new EventHandler(about_ItemToolStripMenuItem_Click);
        }

        private bool ExcelSupport()
        {
            string msgBoxtext = "Would you like to enable Microsoft Excel support\n\n" +
                  "This feature is still in development and might cause unexpected errors (but only if you actually use the feature)\n\n" +
                  "Enabling it will allow you to export your To-Do lists to excel (.xlsx)\n\n" +
                  "It is strongly recommended to choose no if you don't have Microsoft Excel installed\n\n" +
                  "IMPORTANT: ONLY IMPORT EXCEL FILES THAT YOU CREATED USING THIS PROGRAM!\n\n" +
                  "EXPORTING FILES THAT WERE CREATED SOMWHERE ELSE CAN CAUSE CRASHES AND UNHANDLED EXCEPTIONS!\n\n" +
                  "This does not apply to import/export of .txt and .ics files";

            string title = "Microsoft Excel Support";

            MessageBoxButtons buttons = MessageBoxButtons.YesNo;

            //A message box will pop-up when the program is started.
            //If the user chooses yes, it would be possible to export the ToDo list to Excel
            //If the user chooses no, it would be disabled
            DialogResult result = MessageBox.Show(msgBoxtext, title, buttons);
            if (result == DialogResult.Yes)
            {
                excelSupport = true;
            }
            else
            {
                excelSupport = false;
            }

            return excelSupport;
        }

        private void Timer_Seconds_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
        }



        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            string message = "Are you sure you want to exit?";
            string title = "Exit";
            MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
            DialogResult result = MessageBox.Show(message, title, buttons);
            if (result == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
            else
            {
                e.Cancel = false;
            }
        }

        //In the case the user clicks the "New" option InitializeGUI() will be started again
        private void newItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InitializeGUI();
        }

        //In the case the user clicks the "Exit" option
        private void exitItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }



        private void export_ItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //checks if the user has selected an item in the listbox
            int selectedListBoxItem = lstResults.SelectedIndex;

            //I felt like the best approach when exporting to a .ics file would be to allow user to only select one calendar Item at a time
            //Otherwise, for every item in to do list, the user would have to specify the name and file location
            //It makes more sense to only allow to export one calendar item at a time 
            if (selectedListBoxItem == -1)
            {
                MessageBox.Show("You didn't select an item from the lisbox\n\n" +
                    "If you want to export as a calender (.ics) you HAVE to chose an item from the listbox\n\n" +
                    "If you don't do this, the program will run as normal, but your calendar file won't be saved");
            }

            SaveFileDialog saveAs = new SaveFileDialog();//creates an instance of a SaveAs dialogue

            saveAs.Title = "Export"; //Creates a title for the Save File Dialog window

            //the whole point of this is to check if the user wants the excel support or not
            //If excel is supported, an option to export as .xlsx will be added to the combobox with file types in the Save As window
            if (excelSupport)
            {
                saveAs.Filter = "txt File|*.txt|Excel File|*.xlsx|Calendar File|*.ics";
            }
            else
            {
                saveAs.Filter = "txt File|*.txt|Calendar File|*.ics";
            }

            saveAs.ShowDialog(); //Displays the Export window

            if (!String.IsNullOrEmpty(saveAs.FileName))
            {

                ExportManager exportManager = new ExportManager();//creates an instance of an export Manager


                exportManager.Export(excelSupport, saveAs.FilterIndex, saveAs.FileName, taskManager, selectedListBoxItem); //transfering all the neccesarry data to the export manager
                //this was the only way to force EXCEL.EXE to shut down after the export/import has finished
                //There is probably a better way, and from what I've read, calling garbage collector directly doesn't seem to be a good practice
                //However, I have to finish this before the deadline, therefore I leave it like this
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }

            saveAs.Dispose();

        }

        private void import_ItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();

            openFile.Title = "Import"; //Creates a title for the Ipen File Dialog window

            //the whole point of this is to check if the user wants the excel support or not
            //If excel is supported, an option to export as .xlsx will be added to the combobox with file types in the Save As window
            if (excelSupport)
            {
                openFile.Filter = "txt File|*.txt|Excel File|*.xlsx|Calendar File|*.ics";
            }
            else
            {
                openFile.Filter = "txt File|*.txt|Calendar File|*.ics";
            }

            openFile.ShowDialog();

            if(!String.IsNullOrEmpty(openFile.FileName))
            {
                ImportManager importManager = new ImportManager();

                if(importManager.Import(excelSupport, openFile.FilterIndex, openFile.FileName, taskManager) == false)
                {
                    //this was the only way to force EXCEL.EXE to shut down after the export/import has finished
                    //There is probably a better way, and from what I've read, calling garbage collector directly doesn't seem to be a good practice
                    //However, I have to finish this before the deadline, therefore I leave it like this

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    MessageBox.Show("The file you are trying to access does not have the required structure to be read by this program", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
                else
                {

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    UpdateGUI();
                }

            }
        }

        private void about_ItemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 aboutBox = new AboutBox1();
            aboutBox.Show();
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.N)
            {
                InitializeGUI();
            }
        }

        private Task GetTaskFromUserInput()
        {
            string description;
            string strPriority;
            PriorityType priorityType = new PriorityType();

            strPriority = cmbPriority.SelectedItem.ToString();
            strPriority = strPriority.Replace(" ", "_");
            description = txtToDo.Text;

            priorityType = (PriorityType)Enum.Parse(typeof(PriorityType), strPriority);

            Task newTask = new Task();

            newTask.Date = dateTimePicker.Value;
            newTask.Priority = priorityType;
            newTask.Description = description;

            return newTask;

        }

        public void UpdateGUI()
        {
            string[] strToDoList = taskManager.ListToStringArray();

            if(strToDoList != null)
            {
                lstResults.Items.Clear();
                lstResults.Items.AddRange(strToDoList);
            }

            if(lstResults.Items.Count != 0)
            {
                btnDelete.Enabled = true;
                btnEdit.Enabled = true;
            }
            else
            {
                btnDelete.Enabled = false;
                btnEdit.Enabled = false;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (dateTimePicker.Value < new DateTime(2000, 01, 01) || dateTimePicker.Value >= new DateTime(2222, 01, 01))
            {
                MessageBox.Show("Wrong date, must be between years 2000-2222", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if(String.IsNullOrWhiteSpace(txtToDo.Text))
            {
                MessageBox.Show("Please, provide a brief description for your task", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Task newTask = GetTaskFromUserInput();
                if(newTask != null)
                {
                    taskManager.Add(newTask);
                    UpdateGUI();
                }
            }
        }
        //edits an entry
        private void btnEdit_Click(object sender, EventArgs e)
        {
            int index = lstResults.SelectedIndex;

            if (dateTimePicker.Value < new DateTime(2000, 01, 01) || dateTimePicker.Value >= new DateTime(2222, 01, 01))
            {
                MessageBox.Show("Wrong date, must be between years 2000-2222", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (String.IsNullOrWhiteSpace(txtToDo.Text))
            {
                MessageBox.Show("Please, provide a brief description for your task", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Task newTask = GetTaskFromUserInput();
                if (newTask != null)
                {
                    taskManager.Edit(index, newTask);
                    UpdateGUI();
                }
            }
        }
        //deletes an entry
        private void btnDelete_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete this?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if(result == DialogResult.Yes)
            {
                int index = lstResults.SelectedIndex;

                if (index != -1)
                {
                    taskManager.Delete(index);
                    UpdateGUI();
                }
            }
            else if(result == DialogResult.No)
            {
                return;
            }

        }
    }
}
