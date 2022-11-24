using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A6_Volodymyr
{
    public class TaskManager
    {
        public List<Task> taskList; //creating a List<T> with tasks
        //a constructor where we initalize the List<T>
        public TaskManager()
        {
            taskList = new List<Task>();
        }
        //gets the amount of tasks
        public int Count
        {
            get { return taskList.Count; }
        }
        //setting the list to be the property of TaskManager
        public List<Task> list
        {
            get { return taskList; }
        }
        //class responsible for adding a task, and sorting the entire array by date
        public void Add(Task task)
        {
            int count = Count;

            if(count != -1)
            {
                taskList.Add(task);
                taskList.Sort((x, y) => x.Date.CompareTo(y.Date));
            }
        }
        //edting the ToDo list
        public void Edit(int index, Task task)
        {
            taskList.RemoveAt(index);
            taskList.Insert(index, task);
            taskList.Sort((x, y) => x.Date.CompareTo(y.Date));
        }
        //deleting an entry
        public void Delete(int index)
        {
            if (index != -1)
            {
                taskList.RemoveAt(index);
            }
        }
        //converting the array to string, so that it can be displayed in the listbox
        public string[] ListToStringArray() 
        { 
            string[] taskArray = new string[Count]; 
            for (int i = 0; i < Count; i++) taskArray[i] = taskList[i].ToString(); 
            return taskArray; 
        }
    }
}
