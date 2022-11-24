using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace A6_Volodymyr
{
    public class Task
    {
        DateTime DT_Date;
        string strDescription;
        PriorityType enmPriority;

        public Task()
        {

        }

        public Task(DateTime DT_Date, PriorityType enmPriority, string strDescription)
        {

        }
        //properties of Task.cs
        //used to store all the values associated with a single task
        public DateTime Date
        {
            get { return DT_Date; }
            set { DT_Date = value; }
        }        


        public PriorityType Priority
        {
            get { return enmPriority; }
            set { enmPriority = value; }
        }

        public string Description
        {
            get { return strDescription; }
            set { strDescription = value; }
        }
        //convert priority to string
        public string GetPriorityToString()
        {
            string strPriority = enmPriority.ToString();
            strPriority = strPriority.Replace("_", " ");

            return strPriority;
        }
        //convert time to string, separately from the date
        public string GetTimeToString()
        {
            string strPriority = DT_Date.ToShortTimeString();

            return strPriority;
        }
        //formating out input
        public override string ToString()
        {
            return $"{DT_Date.ToShortDateString(), -20}" +
                $"{GetTimeToString(),-23}" +
                $"{GetPriorityToString(), -22}" +
                $"{strDescription}"; 
        }
    }
}
