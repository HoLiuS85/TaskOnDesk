using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TaskOnDesk
{
    class OutlookHandler
    {
        public delegate void CalendarUpdateHandler(object sender, CalendarUpdatedEventArgs e);
        public delegate void TaskUpdateHandler(object sender, TaskUpdatedEventArgs e);
        public event CalendarUpdateHandler OnCalendarUpdated;
        public event TaskUpdateHandler OnTaskUpdated;

        private List<Taskitem> lTasks = new List<Taskitem>();
        private List<Calendaritem> lAppointments = new List<Calendaritem>();
        private Timer tCheckTask, tCheckCalendar;
        private Application application;

        public OutlookHandler()
        {
            //Create Timer to periodically check for new Tasks and Appointments
            tCheckTask = new Timer(new TimerCallback(CheckTaskUpdate), null, 0, -1);
            tCheckCalendar = new Timer(new TimerCallback(CheckCalendarUpdate), null, 0, -1);
        }


        private void CheckTaskUpdate(object state)
        {
            List<Taskitem> lTasksNew = GetTaskItems();

            if (!lTasksNew.All(lTasks.Contains))
            {
                lTasks = lTasksNew;

                if (OnTaskUpdated != null)
                    OnTaskUpdated(this, new TaskUpdatedEventArgs(lTasks));
            }

            tCheckTask.Change(10000, -1);
        }

        private void CheckCalendarUpdate(object state)
        {
            List<Calendaritem> lCalendarNew = GetCalendarItems();

            if (!lCalendarNew.All(lAppointments.Contains))
            {
                lAppointments = lCalendarNew;

                if (OnCalendarUpdated != null)
                    OnCalendarUpdated(this, new CalendarUpdatedEventArgs(lAppointments));
            }
            tCheckCalendar.Change(10000, -1);
        }

        private Application GetOutlookApplication()
        {
            //If we already have a valid instance of Outlook, return it
            if (application != null)
                return application;

            //Check if Outlook is running. 
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                try
                {
                    //If so, obtain the process and cast it to an Application object. 
                    application = Marshal.GetActiveObject("Outlook.Application") as Application;
                }
                catch
                { 
                    //If an error happens while Marshaling Outlook, create a new instance
                    application = new Application();
                }
            }
            else
            {
                //If not, create a new instance of Outlook. 
                application = new Application();
            }

            return application;
        }

        private NameSpace GetOutlookNameSpace()
        {
            NameSpace nameSpace = null;

            //Get the Default Namespace of the Outlook Application
            nameSpace = GetOutlookApplication().GetNamespace("MAPI");

            //Log on with the default outlook profile in the background
            nameSpace.Logon(Missing.Value, Missing.Value, false, false);

            return nameSpace;
        }

        private List<Calendaritem> GetCalendarItems()
        {
            List<Calendaritem> lTemp = new List<Calendaritem>();

            try
            {
                //Set Start to Today midnight to include all appointments of the day
                DateTime dtStart = DateTime.Now.Date;

                //Set End to value configured by the user
                DateTime dtEnd = DateTime.Now.AddDays(Data.Calendar.DaysToInclude);

                //Get all Calendar items from the default outlook calendar and include recurings before sorting
                Items calItems = GetOutlookNameSpace().GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);

                //Restrict Appointments to items matching the filter (between dtStart and dtEnd) ant iterate through
                Items calItemsRestricted = calItems.Restrict(string.Format("[Start] >= '" + dtStart.ToString("g") + "' AND [End] <= '" + dtEnd.ToString("g") + "'"));

                foreach (AppointmentItem calItem in calItemsRestricted)
                {
                    //Check if item is Appointment (reduces errors)
                    if (calItem is AppointmentItem)
                    {
                        //Check wether the appointment is recurring
                        if (calItem.IsRecurring)
                        {
                            //Create Timestemp which matches to current appointment as recurence pattern
                            DateTime first = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, calItem.Start.Hour, calItem.Start.Minute, 0);

                            //Iterate through all days between dtStart and dtEnd to check if the appointment is found
                            for (DateTime cur = first; cur <= dtEnd; cur = cur.AddDays(1))
                            {
                                try
                                {
                                    //Get the occurence of the appointment. This causes an exception if not found in current day
                                    AppointmentItem calItemRecur = calItem.GetRecurrencePattern().GetOccurrence(cur);

                                    //Verify that the appointment is not over yet and add it to the collection
                                    if (calItemRecur.End > DateTime.Now)
                                        lTemp.Add(new Calendaritem(calItemRecur.Start, calItemRecur.End, calItemRecur.Subject, calItemRecur.Location, calItemRecur.EntryID));
                                }
                                catch { /* Appointment not found in current day, maybe tomorrow */  }
                            }
                        }
                        else
                        {
                            //Verify that the appointment is not over yet and add it to the collection
                            if (calItem.End > DateTime.Now)
                                lTemp.Add(new Calendaritem(calItem.Start, calItem.End, calItem.Subject, calItem.Location, calItem.EntryID));
                        }
                    }
                }

                //Sort list by ascending start date and return
                return lTemp.OrderBy(x => x.starttime).ToList();
            }
            catch
            {
                //If an error happens, return the previous list (might happen if outlook is closed)
                return lAppointments;
            }
        }

        private List<Taskitem> GetTaskItems()
        {
            List<Taskitem> lTemp = new List<Taskitem>();

            try
            {
                //Get all Task items from the default outlook Tasks and sort them    
                Items tskItems = GetOutlookNameSpace().GetDefaultFolder(OlDefaultFolders.olFolderTasks).Items;
                tskItems.Sort("[Start]", Type.Missing);

                //Iterate through all Tasks
                foreach (TaskItem tskItem in tskItems)
                {
                    //Verify that the task is not completed yet and add it to the collection
                    if (!tskItem.Complete)
                        lTemp.Add(new Taskitem(tskItem.DueDate, tskItem.Subject, tskItem.EntryID));
                }

                //Sort list by ascending due date and return
                return lTemp.OrderBy(x => x.duedate).ToList();
            }
            catch (System.Exception e)
            {
                //If an error happens, return the previous list (might happen if outlook is closed)
                Debug.WriteLine(e.ToString() + " Message: " + e.Message);
                return lTasks;
            }
        }

        public void NewOutlookItem(OlItemType type)
        {
            try
            {
                //Create new Outlook Task and display UI to user
                if (type is OlItemType.olTaskItem)
                    (GetOutlookApplication().CreateItem(type) as TaskItem).Display();

                //Create new Outlook Appointment and display UI to user
                if (type is OlItemType.olAppointmentItem)
                    (GetOutlookApplication().CreateItem(type) as AppointmentItem).Display();
            }
            catch { /*Might happen if outlook is closed or item not exists, bad luck*/ }
        }

        public void OpenOutlookItem(OlItemType type, string EntryID)
        {
            try
            {
                //Open existing Task Item by EntryId
                if (type is OlItemType.olTaskItem)
                    (GetOutlookNameSpace().GetItemFromID(EntryID) as TaskItem).Display();

                //Open existing Appointment Item by EntryId
                if (type is OlItemType.olAppointmentItem)
                    (GetOutlookNameSpace().GetItemFromID(EntryID) as AppointmentItem).Display();
            }
            catch { /*Might happen if outlook is closed or item not exists, bad luck*/ }
        }
    }
}
