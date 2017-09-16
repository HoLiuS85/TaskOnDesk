using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace TaskOnDesk
{
    public static class Data
    {
        public static HorizontalAlignment ContentAlignment = HorizontalAlignment.Right;

        public static class Task
        {
            public static SolidColorBrush ColorFuture = new SolidColorBrush() { Color = Colors.LightGreen };
            public static SolidColorBrush ColorToday = new SolidColorBrush() { Color = Colors.White };
            public static SolidColorBrush ColorPast = new SolidColorBrush() { Color = Colors.LightPink };
            public static int CheckInterval = 10000;
        }

        public static class Calendar
        {
            public static Int32 DaysToInclude = 7;
            public static int CheckInterval = 10000;
        }

        public static class Window
        {
            public static HorizontalAlignment PosHorizontal = HorizontalAlignment.Right;
            public static VerticalAlignment PosVertical = VerticalAlignment.Top;
            public static Point PosOffset = new Point(10,10);
            public static SolidColorBrush background = new SolidColorBrush() { Color = Colors.Black, Opacity = 0.7f };
        }
    }

    #region Custom EventArgs
    public class TaskUpdatedEventArgs : EventArgs
    {
        public List<Taskitem> List { get; private set; }

        public TaskUpdatedEventArgs(List<Taskitem> list)
        {
            List = list;
        }
    }

    public class CalendarUpdatedEventArgs : EventArgs
    {
        public List<Calendaritem> List { get; private set; }

        public CalendarUpdatedEventArgs(List<Calendaritem> list)
        {
            List = list;
        }
    }
    #endregion

    #region Custom OutlookItems
    public class Calendaritem : OutlookItem
    {
        public DateTime starttime { get; private set; }
        public DateTime endtime { get; private set; }
        public String location { get; private set; }

        public Calendaritem(DateTime start, DateTime end, String subject, String location, String id)
        {
            this.starttime = start;
            this.endtime = end;
            this.subject = subject;
            this.EntryID = id;
            this.location = location;
        }
    }

    public class Taskitem : OutlookItem
    {
        public DateTime duedate { get; private set; }

        public Taskitem(DateTime due, String subject, String id)
        {
            this.duedate = due;
            this.subject = subject;
            this.EntryID = id;
        }
    }

    public class OutlookItem
    {
        public String subject { get; protected set; }
        public String EntryID { get; protected set; }
    }
    #endregion

}
