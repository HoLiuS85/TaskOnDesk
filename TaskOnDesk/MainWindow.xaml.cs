using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Media.Imaging;
using System.Windows.Media.Effects;

namespace TaskOnDesk
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Grid gMain;
        OutlookHandler outlook;
        
        public MainWindow()
        {
            InitializeComponent();

            gMain = InitMainGrid();
            gMain.Children.Add(InitHeaderPanel(0, "dpTaskHeader", "Tasks", @"res\icon_Task.png"));
            gMain.Children.Add(InitLockPanel(0, "dpLockHeader", @"res\icon_LockOpen.png"));
            gMain.Children.Add(InitListView(1, "lvTask"));
            gMain.Children.Add(InitHeaderPanel(2, "dpCalendarHeader", "Calendar", @"res\icon_Calendar.png"));
            gMain.Children.Add(InitListView(3, "lvCalendar"));
            Content = gMain;
            Background = Data.Window.background;

            outlook = new OutlookHandler();
            outlook.OnTaskUpdated += OnTaskUpdated_Outlook;
            outlook.OnCalendarUpdated += OnCalendarUpdated_Outlook;
        }

        private Grid InitMainGrid()
        {
            Grid gTemp = new Grid();
            gTemp.Margin = new Thickness(10,10,10,10);
            gTemp.VerticalAlignment = VerticalAlignment.Top;
            gTemp.HorizontalAlignment = HorizontalAlignment.Right;

            for(int i = 1;i<=4;i++)
                gTemp.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            
            return gTemp;
        }

        private ListView InitListView(int row, string name)
        {
            ListView lvTemp = new ListView()
            {
                Name = name,
                Height = double.NaN,
                Width = double.NaN,
                HorizontalAlignment = Data.ContentAlignment,
                HorizontalContentAlignment = Data.ContentAlignment,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0, 5, 0, 5),
                BorderBrush = Brushes.Transparent,
                Foreground = Brushes.White
            };
            Grid.SetRow(lvTemp, row);

            return lvTemp;

        }

        private DockPanel InitLockPanel(int row, string name, string imagepath)
        {
            Image imgLogo = new Image();
            imgLogo.Source = new BitmapImage(new Uri(imagepath, UriKind.Relative));
            imgLogo.Width = 20;
            imgLogo.Height = 20;
            imgLogo.VerticalAlignment = VerticalAlignment.Center;

            DockPanel dpTemp = new DockPanel();
            dpTemp.Name = name;
            dpTemp.HorizontalAlignment = HorizontalAlignment.Left;
            dpTemp.Children.Add(imgLogo);
            dpTemp.MouseDown += OnClick_Lock;
            Grid.SetRow(dpTemp, row);

            return dpTemp;
        }

        private DockPanel InitHeaderPanel(int row, string name, string text, string imagepath)
        {
            Image imgLogo = new Image();
            imgLogo.Source = new BitmapImage(new Uri(imagepath, UriKind.Relative));
            imgLogo.Width = 20;
            imgLogo.Height = 20;
            imgLogo.VerticalAlignment = VerticalAlignment.Center;
            
            TextBlock tbTemp = new TextBlock();
            tbTemp.Text = text;
            tbTemp.TextDecorations = TextDecorations.Underline;
            tbTemp.Margin = new Thickness(0, 0, 5, 0);
            tbTemp.VerticalAlignment = VerticalAlignment.Center;
            tbTemp.Foreground = Brushes.White;
            tbTemp.FontWeight = FontWeights.Bold;
            tbTemp.FontSize = 14f;
            
            DockPanel dpTemp = new DockPanel();
            dpTemp.Name = name;
            dpTemp.HorizontalAlignment = HorizontalAlignment.Right;
            dpTemp.Children.Add(tbTemp);
            dpTemp.Children.Add(imgLogo);
            dpTemp.PreviewMouseLeftButtonDown += OnClick_Header;
            Grid.SetRow(dpTemp, row);

            return dpTemp;
        }

        private void OnCalendarUpdated_Outlook(object sender, CalendarUpdatedEventArgs e)
        {
            DateTime dtCurrent = DateTime.Now.AddDays(-1);

            Dispatcher.Invoke(new System.Action(() =>
            {
                ListView lvCalendar = LogicalTreeHelper.FindLogicalNode(gMain, "lvCalendar") as ListView;
                lvCalendar.Items.Clear();

                foreach (Calendaritem item in e.List)
                {                    
                    if (item.starttime.DayOfYear > dtCurrent.DayOfYear)
                    {
                        dtCurrent = item.starttime;

                        TextBlock tbTemp = new TextBlock()
                        {
                            Text = dtCurrent.ToString("dddd, dd. MMMM yyyy"),
                            FontWeight = FontWeights.Bold
                        };

                        lvCalendar.Items.Add(new ListViewItem() { Content = tbTemp });   
                    }

                    TextBlock tbAppointment = new TextBlock();
                    tbAppointment.Text = item.starttime.ToString("HH:mm") + " - " + item.endtime.ToString("HH:mm") + " | " + item.subject + " (" + item.location + ")";
                    if (item.starttime < DateTime.Now && item.endtime > DateTime.Now)
                    {
                        tbAppointment.Foreground = Brushes.LightPink;
                        tbAppointment.FontWeight = FontWeights.Bold;
                    }

                    ListViewItem lviTemp = new ListViewItem();
                    lviTemp.Content = tbAppointment;
                    lviTemp.MouseDoubleClick += OnClick_ListViewItem;
                    lviTemp.Tag = item.EntryID;
                    lvCalendar.Items.Add(lviTemp);
                }
            }));
            
        }

        private void OnTaskUpdated_Outlook(object sender, TaskUpdatedEventArgs e)
        {
            Dispatcher.Invoke(new System.Action(() =>
            {
                ListView lvTask = LogicalTreeHelper.FindLogicalNode(gMain, "lvTask") as ListView;
                lvTask.Items.Clear();

                foreach (Taskitem item in e.List)
                {
                    TextBlock tbTemp = new TextBlock();
                    tbTemp.Text = item.duedate.ToShortDateString() + " | " + item.subject;

                    if (item.duedate > DateTime.Now)
                        tbTemp.Foreground = Data.Task.ColorFuture;
                    
                    if (item.duedate.DayOfYear > DateTime.Now.DayOfYear)
                        tbTemp.Foreground = Data.Task.ColorToday;

                    if (item.duedate < DateTime.Now)
                        tbTemp.Foreground = Data.Task.ColorPast;
                    
                    ListViewItem lviTemp = new ListViewItem();
                    lviTemp.Content = tbTemp;
                    lviTemp.MouseDoubleClick += OnClick_ListViewItem;
                    lviTemp.Tag = item.EntryID;
                    lvTask.Items.Add(lviTemp);
                }
            }));
        }
  
        #region UI Events
        private void OnClick_ListViewItem(object sender, MouseButtonEventArgs e)
        {
            ListViewItem lviTemp = sender as ListViewItem;
            ListView lvTemp = lviTemp.Parent as ListView;
            
            if (lvTemp.Name.Equals("lvTask"))
                outlook.OpenOutlookItem(OlItemType.olTaskItem, lviTemp.Tag.ToString());

            if (lvTemp.Name.Equals("lvCalendar"))
                outlook.OpenOutlookItem(OlItemType.olAppointmentItem, lviTemp.Tag.ToString());
        }

        private void OnClick_Header(object sender, MouseButtonEventArgs e)
        {
            if (((DockPanel)sender).Name.Equals("dpTaskHeader"))
                outlook.NewOutlookItem(OlItemType.olTaskItem);

            if (((DockPanel)sender).Name.Equals("dpCalendarHeader"))
                outlook.NewOutlookItem(OlItemType.olAppointmentItem);
        }

        private void OnClick_Lock(object sender, MouseButtonEventArgs e)
        {
            BlurBitmapEffect myBlurEffect = new BlurBitmapEffect();
            myBlurEffect.Radius = 10;
            myBlurEffect.KernelType = KernelType.Box;

            foreach (UIElement uiElement in gMain.Children)
            {
                if (uiElement.BitmapEffect == null)
                {
                    //If Listview Element, apply Blur
                    if (uiElement is ListView)
                    {
                        uiElement.BitmapEffect = myBlurEffect;
                    }
                    
                    if (uiElement is DockPanel)
                    {
                        DockPanel dpTemp = uiElement as DockPanel;

                        if (dpTemp.Name.Equals("dpCalendarHeader") || dpTemp.Name.Equals("dpTaskHeader"))
                        {
                            dpTemp.BitmapEffect = myBlurEffect;
                        }

                        if (dpTemp.Name.Equals("dpLockHeader"))
                        {
                            Image imgTemp = dpTemp.Children[0] as Image;

                            if (imgTemp.Source.ToString().ToLower().Contains("open"))
                            {
                                imgTemp.Source = new BitmapImage(new Uri(@"res\icon_LockClosed.png", UriKind.Relative));
                            }
                            else
                            {
                                imgTemp.Source = new BitmapImage(new Uri(@"res\icon_LockOpen.png", UriKind.Relative));
                            }
                        }
                    }
                }
                else
                {
                    uiElement.BitmapEffect = null;
                }
            }
        }

        private void SizeChanged_Window(object sender, SizeChangedEventArgs e)
        {
            double vTop = Data.Window.PosOffset.Y;
            double vCenter = (SystemParameters.WorkArea.Height / 2) - (ActualHeight / 2);
            double vBottom = SystemParameters.WorkArea.Height - ActualHeight - Data.Window.PosOffset.Y;
            double hLeft = Data.Window.PosOffset.X;
            double hCenter = (SystemParameters.WorkArea.Width / 2) - (ActualWidth / 2);
            double hRight = SystemParameters.WorkArea.Width - ActualWidth - Data.Window.PosOffset.X;

            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Right) && Data.Window.PosVertical.Equals(VerticalAlignment.Top))
            {
                Left = hRight;
                Top = vTop;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Right) && Data.Window.PosVertical.Equals(VerticalAlignment.Center))
            {
                Left = hRight;
                Top = vCenter;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Right) && Data.Window.PosVertical.Equals(VerticalAlignment.Bottom))
            {
                Left = hRight;
                Top = vBottom;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Center) && Data.Window.PosVertical.Equals(VerticalAlignment.Top))
            {
                Left = hCenter;
                Top = vTop;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Center) && Data.Window.PosVertical.Equals(VerticalAlignment.Center))
            {
                Left = hCenter;
                Top = vCenter;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Center) && Data.Window.PosVertical.Equals(VerticalAlignment.Bottom))
            {
                Left = hCenter;
                Top = vBottom;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Left) && Data.Window.PosVertical.Equals(VerticalAlignment.Top))
            {
                Left = hLeft;
                Top = vTop;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Left) && Data.Window.PosVertical.Equals(VerticalAlignment.Center))
            {
                Left = hLeft;
                Top = vCenter;
            }
            if (Data.Window.PosHorizontal.Equals(HorizontalAlignment.Left) && Data.Window.PosVertical.Equals(VerticalAlignment.Bottom))
            {
                Left = hLeft;
                Top = vBottom;
            }
        }
        #endregion

    }
}
