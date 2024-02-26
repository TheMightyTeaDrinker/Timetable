using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using Microsoft.VisualBasic.FileIO;
using System.Timers;

namespace MinimizeToSystemTray
{
    class Program
    {
        static NotifyIcon notifyIcon;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            InitializeNotifyIcon();

            string[] timeStrings = { "8:30", "9:25", "10:35", "12:00", "13:25", "15:05" };

            // Parse times into DateTime objects
            DateTime[] times = timeStrings.Select(timeStr => DateTime.ParseExact(timeStr, "H:mm", null)).ToArray();

            if (times.Any())
            {
                DateTime timeNow = DateTime.Now;
                foreach (DateTime d in times)
                {
                    d.AddMinutes(1);
                    System.Timers.Timer t = new (Math.Abs((int)(timeNow - d).TotalMilliseconds));
                    t.Elapsed += NotifyIcon_DoubleClick;
                    t.Enabled = true;
                }
            }
            NotifyIcon_DoubleClick(new object(), new EventArgs());

            Application.Run();
        }

        private static void InitializeNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = Timetable.Properties.Resources.icon; // Use a default icon for simplicity
            notifyIcon.Text = "Timetable Background App";
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            // Create context menu items
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("Refresh", null, NotifyIcon_DoubleClick);
            contextMenu.Items.Add("Exit", null, exitToolStripMenuItem_Click);

            notifyIcon.ContextMenuStrip = contextMenu;

            notifyIcon.Visible = true; // Make the NotifyIcon visible by default
        }

        private static void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            const int SPI_SETDESKWALLPAPER = 0x0014;
            const int SPIF_UPDATEINIFILE = 0x01;
            const int SPIF_SENDCHANGE = 0x02;

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

            string imagePath = AppDomain.CurrentDomain.BaseDirectory + "Data/" + "Wallpaper.jpg";
            
            Image originalImage = Image.FromFile(imagePath);

            DateTime currentTime = DateTime.Now;

            string[] timeStrings = { "8:30", "9:25", "10:35", "12:00", "13:25", "15:05" };

            // Parse times into DateTime objects
            DateTime[] times = timeStrings.Select(timeStr => DateTime.ParseExact(timeStr, "H:mm", null)).ToArray();

            // Get day number of the week
            int day = ((int)currentTime.DayOfWeek); //Monday = 1

            // Get the week number of the current date
            // Ensure that the specified date is interpreted according to ISO 8601 rules
            System.Globalization.Calendar calendar = System.Globalization.CultureInfo.CurrentCulture.Calendar;
            int weekNumber = calendar.GetWeekOfYear(currentTime, System.Globalization.CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            bool isEven = (weekNumber + 1) % 2 == 0;
            string csvFilePath;

            if (isEven)
            {
                csvFilePath = AppDomain.CurrentDomain.BaseDirectory + "Data/Even.csv";
            } else
            {
                csvFilePath = AppDomain.CurrentDomain.BaseDirectory + "Data/Odd.csv";
            }

            using (Graphics g = Graphics.FromImage(originalImage))
            {
                Point position = new Point(3300, 1500);
                Size size = new(500, 560);
                Rectangle rectangle = new(position, size);
                int textSize = 30;

                // Define the rectangle pen and brush
                Pen rectanglePen = new Pen(Color.Black, 5);
                SolidBrush rectangleBrush = new (Color.White);

                // Define the text parameters (adjust as needed)
                Font textFont = new Font("Arial", textSize);
                Brush textBrush = new SolidBrush(Color.Black);

                int targetColumnIndex = day - 1; // Replace with the desired column index (0-based)
                int startRowIndex = 1; // Replace with the desired start row index (0-based)
                int endRowIndex = 15; // Replace with the desired end row index (0-based)

                string[] selectedColumnValues = new string[endRowIndex - startRowIndex + 1];

                // Check if the CSV file exists
                if (File.Exists(csvFilePath))
                {
                    // Read the CSV file
                    using (TextFieldParser parser = new TextFieldParser(csvFilePath))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");

                        // Move to the start row
                        for (int i = 0; i < startRowIndex; i++)
                        {
                            parser.ReadFields();
                        }

                        // Process each row in the specified range
                        for (int rowIndex = startRowIndex; rowIndex <= endRowIndex && !parser.EndOfData; rowIndex++)
                        {
                            string[] fields = parser.ReadFields();

                            // Check if the target column index is within the array bounds
                            if (targetColumnIndex >= 0 && targetColumnIndex < fields.Length)
                            {
                                selectedColumnValues[rowIndex - startRowIndex] = fields[targetColumnIndex];
                            }
                        }
                    }
                }

                //Debug.WriteLine(selectedColumnValues.Length);

                int start = 0;

                // Filter out times that are not above the current time
                var validTimes = times.Where(time => time > currentTime);

                int index;
                if (validTimes.Any())
                {
                    // Find the minimum among valid times
                    DateTime nextClosestTime = validTimes.Min();

                    index = Array.IndexOf(times, nextClosestTime);
                }
                else index = -1;

                if (index > 0)
                {
                    // Draw the rectangle on the image
                    g.FillRectangle(rectangleBrush, rectangle);
                    g.DrawRectangle(rectanglePen, rectangle);

                    int spacing = 15;
                    for (int i = 0; i < 9; i++)
                    {
                        //Debug.WriteLine(index);

                        if (i % 3 == 0)
                        {
                            spacing += 5;
                        }

                        // Define the text position and draw it
                        PointF textPosition = new PointF(position.X + 10, position.Y + 10 + (i * (textSize + spacing)));
                        int listNum = i + ((index - 1) * 3);
                        //Debug.WriteLine(listNum + " " + selectedColumnValues.Length);
                        if (listNum < 15)
                        {
                            g.DrawString(selectedColumnValues[listNum], textFont, textBrush, textPosition);
                        }
                        else break;
                    }
                } else if (index == 0) // Before School
                {
                    Rectangle newRectangle = new(position, new Size(500, 70));
                    g.FillRectangle(rectangleBrush, newRectangle);
                    g.DrawRectangle(rectanglePen, newRectangle);
                    g.DrawString("Before School", textFont, textBrush, new PointF(position.X + 10, position.Y + 10));
                }
                else // After School
                {
                    Rectangle newRectangle = new(position, new Size(500, 105));
                    g.FillRectangle(rectangleBrush, newRectangle);
                    g.DrawRectangle(rectanglePen, newRectangle);
                    g.DrawString("After School", textFont, textBrush, new PointF(position.X + 10, position.Y + 10));
                    if (day < 5)
                    {
                        g.DrawString("Tomorrow is School", textFont, textBrush, new PointF(position.X + 10, position.Y + 20 + textSize));
                    }
                    else
                    {
                        g.DrawString("Tomorrow is Weekend", textFont, textBrush, new PointF(position.X + 10, position.Y + 20 + textSize));
                    }
                }
            }

            string outputImagePath = AppDomain.CurrentDomain.BaseDirectory + "Data/" + "Wallpaper - Edited.jpg";
            originalImage.Save(outputImagePath, ImageFormat.Jpeg);

            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, outputImagePath, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }

        private static void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Add any cleanup logic if needed
            notifyIcon.Dispose();
            Application.Exit();
        }
    }
}
