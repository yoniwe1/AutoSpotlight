using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Media;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace SimulateSpotlite
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int JumpsSize = 1000;
        private double _accumulatedTime = 0;
        private StreamReader reader = new StreamReader(@"SonataBeethoven.csv");
        //MainSringQuartet
        //SonataBeethoven
        private StreamReader secondReader = new StreamReader(@"SonataBeethoven.csv");
        private double _currentIntervalEndTime = 0;
        DispatcherTimer _dt = new DispatcherTimer();
        DispatcherTimer _timer = new DispatcherTimer();
        Stopwatch _sw = new Stopwatch();
        string _currentTime = string.Empty;
        MediaPlayer player = new MediaPlayer();
        private UIElement _timerTextBlock;
        private PropertyInfo _info;
        private Grid _myGrid;
        TimeSpan _ts = new TimeSpan(0);

        public MainWindow()
        {
            InitializeComponent();
            var myGrid = new Grid();
            CreateGridDynamically(myGrid);
            var firstLine = reader.ReadLine();
            secondReader.ReadLine();
            secondReader.ReadLine();
            if (firstLine != null)
            {
                var words = firstLine.Split(',');
                AddColumnsToGrid(words.Length, myGrid);
                AddLabelsToGrid(words, myGrid);
            }

            Main.Content = myGrid;
            _myGrid = myGrid;
            Main.Show();

            var uri = new Uri(@"Opus012-Violon01_1_3Mvt.mid", UriKind.RelativeOrAbsolute);
            //QuatuorCordes01_Opus18_Num1_Mvt4
            //Opus012-Violon01_1_3Mvt
            player.Open(uri);
            player.Play();

            Thread.Sleep(1000);

            ReadNextLine(myGrid);

            

            _timer.Tick += (s, args) => UpdateCurrentMainChannel();
            _timer.Interval = new TimeSpan(0, 0, 0, 0, JumpsSize);
            _timer.Start();

            Thread.Sleep(1000);

            _dt.Tick += (s, args) => UpdateStopper();
            _dt.Interval = new TimeSpan(0, 0, 0, 0, 1);
            _sw.Start();
            _dt.Start();


            _timerTextBlock = myGrid.Children[1];
            _info = myGrid.Children[0].GetType().GetProperty("Text");
        }

        private void UpdateCurrentMainChannel()
        {
            _accumulatedTime += 1D;
            if (_accumulatedTime > _currentIntervalEndTime)
            {
                ReadNextLine(_myGrid);
            }
        }

        private void UpdateStopper()
        {
            if (!_sw.IsRunning) return;
            _ts = _sw.Elapsed;
            _currentTime = $"{_ts.Minutes:00}:{_ts.Seconds:00}:{_ts.Milliseconds / 10:00}";
            _info?.SetValue(_timerTextBlock, _currentTime);
        }

        private void AddLabelsToGrid(string[] words, Grid myGrid)
        {
            TextBlock dynamicHeadline = new TextBlock();

            dynamicHeadline.Name = $"headlineblock";

            dynamicHeadline.TextAlignment = TextAlignment.Center;

            dynamicHeadline.Foreground = new SolidColorBrush(Colors.Chocolate);

            dynamicHeadline.FontSize = 40;

            dynamicHeadline.FontFamily = new FontFamily("Comic Sans MS");

            dynamicHeadline.HorizontalAlignment = HorizontalAlignment.Center;

            dynamicHeadline.VerticalAlignment = VerticalAlignment.Center;

            dynamicHeadline.Inlines.Add("Automatic Spotlight Simulation");

            dynamicHeadline.Inlines.Add(new LineBreak());

            Run subTitle = new Run("This is a program to simulate automation of a spotlight in a concert");

            subTitle.FontSize = 30;

            subTitle.Foreground = new SolidColorBrush(Colors.Black);

            subTitle.FontStyle = FontStyles.Italic;

            subTitle.BaselineAlignment = BaselineAlignment.Center;

            dynamicHeadline.Inlines.Add(subTitle);

            dynamicHeadline.Inlines.Add(new LineBreak());

            Run explenation = new Run("The simulation will mark the proper dominant channel based on a unique calculation made for the String Quartet classical genre");

            explenation.FontSize = 20;

            explenation.Foreground = new SolidColorBrush(Colors.Black);

            explenation.BaselineAlignment = BaselineAlignment.Center;

            dynamicHeadline.Inlines.Add(explenation);

            Grid.SetRow(dynamicHeadline, 0);

            Grid.SetColumn(dynamicHeadline, 0);

            Grid.SetColumnSpan(dynamicHeadline, words.Length - 1);

            myGrid.Children.Add(dynamicHeadline);

            TextBlock dynamicTextBlock = new TextBlock();

            dynamicTextBlock.Name = $"clocktxtblock";

            //dynamicTextBlock.Background = new SolidColorBrush(Colors.DarkSlateBlue);

            dynamicTextBlock.FontSize = 30;

            dynamicTextBlock.FontFamily = new FontFamily("Comic Sans MS");

            dynamicTextBlock.HorizontalAlignment = HorizontalAlignment.Center;

            dynamicTextBlock.VerticalAlignment = VerticalAlignment.Center;

            Grid.SetRow(dynamicTextBlock, 2);

            Grid.SetColumn(dynamicTextBlock, (int) (Math.Floor(((double) (words.Length - 2)) / 2)));

            Grid.SetColumnSpan(dynamicTextBlock, 2);

            myGrid.Children.Add(dynamicTextBlock);

            for (var i = 1; i < words.Length; i++)
            {

                StackPanel myStackPanel = new StackPanel();
                myStackPanel.Orientation = Orientation.Vertical;

                Image myImage = new Image();
                BitmapImage myImageSource = new BitmapImage();
                myImageSource.BeginInit();
                myImageSource.UriSource = new Uri(@"piano1.PNG", UriKind.RelativeOrAbsolute);
                myImageSource.EndInit();
                myImage.Source = myImageSource;
                myImage.Visibility = Visibility.Hidden;

                myImage.HorizontalAlignment = HorizontalAlignment.Center;

                myImage.VerticalAlignment = VerticalAlignment.Center;

                myImage.Margin = new Thickness(5);

                myImage.Height = 220;



                Label dynamicLabel = new Label();

                var channelNumber = "";

                var channelName = "";

                if (words[i] != null)
                {
                    channelNumber = words[i].Split('_')[0];

                    channelName = words[i].Split('_')[1];
                }

                dynamicLabel.Name = $"Excel_column_{i}";

                dynamicLabel.HorizontalContentAlignment = HorizontalAlignment.Center;

                dynamicLabel.HorizontalAlignment = HorizontalAlignment.Center;

                dynamicLabel.Content = "Channel " + channelNumber + "\n" + channelName;

                dynamicLabel.HorizontalAlignment = HorizontalAlignment.Center;

                dynamicLabel.VerticalAlignment = VerticalAlignment.Center;

                dynamicLabel.FontSize = 30;

                dynamicLabel.FontWeight = FontWeights.Bold;

                dynamicLabel.Margin = new Thickness(10);

                dynamicLabel.FontFamily = new FontFamily("Comic Sans MS");

                myStackPanel.Children.Add(myImage);

                myStackPanel.Children.Add(dynamicLabel);

                Grid.SetRow(myStackPanel, 1);

                Grid.SetColumn(myStackPanel, i - 1);

                myGrid.Children.Add(myStackPanel);
            }
        }

        private void AddColumnsToGrid(int wordsLength, Grid myGrid)
        {
            for (var i = 0; i < wordsLength - 1; i++)
            {
                ColumnDefinition gridCol = new ColumnDefinition();
                myGrid.ColumnDefinitions.Add(gridCol);
            }

        }

        private void CreateGridDynamically(Grid myGrid)
        {

            myGrid.Background = new SolidColorBrush(Colors.LightSteelBlue);

            RowDefinition gridRow1 = new RowDefinition();
            gridRow1.Height = new GridLength(30, GridUnitType.Star);

            RowDefinition gridRow2 = new RowDefinition();
            gridRow2.Height = new GridLength(50, GridUnitType.Star);

            RowDefinition gridRow3 = new RowDefinition();
            gridRow3.Height = new GridLength(20, GridUnitType.Star);

            myGrid.RowDefinitions.Add(gridRow1);

            myGrid.RowDefinitions.Add(gridRow2);

            myGrid.RowDefinitions.Add(gridRow3);

        }

        private void ReadNextLine(Grid myGrid)
        {
            var line = reader.ReadLine();
            
            if (line != null)
            {
                var values = line.Split(',');
                if (values[1] == "Ending")
                {
                    StopProcess();
                }
                var nextTime = secondReader.ReadLine()?.Split(',')[0];
                _currentIntervalEndTime = double.Parse(nextTime ?? throw new InvalidOperationException(), System.Globalization.CultureInfo.InvariantCulture);
                var columnNumber = 0;
                for (int counter = 1; counter < values.Length; counter++)
                {
                    if (Int32.Parse(values[counter]) == 1)
                    {
                        columnNumber = counter;
                        break;
                    }
                }
                for (int i = 2; i < myGrid.Children.Count; i++)
                {
                    UIElement e = myGrid.Children[i];
                    int excelColumn = Int32.Parse(((StackPanel)(e)).Children[1].GetValue(FrameworkElement.NameProperty).ToString().Split('_')[2]);
                    if (columnNumber == excelColumn)
                    {
                        EnlightChannel(myGrid, i);
                        break;
                    }
                }
            }
        }

        private void StopProcess()
        {
            _sw.Stop();
            _dt.Stop();
            _timer.Stop();
            Main.Close();
        }

        private void EnlightChannel(Grid myGrid, int childNumToLight)
        {
            for (int i = 2; i < myGrid.Children.Count; i++)
            {
                if (i != childNumToLight)
                {
                    ((StackPanel)(myGrid.Children[i])).Children[0].GetType().GetProperty("Visibility")?.SetValue((((StackPanel)(myGrid.Children[i])).Children[0]), Visibility.Hidden);
                }
            }
            ((StackPanel)(myGrid.Children[childNumToLight])).Children[0].GetType().GetProperty("Visibility")?.SetValue(((StackPanel)(myGrid.Children[childNumToLight])).Children[0], Visibility.Visible);
        }
    }
}

