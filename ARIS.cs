using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Recognition.SrgsGrammar;
using System.Runtime.InteropServices;
using System.Management;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Web;
using System.Speech.Synthesis;
using System.Threading;
using System.Net.NetworkInformation;

namespace A.R.I.S_Intelligent_System
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer ARIS = new SpeechSynthesizer();
        string Temperature, Condition, Humidity, WindSpeed, Town, TFCond, TFHigh, TFLow;
        string News1, News2, News3, News4, News5, News6, News7, News8, News9, News10;
        string News11, News12, News13, News14, News15, News16, News17, News18, News19, News20;
        string result;
        DateTime Time = DateTime.Now;
        String QEvent;
        int Timer = 20;
        int ats1 = 300;
        int ats2 = 20;
        int light = 10;
        Random ans = new Random();
        int T = 0;
        int X = 10;

        public Form1()
        {
            try
            {
                InitializeComponent();
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Commands.txt")))));
                recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                recognizer.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRejected);
                recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);
                NetworkChange.NetworkAvailabilityChanged += new NetworkAvailabilityChangedEventHandler(NetworkChange_NetworkAvailabilityChanged);
            }
            catch
            {
                MessageBox.Show("No Microphone Detected , Please Connect a Suitable Microphone");
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            int res;
            string time = Time.GetDateTimeFormats('t')[0];
            string speech = e.Result.Text;
 
            try
            {
                switch (speech)
                {

                    //GREETINGS
                    case "hello aris":
                    case "welcome aris":
                    case "goodmorning aris":
                        if (Time.Hour >= 5 && Time.Hour < 12)
                        { ARIS.SpeakAsync("Goodmorning sir"); }
                        if (Time.Hour >= 12 && Time.Hour < 18)
                        { ARIS.SpeakAsync("Good afternoon sir"); }
                        if (Time.Hour >= 18 && Time.Hour < 24)
                        { ARIS.SpeakAsync("Good evening sir"); }
                        if (Time.Hour < 5)
                        { ARIS.SpeakAsync("Hello, it is getting late sir"); }
                        ARIS.SpeakAsync("what can i do for you");
                        break;
                    case "goodbye aris":
                    case "good night aris":
                    case "see you soon":
                        ARIS.Speak("GoodBye Sir");
                        ARIS.Speak("Until Next Time");
                        Application.Exit();
                        break;
                    case "i love you":
                        ARIS.SpeakAsync("I Love You too");
                        break;


                    //DATE / TIME / WEATHER
                    case "what time is it":
                    case "what is the time":
                    case "what is time":
                        ARIS.SpeakAsync(time);
                        break;
                    case "what day is it":
                    case "what day is this":
                    ARIS.SpeakAsync(DateTime.Today.ToString("dddd"));
                        break;
                    case "what is the date":
                    case "what is today's date":
                    case "what is date":
                        ARIS.SpeakAsync(DateTime.Today.ToString("MM-dd-yyyy"));
                        break;
                    case "how is the weather today":
                    case "what is today's weather":
                    case "weather today":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            GetWeather();
                            ARIS.SpeakAsync("Today's Weather as follows " + " The Temperature is " + Temperature + " Degree fahrenheit " + " The Humidity is " +
                            Humidity + "Percent" + " The Wind Speed is " + WindSpeed + " Miles Per Hour" + " and The Weather Condition is " + Condition);
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "what is tomorrow's weather forecast":
                    case "what is tomorrow's weather":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {

                            GetWeather();
                            ARIS.SpeakAsync("Tomorrows's Weather Forecast as follows " + " The High Temperature is " + TFHigh + " Degree fahrenheit" + "The Low Temperature is" +
                                                TFLow + "Degree fahrenheit" + " and The Weather Condition is " + TFCond);
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;



                    //Programs Launch
                    case "open google chrome":
                        ARIS.SpeakAsync("Right Away");
                        System.Diagnostics.Process.Start("chrome.exe");
                        break;
                    case "open itunes":
                        ARIS.SpeakAsync("Connect Your I Phone");
                        System.Diagnostics.Process.Start("iTunes.exe");
                        break;
                    case "open media player":
                        ARIS.SpeakAsync("Right Away");
                        System.Diagnostics.Process.Start("wmplayer.exe");
                        break;
                    case "open photoshop":
                        ARIS.SpeakAsync("Let's Edit Some Photos");
                        System.Diagnostics.Process.Start("photoshop.exe");
                        break;
                    case "open visual studio":
                        ARIS.SpeakAsync("Let's Begin Coding");
                        System.Diagnostics.Process.Start("devenv.exe");
                        break;
                    case "open free studio":
                        ARIS.SpeakAsync("Ok Sir");
                        System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Common Files\DVDVideoSoft\FreeStudioManager.exe");
                        break;


                    //WEBSITES
                    case "open facebook":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.facebook.com");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open google":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.google.com");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open yahoo":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.yahoo.com");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open youtube":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.youtube.com");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open wikipedia":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.wikipedia.org/");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "how is movies today":
                    case "movies today":
                    case "what is today's movies":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("I Will Check For You Sir");
                            System.Diagnostics.Process.Start("http://www.imdb.com/");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open i jailbreak":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.ijailbreak.com/");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open day 7":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Ok sir");
                            System.Diagnostics.Process.Start("http://www.youm7.com/");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open cnn":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Ok sir");
                            System.Diagnostics.Process.Start("http://www.cnn.com/");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "open accuweather":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Right Away Sir");
                            System.Diagnostics.Process.Start("http://www.accuweather.com/en/eg/al-mansurah/126814/weather-forecast/126814");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;


                    //SHUTDOWN RESTART LOG OFF
                    case "shutdown computer":
                        QEvent = "shutdown";
                        ARIS.SpeakAsync("Your Computer Will Shutdown in 20 Seconds");
                        label4.Visible = true;
                        Timer1.Enabled = true;
                        break;
                    case "log off computer":
                        QEvent = "logoff";
                        ARIS.SpeakAsync("Your Computer Will Logoff in 20 Seconds");
                        label4.Visible = true;
                        Timer1.Enabled = true;
                        break;
                    case "restart computer":
                        QEvent = "restart";
                        ARIS.SpeakAsync("Your Computer Will Restart in 20 Seconds");
                        label4.Visible = true;
                        Timer1.Enabled = true;
                        break;
                    case "abort operation":
                        ARIS.SpeakAsync("Operation Terminated");
                        Timer = 20;
                        label4.Text = "20";
                        Timer1.Enabled = false;
                        label4.Visible = false;
                        break;


                    //Interaction
                    case "stop talking":
                        ARIS.SpeakAsyncCancelAll();
                        break;
                    case "out of the way":
                    case "move a way":
                        if (WindowState == FormWindowState.Normal)
                        {
                            WindowState = FormWindowState.Minimized;
                            ARIS.SpeakAsync("My apologies");
                        }
                        break;
                    case "come back":
                    case "come back aris":
                        if (WindowState == FormWindowState.Minimized)
                        {
                            ARIS.SpeakAsync("Here I am");
                            WindowState = FormWindowState.Normal;
                        }
                        break;
                    case "show commands":
                    case "what can i say":
                    case "commands":
                        string[] commands = File.ReadAllLines(@"Commands.txt");
                        ARIS.SpeakAsync("Here we are");
                        LstBoxCommands.Items.Clear();
                        LstBoxCommands.SelectionMode = SelectionMode.None;
                        LstBoxCommands.Visible = true;
                        foreach (string command in commands)
                        {
                            LstBoxCommands.Items.Add(command);
                        }
                        break;
                    case "hide commands":
                        ARIS.SpeakAsync("Very Well");
                        LstBoxCommands.Visible = false;
                        break;
                    case "go full screen":
                        FormBorderStyle = FormBorderStyle.None;
                        WindowState = FormWindowState.Maximized;
                        TopMost = true;
                        ARIS.SpeakAsync("expanding");
                        break;
                    case "exit full screen":
                        FormBorderStyle = FormBorderStyle.Sizable;
                        WindowState = FormWindowState.Normal;
                        TopMost = false;
                        ARIS.SpeakAsync("exiting");
                        break;


                    //SOCIAL
                    case "how are you":
                    case "how are you doing":
                    case "how are you today":
                    case "are you ok":
                    case "are you fine":
                        res = ans.Next(1, 3);
                        if (res == 1)
                        {
                            ARIS.SpeakAsync("I'm Fine");
                            ARIS.SpeakAsync("How are you sir?");
                        }
                        else if (res == 2)
                        {
                            ARIS.SpeakAsync("I'm Fine");
                            ARIS.SpeakAsync("How are you today?");
                        }
                        break;
                    case "i'm fine":
                    case "i'm ok":
                    case "i'm good":
                        ARIS.SpeakAsync("Glad to hear it");
                        break;
                    case "are you my friend aris":
                        ARIS.SpeakAsync("Of Course Sir");
                        break;
                    case "are you ready aris":
                        ARIS.SpeakAsync("I'm always ready sir");
                        ARIS.SpeakAsync("Let's Begin The Work");
                        break;
                    case "can you hear me":
                        ARIS.SpeakAsync("yes , I Hear You");
                        break;
                    case "tell me about your self":
                        ARIS.SpeakAsync("Hello Every one I'm an Artificial Intelligent system , i'm your personal assistant that helps you in the daily life");
                        break;


                    //Information of Computer
                    case "what is the disk usage percentage":
                    case "disk usage percentage":
                    case "disk usage":
                        double x;
                        x = 100 - (int)PerformanceCounter2.NextValue();
                        ARIS.SpeakAsync("The Disk Usage Percentage is" + x + "Percent");
                        break;
                    case "what is the ram memory usage percentage":
                    case "ram memory usage":
                    case "ram usage percentage":
                        double y;
                        y = (int)PerformanceCounter1.NextValue();
                        ARIS.SpeakAsync("The ram memory usage percentage is" + y + "Percent");
                        break;
                    case "charging status":
                        PowerLineStatus status = SystemInformation.PowerStatus.PowerLineStatus;
                        if (status == PowerLineStatus.Offline)
                        {
                            ARIS.SpeakAsync("Running on Battery");
                        }
                        else
                        {
                            ARIS.SpeakAsync("Running on Charger");
                        }
                        break;
                    case "what is the battery power usage percentage":
                    case "battery power percent":
                        PowerStatus power1 = SystemInformation.PowerStatus;
                        double percent = (power1.BatteryLifePercent) * 100;
                        ARIS.SpeakAsync("The Battery Power Usage Percentage is" + string.Format("{0:0}", percent) + "Percent");
                        break;
                    case "what is the remaining battery time":
                    case "battery remaining time":
                        PowerStatus power2 = SystemInformation.PowerStatus;
                        double c;
                        double d;

                        c = power2.BatteryLifeRemaining / 3600;
                        d = power2.BatteryLifeRemaining / 60;
                        if (c < 1)
                        {
                            ARIS.SpeakAsync("The Battery Remaining Time is" + d + "Minutes");
                        }
                        else
                        {
                            ARIS.SpeakAsync("The Battery Remaining Time is" + c + "Hours");
                        }
                        break;
                    case "what is the available ram size":
                    case "available ram size":
                        double w;
                        w = (int)performanceCounter3.NextValue();
                        ARIS.SpeakAsync("The Available Ram Size is" + w + "Mega Bytes");
                        break;
                    case "tell me the computer's name":
                    case "computer's name":
                        ARIS.SpeakAsync(SystemInformation.ComputerName);
                        break;
                    case "what is the computer's operating system version":
                    case "operating system version":
                        ARIS.SpeakAsync(Environment.OSVersion.VersionString);
                        break;


                    //Multiple Searching
                    case "search videos":
                        QEvent = "video";
                        ARIS.SpeakAsync("Give me Title sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "search photos":
                        QEvent = "photo";
                        ARIS.SpeakAsync("Give me Title sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "search web":
                        QEvent = "web";
                        ARIS.SpeakAsync("Give me Title sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "search wikipedia":
                        QEvent = "wikipedia";
                        ARIS.SpeakAsync("Give me Title sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "i want to buy something":
                        QEvent = "product";
                        ARIS.SpeakAsync("what is it sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "search title":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Searching");
                            TextBox1.Visible = false;
                            if (QEvent == "video")
                            {
                                System.Diagnostics.Process.Start("http://www.youtube.com/results?search_query=" + TextBox1.Text);
                            }
                            else if (QEvent == "photo")
                            {
                                System.Diagnostics.Process.Start("http://www.google.com.eg/search?um=1&hl=ar&biw=1422&bih=735&tbm=isch&sa=1&q=" + TextBox1.Text);
                            }
                            else if (QEvent == "web")
                            {
                                System.Diagnostics.Process.Start("http://www.google.com.eg/#q=" + TextBox1.Text + "&fp=b95952b3235a197d");
                            }
                            else if (QEvent == "wikipedia")
                            {
                                System.Diagnostics.Process.Start("http://en.wikipedia.org/wiki/" + TextBox1.Text);
                            }
                            else if (QEvent == "product")
                            {
                                System.Diagnostics.Process.Start("http://egypt.souq.com/eg-en/" + TextBox1.Text + "/s/");
                            }
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "clear text box":
                        TextBox1.Clear();
                        break;
                    case "cancel search":
                        ARIS.SpeakAsync("Canceling");
                        TextBox1.Clear();
                        TextBox1.Visible = false;
                        break;


                    //ATS Mode
                    case "run a t s mode":
                        ARIS.SpeakAsync("A , T , S mode is Now Running");
                        Timer2.Start();
                        Label2.Visible = true;

                        break;
                    case "yes i am awake":
                        Label2.Text = "300";
                        ats1 = 300;
                        Timer2.Start();
                        Timer3.Stop();
                        ats2 = 20;
                        Label3.Text = "20";
                        Label3.Visible = false;
                        break;
                    case "cancel a t s mode":
                        ARIS.SpeakAsync("A , T , S Mode canceled");
                        Label2.Text = "300";
                        ats1 = 300;
                        Timer2.Stop();
                        Timer3.Stop();
                        Label3.Text = "20";
                        ats2 = 20;
                        Label2.Visible = false;
                        Label3.Visible = false;
                        break;


                    //Media Control
                    case "open file":
                        ARIS.SpeakAsync("Choose Your File Sir");
                        OpenFileDialog1.Title = "Please Select a File";
                        OpenFileDialog1.InitialDirectory = "C";
                        OpenFileDialog1.ShowDialog();
                        break;
                    case "increase volume":
                        ARIS.SpeakAsync("Increasing Volume");
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_UP);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_UP);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_UP);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_UP);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_UP);
                        break;
                    case "decrease volume":
                        ARIS.SpeakAsync("Decreasing Volume");
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                    (IntPtr)APPCOMMAND_VOLUME_DOWN);
                        break;
                    case "close volume":
                        ARIS.SpeakAsync("Closing Volume");
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle,
                 (IntPtr)APPCOMMAND_VOLUME_MUTE);
                        break;


                    //News Feeds
                    case "what is today's news":
                    case "news today":
                    case "what is the news":
                        ARIS.SpeakAsync("Local News or World News");
                        break;
                    case "local news":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            get_local_news();
                            ARIS.SpeakAsync("The Local News For Today is as Follows ");
                            ARIS.SpeakAsync(News1);
                            ARIS.SpeakAsync(News2);
                            ARIS.SpeakAsync(News3);
                            ARIS.SpeakAsync(News4);
                            ARIS.SpeakAsync(News5);
                            ARIS.SpeakAsync(News6);
                            ARIS.SpeakAsync(News7);
                            ARIS.SpeakAsync(News8);
                            ARIS.SpeakAsync(News9);
                            ARIS.SpeakAsync(News10);
                            ARIS.SpeakAsync(" End of The News Report");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;
                    case "world news":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            get_world_news();
                            ARIS.SpeakAsync("The World News For Today is as Follows ");
                            ARIS.SpeakAsync(News11);
                            ARIS.SpeakAsync(News12);
                            ARIS.SpeakAsync(News13);
                            ARIS.SpeakAsync(News14);
                            ARIS.SpeakAsync(News15);
                            ARIS.SpeakAsync(News16);
                            ARIS.SpeakAsync(News17);
                            ARIS.SpeakAsync(News18);
                            ARIS.SpeakAsync(News19);
                            ARIS.SpeakAsync(News20);
                            ARIS.SpeakAsync(" End of The News Report");
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;


                    //General Information
                    case "general information":
                        ARIS.SpeakAsync("What do you want to know sir");
                        TextBox1.Visible = true;
                        TextBox1.Text = "";
                        TextBox1.Select();
                        break;
                    case "search information":
                        if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() == true)
                        {
                            ARIS.SpeakAsync("Searching for Information");
                            get_information();
                            ARIS.SpeakAsync(result);
                        }
                        else
                        {
                            ARIS.SpeakAsync("there is no internet connection , please connect to the internet");
                        }
                        break;



                    //Calculations
                    case "mathematics":
                        ARIS.SpeakAsync("enter first and second number");
                        TextBox2.Visible = true;
                        TextBox3.Visible = true;
                        TextBox2.Clear();
                        TextBox3.Clear();
                        break;
                    case "summation":
                        double sum;
                        sum = double.Parse(TextBox2.Text) + double.Parse(TextBox3.Text);
                        ARIS.SpeakAsync("The result is " + sum);
                        break;
                    case "multiplication":
                        double mul;
                        mul = double.Parse(TextBox2.Text) * double.Parse(TextBox3.Text);
                        ARIS.SpeakAsync("The result is " + mul);
                        break;
                    case "subtraction":
                        double sub;
                        sub = double.Parse(TextBox2.Text) - double.Parse(TextBox3.Text);
                        ARIS.SpeakAsync("The result is " + sub);
                        break;
                    case "division":
                        double div;
                        div = double.Parse(TextBox2.Text) / double.Parse(TextBox3.Text);
                        ARIS.SpeakAsync("The result is " + div);
                        break;
                    case "cancel operation":
                        TextBox2.Visible = false;
                        TextBox3.Visible = false;
                        break;


                    // Media Voice Control

                    case "play":
                        SendKeys.Send(" ");
                        SendKeys.Send("^+p");
                        break;

                    case "pause":
                        SendKeys.Send(" ");
                        SendKeys.Send("^+p");
                        break;

                    case "full screen":
                    case "normal screen":
                        SendKeys.Send("f");
                        SendKeys.Send("%{RETURN}");
                        break;

                    case "stop":
                        SendKeys.Send("s");
                        SendKeys.Send(".");
                        SendKeys.Send("^+s");
                        SendKeys.Send("^.");
                        break;

                    case "next":
                        SendKeys.Send("n");
                        SendKeys.Send("{PGDN}");
                        SendKeys.Send("^+n");
                        SendKeys.Send("^{RIGHT}");
                        break;

                    case "previous":
                        SendKeys.Send("p");
                        SendKeys.Send("{PGUP}");
                        SendKeys.Send("^+t");
                        SendKeys.Send("^{LEFT}");
                        break;

                    case "faster":
                        SendKeys.Send("]");

                        break;

                    case "slower":
                        SendKeys.Send("[");

                        break;

                    case "normal speed":
                        SendKeys.Send("=");

                        break;

                    case "volume up":
                        SendKeys.Send("^{UP}");
                        SendKeys.Send("{UP}");
                        SendKeys.Send("^+u");
                        SendKeys.Send("^{UP}");
                        break;

                    case "volume down":
                        SendKeys.Send("^{DOWN}");
                        SendKeys.Send("{DOWN}");
                        SendKeys.Send("^+d");
                        SendKeys.Send("^{DOWN}");
                        break;


                    // Home Automation

                    case "connect to home":
                        serialPort1.Open();
                        ARIS.SpeakAsync("connected");
                        serialPort1.Write("e");
                        ARIS.SpeakAsync("system is now online");
                        timer5.Start();
                        ARIS.SpeakAsync("counting started");
                        break;

                    case "open light one":
                    case "light one":
                        serialPort1.Write("a");
                        ARIS.SpeakAsync("opened");
                        break;

                    case "open light two":
                    case "light two":
                        serialPort1.Write("b");
                        ARIS.SpeakAsync("opened");
                        break;

                    case "open light three":
                    case "light three":
                        serialPort1.Write("c");
                        ARIS.SpeakAsync("opened");
                        break;

                    case "open fan":
                        serialPort1.Write("d");
                        ARIS.SpeakAsync("opened");
                        break;


                    case "close light one":
                        serialPort1.Write("f");
                        ARIS.SpeakAsync("closed");
                        break;

                    case "close light two":
                        serialPort1.Write("g");
                        ARIS.SpeakAsync("closed");
                        break;

                    case "close light three":
                        serialPort1.Write("h");
                        ARIS.SpeakAsync("closed");
                        break;

                    case "close fan":
                        serialPort1.Write("i");
                        ARIS.SpeakAsync("closed");
                        break;


                    case "close all devices":
                        serialPort1.Write("f");
                        serialPort1.Write("g");
                        serialPort1.Write("h");
                        serialPort1.Write("i");
                        ARIS.SpeakAsync("all closed");
                        break;

                    case "open all devices":
                        serialPort1.Write("a");
                        serialPort1.Write("b");
                        serialPort1.Write("c");
                        serialPort1.Write("d");
                        ARIS.SpeakAsync("all opened");
                        break;

                    case "disconnect":
                        serialPort1.Write("j");
                        serialPort1.Close();
                        ARIS.SpeakAsync("disconnected");
                        timer5.Stop();
                        ARIS.SpeakAsync("the system have been online for " + T + " Seconds");
                        T = 0;
                        break;

                    case "close all devices after ten seconds":
                        ARIS.SpeakAsync("all devices will be closed in ten seconds");
                        timer4.Start();
                        break;

                    case "open all devices for ten seconds":
                        ARIS.SpeakAsync("all devices will be open for ten seconds");
                        serialPort1.Write("a");
                        serialPort1.Write("b");
                        serialPort1.Write("c");
                        serialPort1.Write("d");
                        timer6.Start();
                        break;

                }
            }
            catch
            {
                MessageBox.Show("There is a Command Error , Please Try Again");
            }
        }
        private void GetWeather()
        {
            try
            {
                string query = String.Format("http://weather.yahooapis.com/forecastrss?w=2345233");
                XmlDocument wData = new XmlDocument();
                wData.Load(query);

                XmlNamespaceManager man = new XmlNamespaceManager(wData.NameTable);
                man.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

                XmlNode channel = wData.SelectSingleNode("rss").SelectSingleNode("channel");
                XmlNodeList nodes = wData.SelectNodes("/rss/channel/item/yweather:forecast", man);

                Temperature = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", man).Attributes["temp"].Value;

                Condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", man).Attributes["text"].Value;

                Humidity = channel.SelectSingleNode("yweather:atmosphere", man).Attributes["humidity"].Value;

                WindSpeed = channel.SelectSingleNode("yweather:wind", man).Attributes["speed"].Value;

                Town = channel.SelectSingleNode("yweather:location", man).Attributes["city"].Value;

                TFCond = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["text"].Value;

                TFHigh = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["high"].Value;

                TFLow = channel.SelectSingleNode("item").SelectSingleNode("yweather:forecast", man).Attributes["low"].Value;
            }
            catch
            {
                MessageBox.Show("There is an error in Getting Weather");
            }
        }

        private void Timer1_Tick(object sender, EventArgs e)
        {
            try
            { 
            Timer = Timer - 1;
            label4.Text = Timer.ToString();

            if (Timer == 0)
            {
                Timer1.Enabled = false;
                Timer = 20;
                label4.Text = "20";

                switch (QEvent)
                {
                    case "shutdown":
                        System.Diagnostics.Process.Start("shutdown", "-s");
                        break;
                    case "logoff":
                        System.Diagnostics.Process.Start("shutdown", "-l");
                        break;
                    case "restart":
                        System.Diagnostics.Process.Start("shutdown", "-r");
                        break;
                }
            }
            }
            catch
            {
                MessageBox.Show("There is an error in Shutdown , Logoff , Restart");
            }
        }

        private void recognizer_SpeechDetected(Object sender, SpeechDetectedEventArgs e)
        {
            Label1.ForeColor = Color.Blue;
            Label1.BackColor = Color.Transparent;
            Label1.Text = "I Detected That you Said something";
        }
        private void recognizer_SpeechRejected(Object sender, SpeechRecognitionRejectedEventArgs e)
        {
            Label1.ForeColor = Color.Red;
            Label1.BackColor = Color.Transparent;
            Label1.Text = "I Can't Recognize What You Said";
        }

        private void Timer2_Tick(object sender, EventArgs e)
        {
            try
            { 
            ats1 = ats1 - 1;
            Label2.Text = ats1.ToString();
            if (ats1 == 0)
            {
                Timer2.Stop();
                Timer3.Start();
                ARIS.SpeakAsync("Are You Awake Sir ?");
                Label3.Visible = true;
            }
            }
            catch
            {
                MessageBox.Show("There is an error in ATS Mode");
            }
        }

        private void Timer3_Tick(object sender, EventArgs e)
        {
            try
            {
            ats2 = ats2 - 1;
            label4.Text = ats2.ToString();

            if (ats2 == 0)
            {
                System.Diagnostics.Process.Start("shutdown", "-s");
            }
            }
            catch
            {
                MessageBox.Show("There is an error in ATS Mode");
            }
        }
        private void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            System.IO.Stream strm;
            strm = OpenFileDialog1.OpenFile();
            System.Diagnostics.Process.Start(OpenFileDialog1.FileName.ToString());
        }
        private void get_local_news()
        {
            try
            {
                string query = String.Format("http://www.arabnews.com/cat/2/rss.xml");
                XmlDocument nData = new XmlDocument();
                nData.Load(query);

                XmlNamespaceManager man = new XmlNamespaceManager(nData.NameTable);
                man.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");

                XmlNode channel = nData.SelectSingleNode("rss").SelectSingleNode("channel");
                XmlNodeList nodes = nData.SelectNodes("/rss/channel/item/title", man);

                News1 = channel.SelectSingleNode("item").SelectSingleNode("title").InnerText.ToString();
                News2 = channel.SelectSingleNode("item").NextSibling.SelectSingleNode("title").InnerText.ToString();
                News3 = channel.SelectSingleNode("item").NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News4 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News5 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News6 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News7 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News8 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News9 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
                News10 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            }
            catch
            {
                MessageBox.Show("There is an Error in Getting News");
            }
        }
        private void get_world_news()
        {
            try
            { 
            string query = String.Format("http://news.yahoo.com/rss/world");
            XmlDocument nData = new XmlDocument();
            nData.Load(query);

            XmlNamespaceManager man = new XmlNamespaceManager(nData.NameTable);
            man.AddNamespace("media", "http://search.yahoo.com/mrss/");

            XmlNode channel = nData.SelectSingleNode("rss").SelectSingleNode("channel");
            XmlNodeList nodes = nData.SelectNodes("/rss/channel/item/title", man);

            News11 = channel.SelectSingleNode("item").SelectSingleNode("title").InnerText.ToString();
            News12 = channel.SelectSingleNode("item").NextSibling.SelectSingleNode("title").InnerText.ToString();
            News13 = channel.SelectSingleNode("item").NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News14 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News15 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News16 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News17 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News18 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News19 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            News20 = channel.SelectSingleNode("item").NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.SelectSingleNode("title").InnerText.ToString();
            }
            catch
            {
                MessageBox.Show("There is an Error in Getting News");
            }
        }
        private void get_information()
        {
            try
            { 
            string query = String.Format("http://api.wolframalpha.com/v2/query?input=" + TextBox1.Text + "&appid=35V96L-U6JHV2H8HK");
            XmlDocument iData = new XmlDocument();
            iData.Load(query);

            XmlNode queryresult = iData.SelectSingleNode("queryresult").SelectSingleNode("pod").NextSibling;

            result = queryresult.SelectSingleNode("subpod").SelectSingleNode("plaintext").InnerText.ToString();
            }
            catch
            {
                MessageBox.Show("There is an Error in Getting General Information");
            }
        }

        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg,
            IntPtr wParam, IntPtr lParam);

        void NetworkChange_NetworkAvailabilityChanged(object sender, NetworkAvailabilityEventArgs e)
        {
            if (e.IsAvailable)
            {
                ARIS.SpeakAsync("Network is Available");
            }
            else
            {
                ARIS.SpeakAsync("Network is not available");
            }
        }
        public enum DeviceEvent : int
        {
            Arrival = 0x8000,                  //DBT_DEVICEARRIVAL
            QueryRemove = 0x8001,                   //DBT_DEVICEQUERYREMOVE
            QueryRemoveFailed = 0x8002,      //DBT_DEVICEQUERYREMOVEFAILED
            RemovePending = 0x8003,                   //DBT_DEVICEREMOVEPENDING
            RemoveComplete = 0x8004,             //DBT_DEVICEREMOVECOMPLETE
            Specific = 0x8005,                         //DBT_DEVICEREMOVECOMPLETE
            Custom = 0x8006                               //DBT_CUSTOMEVENT
        }
        protected override void WndProc(ref Message m)
        {
            try
            { 
            base.WndProc(ref m);
            const int WM_DEVICECHANGE = 0x0219;
            DeviceEvent lEvent;

            if (m.Msg == WM_DEVICECHANGE)
            {
                lEvent = (DeviceEvent)m.WParam.ToInt32();

                if (lEvent == DeviceEvent.Arrival)
                    ARIS.SpeakAsync("USB Storage DEVICE DETECTED");
                else if (lEvent == DeviceEvent.RemoveComplete)
                    ARIS.SpeakAsync("USB Storage DEVICE REMOVED");
            }
            }
            catch
            {
                MessageBox.Show("There is an Error in EHL");
            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            light = light - 1;
            if (light == 0)
            {
                light = 10;
                timer4.Stop();
                serialPort1.Write("f");
                serialPort1.Write("g");
                serialPort1.Write("h");
                serialPort1.Write("i");
                ARIS.SpeakAsync("all closed");
            }
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            T = T + 1;
        }

        private void timer6_Tick(object sender, EventArgs e)
        {
              X = X - 1;
            if (X == 0)
            {
                X = 10;
                timer6.Stop();
                serialPort1.Write("f");
                serialPort1.Write("g");
                serialPort1.Write("h");
                serialPort1.Write("i");
                ARIS.SpeakAsync("all closed");
            }
        }
        
    }
}

