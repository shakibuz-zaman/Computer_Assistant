using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Xml;
using System.IO.Ports;
using System.IO;

namespace Computer_Assistant
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer s = new SpeechSynthesizer();
        String temp;
        String commandstor;
        String condition;
        Choices list = new Choices();
        Boolean wmode = false;
        string openinput = "open facebook" ;
        PowerStatus pw = SystemInformation.PowerStatus;
        SerialPort port = new SerialPort("COM3",9600,Parity.None,8,StopBits.One);
        public Form1()

        {
            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();



            /*list.Add(new String[] { "i am a student","write something on word","are you ok","this is fun","i am fine","hello", "how are you", "what time is it", "what is today", "wake", "sleep", "open bing", "who are you", "open office", "close office",
            "whats the weather like", "whats the temperature","play","resume","open this","pause","go next","last","move forward","light on","light off","open my computer",
            "up","down","right","left","switch to write mode","close write mode","This","is","my","working","project","type","something","on","office","for","me",
            "go to new line","close last term","close last letter","boom","tell me about bettery charge",
            "connect the charger","disconnect the charger","dot"});*/

            list.Add(new String[] { "y", "u", "h", "j", "n", "m", "six", "seven", "start writing", "stop writing" });
            //list.Add(File.ReadAllLines(@"D\projectt\file.txt")
            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeachRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);


            }catch{return;}



            s.SelectVoiceByHints(VoiceGender.Female);
            s.Speak("Hi, I am Gedion");
            //s.Speak("open bing");
            InitializeComponent();

        }
        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='dhaka, bangladesh')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;
                if (input == "temp")
                {
                    return temp;
                }
                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }
        public static void killprog(String ss)
        {
            System.Diagnostics.Process[] procs = null;
            try
            {
                procs = Process.GetProcessesByName(ss);
                Process prog = procs[0];
                if(!prog.HasExited)
                    prog.Kill();
            }
            finally
            {
                if(procs!=null)
                {
                    foreach(Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
        }

        String extractopen(String oi)
        {
            String data="";
            Int16 i;
            Console.WriteLine(oi[1]);
            for (i = 0; i < oi.Length; i++)
            {
                data = data + oi[i];
            }
            return data;
        }
        void say(String h)
        {
            s.Speak(h);
            textBox2.AppendText(h+"\n");
            //textBox1.AppendText(h+"\n");
        }
        /*string loc(string s)
        {
            var map = new Dictionary<string, string>();
            map.Add("open office", "");
            map.Add("open codeblocks", "C:\Program Files (x86)\CodeBlocks\codeblocks.exe");
            map.Add("open chrome","C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
            return map[s];
        }*/
    
        private void rec_SpeachRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;

            Console.WriteLine(r);
            if (wmode)
            {
                Console.WriteLine("Wmode is on\n");
                if (r == "y" || r == "u" || r == "h" || r == "j" || r == "n" || r == "m" || r == "y" || r == "six" || r == "seven")
                {
                    if (r == "six")
                        SendKeys.Send("6");
                    else if (r == "seven")
                        SendKeys.Send("7");
                    else SendKeys.Send(r);
                }

                else if (r == "stop writing")
                {
                    say("writing mode switched off");
                    wmode = false;
                }
                else if (r == "go to new line")
                {
                    SendKeys.Send("{ENTER}");
                }
                else if (r == "dot")
                {
                    SendKeys.Send(".");
                }
                else if (r == "close last term")
                {
                    for (int i = 0; i < commandstor.Length + 1; i++)
                    {
                        SendKeys.Send("{Backspace}");
                    }
                    SendKeys.Send(" ");
                }
                else if (r == "close last letter")
                {
                    SendKeys.Send("{Backspace}");
                    SendKeys.Send("{Backspace}");
                    SendKeys.Send(" ");
                }
                else
                {
                    for (int i = 0; i < r.Length; i++)
                    {
                        // char ch = "{UP}";
                        String s = "" + r[i];
                        SendKeys.Send(s);

                    }
                    SendKeys.Send(" ");
                }
            }

            else
            {
                Console.WriteLine("Wmode is ofF\n");
                if (r == "hello")
                {
                    //SendKeys.Send("B");
                    //SendKeys.Send("C");
                    //SendKeys.Send(" ");
                    say("hi");
                }
                if (r == "tell me about bettery charge")
                {
                    //int bc = pw.BatteryLifeRemaining;
                    float bc = pw.BatteryLifePercent;
                    bc = bc * 100;
                    String sr=bc.ToString();
                    //Console.
                    String s = "you have left " + sr + " percent charge on your bettery";
                    say(s);
                    if (bc<=80)
                    {
                        say("your charge is critically low . charger connected");
                        port.Open();
                        port.WriteLine("A");
                        port.Close();
                    }

                }
                if (r == "how are you")
                    say("I am fine");
                if (r == "what time is it")
                    say(DateTime.Now.ToString("hh:mm tt"));
                if (r == "what is today")
                    say(DateTime.Now.ToString("M/d/yyyy"));
                if (r == "open bing")
                {
                    say("openign bing");
                    Process.Start("http://bing.com");
                    //openinput = extractopen(openinput);
                    //Console.WriteLine(openinput);
                }
                if (r == "who are you")
                    say("i am gedion");
                textBox1.AppendText(r + "\n");
                if (r == "open office")
                {
                    Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Word 2013.lnk");
                }
                if (r == "open my computer")
                {
                    Process.Start(@"C:\Users\ASUS\Desktop\This PC.lnk");
                    SendKeys.Send("{DOWN}");
                    SendKeys.Send("{DOWN}");
                    SendKeys.Send("{DOWN}");
                    SendKeys.Send("{RIGHT}");
                }
                if (r == "close office")
                {
                    //killprog("WINWORD.EXE");
                    killprog("WINWORD");
                }
                if (r == "whats the weather like")
                {
                    say("The sky is " + GetWeather("cond") + ".");
                }
                if (r == "whats the temperature")
                {

                    say("It is " + GetWeather("temp") + "degree farenhite");
                }
                if (r == "pause")
                {
                    SendKeys.Send(" ");
                }
                if (r == "resume")
                {
                    SendKeys.Send(" ");
                }
                if (r == "open this")
                {
                    SendKeys.Send("{ENTER}");
                }
                if (r == "move forward")
                {
                    SendKeys.Send("^{RIGHT}");
                }
                if (r == "move forward")
                {
                    SendKeys.Send("^{LEFT}");
                }
                if (r == "go next")
                {
                    SendKeys.Send("{PGDN}");
                }
                if (r == "right")
                {
                    SendKeys.Send("{RIGHT}");
                }
                if (r == "left")
                {
                    SendKeys.Send("{LEFT}");
                }
                if (r == "up")
                {
                    SendKeys.Send("{UP}");
                }
                if (r == "down")
                {
                    SendKeys.Send("{DOWN}");
                }
                if (r == "last")
                {
                    SendKeys.Send("{PGUP}");
                }
                if (r == "show tabs")
                {
                    SendKeys.Send("%{Tab}");
                }
                if (r == "light on" || r == "connect the charger")
                {
                    say("charger connected");
                    port.Open();
                    port.WriteLine("A");
                    port.Close();
                }
                

                if (r == "light off" || r == "disconnect the charger")
                {
                    say("charger disconnected");
                    port.Open();
                    port.WriteLine("B");
                    port.Close();
                }
                if (r == "start writing")
                {
                    say("writing mode switched on");
                    wmode = true;
                }
            }
            commandstor = r;

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
