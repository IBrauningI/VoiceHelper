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
using System.IO.Ports;
using System.Xml;
using System.Net;
using System.IO;

namespace VoiceHelper
{

   
    public partial class Form1 : Form
    {
        SpeechSynthesizer s = new SpeechSynthesizer();
        Boolean wake = true;
        String temp;
        String condition;


        String name = "Vladyslav";
        Boolean var1 = true;




        Choices list = new Choices();

        public Boolean search = false;
        public Form1()
        {
            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

            list.Add(File.ReadAllLines(@"C:\Users\IBRAU\OneDrive\Робочий стіл\VoiceHelper commands\commands.txt"));



            Grammar gr = new Grammar(new GrammarBuilder(list));

            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeachRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);

            }

            catch { return; }

            s.SelectVoiceByHints(VoiceGender.Female);
            s.Speak("Hello, My name is VoiceHelper,how can I help you?");

            InitializeComponent();
        }

        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='Miami, state')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            try
            {
                wData.Load(query);
            }

            catch
            {
                MessageBox.Show("No internet connection");
                return "No internet";
            }


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

        public void killProg(String s)
        {

            System.Diagnostics.Process[] procs = null;


            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];

                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }
            catch
            {
                say("Notepad is not open.");
            }
            finally
            {
                if (procs != null)
                {
                    foreach (Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
            procs = null;
        }


        public void restart()
        {
            Process.Start(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint 2016.lnk");
        }
        public void say(string h)
        {
            s.Speak(h);
           // wake = false;
            textBox2.AppendText(h + "\r\n");
        }


        // команди
        private void rec_SpeachRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;

            if (search)
            {
                Process.Start(@"https://google.com.ua/#q=" + r);
            }

            //Можна говорити команди,тільки після цієї команди.Щоб забрати в say wake=false; забрати, і Boolean wake = true поставити
            if (r == "Helper")
            {
                wake = true;
            }

            if (r == "Wake")
            {
                wake = true;
                label3.Text = "State: Awake";
            }


            if (r == "Sleep")
            {
                wake = false;
                label3.Text = "State: Sleep mode";
            }

            if (wake == true  && search == false )
            {

                if (r == "Find for me")
                {
                    search = true;
                }

                if (r=="Do I like football")
                {
                    if (var1)
                    {
                        say("Yes," +name+",you love it");
                    }
                    if (!var1)
                    {
                        say("No," + name + "you hate it");
                    }
                }

                if (r == "What is my name")
                {
                    say("Your name is," + name);
                }

                if (r == "Delete")
                {
                    SendKeys.Send("{DEL}");
                }
                if (r == "New slide")
                {
                    SendKeys.Send("^{M}");
                }


                if (r == "Up")
                {
                    SendKeys.Send("{UP}");
                }

                if (r == "Down")
                {
                    SendKeys.Send("{DOWN}");
                }

                if (r == "Left")
                {
                    SendKeys.Send("{LEFT}");
                }

                if (r == "Right")
                {
                    SendKeys.Send("{RIGHT}");
                }

                if (r == "Aproach")
                {
                    SendKeys.Send("^{+}");
                }

                if (r == "Remove")
                {
                    SendKeys.Send("^{-}");
                }
                if (r == "Back")
                {
                    SendKeys.Send("{ESC}");
                }


                if (r == "Enter")
                {
                    SendKeys.Send("{ENTER}");
                }

                if (r == "Presentation")
                {
                    Process.Start(@"D:\CourseWork.pptx");
                }

                if (r == "Start")
                {
                    SendKeys.Send("{F5}");
                }

                if (r == "Last")
                {
                    SendKeys.Send("{P}");
                }

                if (r == "Next")
                {
                    SendKeys.Send("{N}");
                }


                if (r == "Last song")
                {
                    SendKeys.Send("^{LEFT}");
                }
                if (r == "Next song")
                {
                    SendKeys.Send("^{RIGHT}");
                }
                if (r == "Spotify")
                {
                    Process.Start(@"C:\Users\IBRAU\AppData\Roaming\Spotify\Spotify.exe");
                }

                if (r == "Play" || r == "Pause")
                {
                    SendKeys.Send(" ");
                }

                if (r == "Minimaze")
                {
                    this.WindowState = FormWindowState.Minimized;
                }

                if (r == "Unminimaze")
                {
                    this.WindowState = FormWindowState.Normal;
                }

                if (r == "Maximaze")
                {
                    this.WindowState = FormWindowState.Maximized;
                    
                }


                if (r == "Whats the weather like")
                {
                    say("The sky is" + GetWeather("cond" + "."));
                }

                if (r == "Whats the temperature")
                {
                    say("It is" + GetWeather("temp" + "degrees."));
                }
                if (r == "Open notepad")
                {
                    Process.Start(@"C:\Program Files (x86)\Notepad++\notepad++.exe");
                }

                if (r == "Close notepad")
                {
                    killProg("notepad++.exe");
                }
                //Що я кажу
                if (r == "Hello")
                {
                    //відповідь
                    say("Hi");
                }

                if (r == "How are you")
                {
                    say("I'm great, and you ?");
                }

                if (r == "Good")
                {
                    say("Nice,to hear that");
                }


                if (r == "What is today")
                {
                    say(DateTime.Now.ToString("M/d/yyyy"));
                }

                if (r == "What time is it")
                {
                    say(DateTime.Now.ToString("h:mm tt"));
                }

                if (r == "Open google")
                {
                    Process.Start("http://google.com.ua");
                }

                if (r == "Restart" || r == "Update")
                {
                    restart();
                }
            }
            textBox1.AppendText(r + "\r\n");


        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }
    }
}


