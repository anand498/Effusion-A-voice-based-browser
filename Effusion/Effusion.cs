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
using System.Windows;
using System.Xml;
using System.Net;
using System.IO;
using System.Collections;
using System.Web;
using System.Text.RegularExpressions;

namespace Effusion
{
    public partial class Effusion : Form
    {
        int var = 0;
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        public Effusion()
        {
            InitializeComponent();
        }
        WebBrowser web = new WebBrowser();
        int i = 0;


        private void Form1_Load(object sender, EventArgs e)
        {
            web = new WebBrowser();
            web.ScriptErrorsSuppressed = true;
            web.Dock = DockStyle.Fill;
            web.Visible = true;
            web.DocumentCompleted += Web_DocumentCompleted;
            tabControl1.TabPages.Add("New tab");
            tabControl1.SelectTab(i);
            tabControl1.SelectedTab.Controls.Add(web);
            i += 1;
            Choices commands = new Choices();
            commands.Add(new string[] { "mail","open" ,"close", " history","V I T",
               "back", "previous", "find" , "new tab" , "next","face book","quora", "refresh", "duck duck go","yaahoo","exit","google","google","youtube","moodle","one"});
            GrammarBuilder grammar = new GrammarBuilder();
            grammar.Append(commands);
            Grammar g = new Grammar(grammar);

            recEngine.LoadGrammarAsync(g);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += RecEngine_SpeechRecognized; ;

        }

        private void RecEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string result = e.Result.Text;
            switch (result)
            {
                case "exit":
                    this.Close();
                    break;
                /*case "duck duck go":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.duckduckgo.com";
                    goToolStripMenuItem_Click(sender, e);
                    break;*/
                case "reload":
                    refresh();
                    break;
                case "refresh":
                    refresh();
                    break;
                case "face book":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.facebook.com";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "quora":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.quora.com";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "open":
                    addTabToolStripMenuItem_Click(sender, e);
                    break;
                case "new tab":
                    addTabToolStripMenuItem_Click(sender, e);
                    break;
                case "close":
                    closeTabToolStripMenuItem_Click(sender, e);
                    break;
                case "back":
                    toolStripMenuItem1_Click(sender, e);
                    break;
                case "previous":
                    toolStripMenuItem1_Click(sender, e);
                    break;
                case "close tab":
                    closeTabToolStripMenuItem_Click(sender, e);
                    break;
                case "next":
                    toolStripMenuItem2_Click(sender, e);
                    break;
                case "history":
                    historyToolStripMenuItem_Click(sender,e);
                    break;
                case "V I T":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://academicscc.vit.ac.in";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "google":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.google.com";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "moodle":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "http://moodlecc.vit.ac.in";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "youtube":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.youtube.com";
                    goToolStripMenuItem_Click(sender, e);
                    break;
                case "mail":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://google.co.in";
                    goToolStripMenuItem_Click(sender, e);
                    break;
            }
        }

        void Web_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            tabControl1.SelectedTab.Text = ((WebBrowser)tabControl1.SelectedTab.Controls[0]).DocumentTitle;

        }

        private void goToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((WebBrowser)tabControl1.SelectedTab.Controls[0]).Navigate(toolStripComboBox1.Text);
            if(!toolStripComboBox1.Items.Contains(toolStripComboBox1.Text))
            {
                toolStripComboBox1.Items.Add(toolStripComboBox1.Text);
            }
            addHistory();

        }

        private void refresh()
        {
            ((WebBrowser)tabControl1.SelectedTab.Controls[0]).Refresh();

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ((WebBrowser)tabControl1.SelectedTab.Controls[0]).GoBack();


        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ((WebBrowser)tabControl1.SelectedTab.Controls[0]).GoForward();

        }

        private void addTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
            web = new WebBrowser();
            web.ScriptErrorsSuppressed = true;
            web.Dock = DockStyle.Fill;
            web.Visible = true;
            //web.ProgressChanged += new WebBrowserProgressChangedEventHandler(Effusion_ProgressChanged);
            web.DocumentCompleted += Web_DocumentCompleted;
            tabControl1.TabPages.Add("New tab");
            tabControl1.SelectTab(i);
            tabControl1.SelectedTab.Controls.Add(web);
            i += 1;
        }

        //private void Effusion_ProgressChanged(object sender, WebBrowserProgressChangedEventArgs e)
        //{
        //    if (e.CurrentProgress < e.MaximumProgress)
        //        toolStripProgressBar1.Value = (int)e.CurrentProgress;
        //    else toolStripProgressBar1.Value = toolStripProgressBar1.Minimum;
        //}

        private void addHistory()
        {
            XmlDocument myXml = new XmlDocument();
            if (!File.Exists("history.xml")){
                XmlElement root = myXml.CreateElement("history");
                myXml.AppendChild(root);
                XmlElement el = myXml.CreateElement("item");
                el.SetAttribute("url", toolStripComboBox1.Text);
                el.SetAttribute("date", DateTime.Now.ToString());
                root.AppendChild(el);
                myXml.Save("history.xml");
            }
            else
            {
                myXml.Load("history.xml");
                XmlElement el = myXml.CreateElement("item");
                el.SetAttribute("url", toolStripComboBox1.Text);
                el.SetAttribute("date", DateTime.Now.ToString());
                myXml.DocumentElement.AppendChild(el);
                myXml.Save("history.xml");
            }
        }
         
        private void closeTabToolStripMenuItem_Click(object sender, EventArgs e)
        {
                if(tabControl1.TabPages.Count -1> 0)
            {
                tabControl1.TabPages.RemoveAt(tabControl1.SelectedIndex);
                tabControl1.SelectTab(tabControl1.TabPages.Count - 1);
                i -= 1;
            }
        }
        /*
        private void gotonum(object sender, EventArgs e,int num)
        {
            tabControl1.SelectTab(num);
        }*/

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Space))
            {
                if (var == 0)
                {
                    var = 1;
                    recEngine.RecognizeAsync(RecognizeMode.Multiple);
                    richTextBox1.Text = "Listening\n";
                    return true;
                }
                if (var == 1)
                {
                    var = 0;
                    recEngine.RecognizeAsyncStop();
                    richTextBox1.Text = "Stopped Listening \n";
                }
            }
            if (keyData == Keys.Enter)
            {
                if (!toolStripComboBox1.Items.Contains(toolStripComboBox1.Text))
                {
                    string source = toolStripComboBox1.Text;
                    string regexp = @"^([a-zA-Z].)[a-zA-Z0-9\-\.]+\.(com|edu|gov|mil|net|org|biz|info|name|museum|us|ca|uk)(\:[0-9]+)*(/($|[a-zA-Z0-9\.\,\;\?\'\\\+&amp;%\$#\=~_\-]+))*$";
                    Regex reg = new Regex(regexp, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    if (reg.IsMatch(source))
                    {
                        toolStripComboBox1.Items.Add(toolStripComboBox1.Text);
                        ((WebBrowser)tabControl1.SelectedTab.Controls[0]).Navigate(toolStripComboBox1.Text);
                        if (!toolStripComboBox1.Items.Contains(toolStripComboBox1.Text))
                        {
                            toolStripComboBox1.Items.Add(toolStripComboBox1.Text);
                        }
                    }
                    else
                    {
                        web = new WebBrowser();
                        string query = "";
                        query = "https://www.google.com/search?q=" + toolStripComboBox1.Text;
                        ((WebBrowser)tabControl1.SelectedTab.Controls[0]).Navigate(query);
                    }
                }
            }
                    addHistory();
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void historyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addTabToolStripMenuItem_Click(sender, e);
            toolStripComboBox1.Text = "file:///C:/Users/Anand/source/repos/Effusion/Effusion/bin/Debug/history.xml";
            goToolStripMenuItem_Click(sender, e);
        }
    }
}
