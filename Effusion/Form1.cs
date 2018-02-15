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

namespace Effusion
{
    public partial class Form1 : Form
    {
        int var = 0;
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        public Form1()
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
            commands.Add(new string[] { "hello", "sorry" ,"open" ,"close", " history",
               "back", "previous", "find" , "new tab" , "next","googal"});
            GrammarBuilder grammar = new GrammarBuilder();
            grammar.Append(commands);
            Grammar g = new Grammar(grammar);

            recEngine.LoadGrammarAsync(g);
            recEngine.SetInputToDefaultAudioDevice();


            recEngine.SpeechRecognized += RecEngine_SpeechRecognized; ;

        }

        private void RecEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text)
            {
                case "googal":
                    addTabToolStripMenuItem_Click(sender, e);
                    toolStripComboBox1.Text = "https://www.google.com";

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
            web.DocumentCompleted += Web_DocumentCompleted;
            tabControl1.TabPages.Add("New tab");
            tabControl1.SelectTab(i);
            tabControl1.SelectedTab.Controls.Add(web);
            i += 1;
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
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Space))
            {
                if (var == 0)
                {
                    var = 1;
                    recEngine.RecognizeAsync(RecognizeMode.Multiple);
                    //button2.Enabled = true;
                    //button1.Enabled = false;
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
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
