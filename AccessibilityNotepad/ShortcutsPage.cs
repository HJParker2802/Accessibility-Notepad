using System.Security.Cryptography;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics.Contracts;

namespace AccessibilityNotepad
{
    public partial class ShortcutsPage : Form
    {
        public ShortcutsPage()
        {
            InitializeComponent();
        }

        SpeechSynthesizer SpeechSynthesizerObj;
        static bool SpeechToTextOn = false;
        private void ShortcutsPage_Load(object sender, EventArgs e)
        {
            SpeechSynthesizerObj = new SpeechSynthesizer();
            textBox1.Multiline = true;
            FileDarkMode();
        }
        private void FileClose()
        {
            Close();
        }

        private void FileCopy()
        {
            textBox1.Copy();
        }
        private void FileSelectAll()
        {
            textBox1.SelectAll();
        }
        private void FileTextColour()
        {
            ColorDialog colorDlg = new ColorDialog();
            colorDlg.AllowFullOpen = false;
            colorDlg.AnyColor = true;
            colorDlg.SolidColorOnly = false;
            colorDlg.Color = Color.Red;

            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.ForeColor = colorDlg.Color;
                lblTitle.ForeColor = colorDlg.Color;
            }
        }
        private void FileBackColour()
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.BackColor = colorDlg.Color;
                BackColor = colorDlg.Color;
            }
        }
        private void FileDarkMode()
        {
            if (Notepad.IsDarkModeOn == true)
            {
                LightModeOn();
            }
            else if (Notepad.IsDarkModeOn == false)
            {
                DarkModeOn();
            }   
        }
        private void DarkModeOn()
        {
            BackColor = Color.Black; 
            textBox1.BackColor = Color.Black;
            textBox1.ForeColor = Color.White;
            lblShortcuts.ForeColor = Color.White;
            Notepad.IsDarkModeOn = false;
        }
        private void LightModeOn()
        {
            BackColor = Color.White;
            textBox1.BackColor = Color.White;
            textBox1.ForeColor = Color.Black;
            lblShortcuts.ForeColor = Color.Black;
            Notepad.IsDarkModeOn = true;
        }
            
                private void FileTextToSpeech()
        {
            //Needs System.Speech to work, does not run on college machines atm 
            if (SpeechSynthesizerObj.State == SynthesizerState.Paused)
            {
                SpeechSynthesizerObj.Resume();
            }
            else if (SpeechSynthesizerObj.State != SynthesizerState.Speaking)
            {
                SpeechSynthesizerObj.SpeakAsync(textBox1.Text);
            }
            else if (SpeechSynthesizerObj.State == SynthesizerState.Speaking)
            {
                SpeechSynthesizerObj.Pause();
            }
        }

        static SpeechSynthesizer SS = new SpeechSynthesizer();
        static PromptBuilder PB = new PromptBuilder();
        static SpeechRecognitionEngine SRE = new SpeechRecognitionEngine();
        static Choices clist = new Choices();

        private void FileSpeechToText()
        {//https://www.youtube.com/watch?v=EgfsGUbhH54
            if (SpeechToTextOn == false) { SpeechToTextOn = true; }
            else if (SpeechToTextOn == true) { SpeechToTextOn = false; }

            if (SpeechToTextOn == true)
            {
                MessageBox.Show("Speech To Text is On");
                //clist.Add(new string[] { "Hello", "Good Morning", "Son of a bitch", "Son of a witch" });
                //Grammar GR = new Grammar(new GrammarBuilder(clist));
                Grammar GR = new DictationGrammar();
                try
                {
                    SRE.RequestRecognizerUpdate();
                    SRE.LoadGrammar(GR);
                    SRE.SpeechRecognized += SRE_SpeechRecognized;
                    SRE.SetInputToDefaultAudioDevice();
                    SRE.RecognizeAsync(RecognizeMode.Multiple);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error with Text To Speech"); }
            }
            else { MessageBox.Show("Speech To Text is Off"); }
            }

        private void SRE_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
        {
            textBox1.Text = textBox1.Text + e.Result.Text.ToString() + Environment.NewLine;
        }

        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.Control | Keys.W:
                    {
                        FileClose();
                        return true;
                    }

                case Keys.Control | Keys.C:
                    {
                        FileCopy();
                        return true;
                    }

                case Keys.Control | Keys.A:
                    {
                        FileSelectAll();
                        return true;
                    }

                case Keys.Control | Keys.F2:
                    {
                        FileTextColour();
                        return true;
                    }
                case Keys.Control | Keys.F3:
                    {
                        FileBackColour();
                        return true;
                    }
                case Keys.Control | Keys.D:
                    {
                        FileDarkMode();
                        return true;
                    }
                case Keys.Control | Keys.F5:
                    {
                        FileTextToSpeech();
                        return true;
                    }
                case Keys.Control | Keys.F6:
                    {
                        FileSpeechToText();
                        return true;
                    }
            }
            return base.ProcessCmdKey(ref message, keys);
        }
    }
}
