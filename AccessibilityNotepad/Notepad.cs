using System.Security.Cryptography;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics.Contracts;

namespace AccessibilityNotepad
{
    public partial class Notepad : Form
    {
        public Notepad()
        {
            InitializeComponent();
        }
        SpeechSynthesizer SpeechSynthesizerObj;
        static bool SpeechToTextOn = false;
        private void Notepad_Load(object sender, EventArgs e)
        {
            SpeechSynthesizerObj = new SpeechSynthesizer();
            NotepadText.Multiline = true;
        }
        /// ///////////////////////////////////////////////
        // SHORTCUTS
        private void FileOpen()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    NotepadText.Text = File.ReadAllText(openFileDialog.FileName);
                }
            }
        }

        private void FileSave()
        {
            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.DefaultExt = ".txt";

            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                File.WriteAllText(saveFileDialog1.FileName, NotepadText.Text);
            }
        }

        private void FileClose()
        {
            Close();
        }

        private void FileUndo()
        {
            NotepadText.Undo();
        }

        private void FileRedo()
        {
            //NotepadText.Redo();
        }

        private void FileCut()
        {
            NotepadText.Cut();
        }
        private void FileCopy()
        {
            NotepadText.Copy();
        }
        private void FilePaste()
        {
            NotepadText.Paste();
        }
        private void FileSelectAll()
        {
            NotepadText.SelectAll();
        }
        private void FileTimeDate()
        {
            NotepadText.Text += System.DateTime.Now.ToString();
        }
        private void FileTextFont()
        {
            //NotepadText.ForeColor = Console.ReadLine();
            FontDialog dlg = new FontDialog();
            //dlg.ShowDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string fontName;
                float fontSize;
                fontName = dlg.Font.Name;
                fontSize = dlg.Font.Size;
                //MessageBox.Show(fontName + "    " + fontSize);
                NotepadText.Font = dlg.Font;
                //menuStrip1.Font = new Font(menuStrip1.Font.FontFamily, 9);
                //menuStrip1.Font = new Font(menuStrip1.Font.Name, "Arial");
                //menuStrip1.Font = new Font(menuStrip1.Font.FontFamily, 9);
                
                menuStrip1.Font = dlg.Font; 
                menuStrip1.Font = new Font(menuStrip1.Font.FontFamily, 9);
            }
        }
        private void FileTextColour()
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                NotepadText.ForeColor = colorDlg.Color;
                menuStrip1.ForeColor = colorDlg.Color;
                statusStrip1.ForeColor = colorDlg.Color; 
            }
        }
        private void FileBackColour()
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                NotepadText.BackColor = colorDlg.Color;
                menuStrip1.BackColor = colorDlg.Color;
                statusStrip1.BackColor = colorDlg.Color;
            }
        }
        private void FileDarkMode()
        {
            if (NotepadText.BackColor != Color.Black)
            {
                menuStrip1.BackColor = Color.Black;
                menuStrip1.ForeColor = Color.White;
                NotepadText.BackColor = Color.Black;
                statusStrip1.BackColor = Color.Black;
                statusStrip1.ForeColor = Color.White;
                NotepadText.ForeColor = Color.White;

                fileToolStripMenuItem.BackColor = Color.Black;
                fileToolStripMenuItem.ForeColor = Color.White;
                newToolStripMenuItem.BackColor = Color.Black;
                newToolStripMenuItem.ForeColor = Color.White;
                openToolStripMenuItem.BackColor = Color.Black;
                openToolStripMenuItem.ForeColor = Color.White;
                saveToolStripMenuItem.BackColor = Color.Black;
                saveToolStripMenuItem.ForeColor = Color.White;
                exitToolStripMenuItem.BackColor = Color.Black;
                exitToolStripMenuItem.ForeColor = Color.White;
                editToolStripMenuItem.BackColor = Color.Black;
                editToolStripMenuItem.ForeColor = Color.White;
                undoToolStripMenuItem.BackColor = Color.Black;
                undoToolStripMenuItem.ForeColor = Color.White;
                redoToolStripMenuItem.BackColor = Color.Black;
                redoToolStripMenuItem.ForeColor = Color.White;
                cutToolStripMenuItem.BackColor = Color.Black;
                cutToolStripMenuItem.ForeColor = Color.White;
                copyToolStripMenuItem.BackColor = Color.Black;
                copyToolStripMenuItem.ForeColor = Color.White;
                pasteToolStripMenuItem.BackColor = Color.Black;
                pasteToolStripMenuItem.ForeColor = Color.White;
                selectAllToolStripMenuItem.BackColor = Color.Black;
                selectAllToolStripMenuItem.ForeColor = Color.White;
                dateTimeToolStripMenuItem.BackColor = Color.Black;
                dateTimeToolStripMenuItem.ForeColor = Color.White;
                settingsToolStripMenuItem.BackColor = Color.Black;
                settingsToolStripMenuItem.ForeColor = Color.White;
                fontToolStripMenuItem.BackColor = Color.Black;
                fontToolStripMenuItem.ForeColor = Color.White;
                colourToolStripMenuItem.BackColor = Color.Black;
                colourToolStripMenuItem.ForeColor = Color.White;
                backgroundColourToolStripMenuItem.BackColor = Color.Black;
                backgroundColourToolStripMenuItem.ForeColor = Color.White;
                textToSpeechToolStripMenuItem.BackColor = Color.Black;
                textToSpeechToolStripMenuItem.ForeColor = Color.White;
                speechToTextToolStripMenuItem.BackColor = Color.Black;
                speechToTextToolStripMenuItem.ForeColor = Color.White;
                wordWrapToolStripMenuItem.BackColor = Color.Black;
                wordWrapToolStripMenuItem.ForeColor = Color.White;
                helpToolStripMenuItem.BackColor = Color.Black;
                helpToolStripMenuItem.ForeColor= Color.White;
                shortcutsToolStripMenuItem.BackColor = Color.Black;
                shortcutsToolStripMenuItem.ForeColor = Color.White;
                aboutNotepadToolStripMenuItem.BackColor = Color.Black;
                aboutNotepadToolStripMenuItem.ForeColor = Color.White;
                darkModeToolStripMenuItem.BackColor = Color.Black;
                darkModeToolStripMenuItem.ForeColor = Color.White;

                menuStrip1.Renderer = new ToolStripProfessionalRenderer(new DarkMode());
            }
            else
            {
                menuStrip1.BackColor = Color.White;
                menuStrip1.ForeColor = Color.Black;
                NotepadText.BackColor = Color.White;
                statusStrip1.BackColor = Color.White;
                statusStrip1.ForeColor = Color.Black;
                NotepadText.ForeColor = Color.Black;

                fileToolStripMenuItem.BackColor = Color.White;
                fileToolStripMenuItem.ForeColor = Color.Black;
                newToolStripMenuItem.BackColor = Color.White;
                newToolStripMenuItem.ForeColor = Color.Black;
                openToolStripMenuItem.BackColor = Color.White;
                openToolStripMenuItem.ForeColor = Color.Black;
                saveToolStripMenuItem.BackColor = Color.White;
                saveToolStripMenuItem.ForeColor = Color.Black;
                exitToolStripMenuItem.BackColor = Color.White;
                exitToolStripMenuItem.ForeColor = Color.Black;
                editToolStripMenuItem.BackColor = Color.White;
                editToolStripMenuItem.ForeColor = Color.Black;
                undoToolStripMenuItem.BackColor = Color.White;
                undoToolStripMenuItem.ForeColor = Color.Black;
                redoToolStripMenuItem.BackColor = Color.White;
                redoToolStripMenuItem.ForeColor = Color.Black;
                cutToolStripMenuItem.BackColor = Color.White;
                cutToolStripMenuItem.ForeColor = Color.Black;
                copyToolStripMenuItem.BackColor = Color.White;
                copyToolStripMenuItem.ForeColor = Color.Black;
                pasteToolStripMenuItem.BackColor = Color.White;
                pasteToolStripMenuItem.ForeColor = Color.Black;
                selectAllToolStripMenuItem.BackColor = Color.White;
                selectAllToolStripMenuItem.ForeColor = Color.Black;
                dateTimeToolStripMenuItem.BackColor = Color.White;
                dateTimeToolStripMenuItem.ForeColor = Color.Black;
                settingsToolStripMenuItem.BackColor = Color.White;
                settingsToolStripMenuItem.ForeColor = Color.Black;
                fontToolStripMenuItem.BackColor = Color.White;
                fontToolStripMenuItem.ForeColor = Color.Black;
                colourToolStripMenuItem.BackColor = Color.White;
                colourToolStripMenuItem.ForeColor = Color.Black;
                backgroundColourToolStripMenuItem.BackColor = Color.White;
                backgroundColourToolStripMenuItem.ForeColor = Color.Black;
                textToSpeechToolStripMenuItem.BackColor = Color.White;
                textToSpeechToolStripMenuItem.ForeColor = Color.Black;
                speechToTextToolStripMenuItem.BackColor = Color.White;
                speechToTextToolStripMenuItem.ForeColor = Color.Black;
                wordWrapToolStripMenuItem.BackColor = Color.White;
                wordWrapToolStripMenuItem.ForeColor = Color.Black;
                helpToolStripMenuItem.BackColor = Color.White;
                helpToolStripMenuItem.ForeColor = Color.Black;
                shortcutsToolStripMenuItem.BackColor = Color.White;
                shortcutsToolStripMenuItem.ForeColor = Color.Black;
                aboutNotepadToolStripMenuItem.BackColor = Color.White;
                aboutNotepadToolStripMenuItem.ForeColor = Color.Black;
                darkModeToolStripMenuItem.BackColor = Color.White;
                darkModeToolStripMenuItem.ForeColor = Color.Black;

                menuStrip1.Renderer = new ToolStripProfessionalRenderer(new LightMode());
            }
        }
        private void FileOpenHelp()
        {
            ShortcutsPage frm = new ShortcutsPage();
            frm.Show();
        }
        private void FileAbout()
        {
            MessageBox.Show("This notepad is epic and you are not");
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
                SpeechSynthesizerObj.SpeakAsync(NotepadText.Text);
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
            NotepadText.Text = NotepadText.Text + e.Result.Text.ToString() + Environment.NewLine;
        }

        //END OF SHORTCUTS 
        //=============================================================================









        //////////////////////////////////////////////////////
        //Menustrip


        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileOpen();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileSave();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileClose();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileUndo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileRedo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileCut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileCopy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FilePaste();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileSelectAll();
        }

        private void dateTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileTimeDate();
        }

        private void fontToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileTextFont();
        }

        private void colourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileTextColour();
        }

        private void textToSpeechToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileTextToSpeech();
        }

        private void speechToTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileSpeechToText();
        }

        private void backgroundColourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileBackColour();
        }

        private void shortcutsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileOpenHelp();
        }

        private void aboutNotepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileAbout();
        }

        private void darkModeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileDarkMode();
        }

        //END MENU STRIP
        //===========================================================================================================





        //SHORTCUT INPUTS
        //============================
        protected override bool ProcessCmdKey(ref Message message, Keys keys)
        {
            switch (keys)
            {
                case Keys.Control | Keys.O:
                    {
                        FileOpen();
                        return true;
                    }

                case Keys.Control | Keys.S:
                    {
                        FileSave();
                        return true;
                    }
                case Keys.Control | Keys.W:
                    {
                        FileClose();
                        return true;
                    }
                case Keys.Control | Keys.U:
                    {
                        FileUndo();
                        return true;
                    }
                case Keys.Control | Keys.R:
                    {
                        FileRedo();
                        return true;
                    }
                case Keys.Control | Keys.X:
                    {
                        FileCut();
                        return true;
                    }
                case Keys.Control | Keys.C:
                    {
                        FileCopy();
                        return true;
                    }
                case Keys.Control | Keys.P:
                    {
                        FilePaste();
                        return true;
                    }
                case Keys.Control | Keys.A:
                    {
                        FileSelectAll();
                        return true;
                    }
                case Keys.Control | Keys.T:
                    {
                        FileTimeDate();
                        return true;
                    }
                case Keys.Control | Keys.F1:
                    {
                        FileTextFont();
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
                case Keys.Control | Keys.H:
                    {
                        FileOpenHelp();
                        return true;
                    }
            }
            return base.ProcessCmdKey(ref message, keys);
        }

        private void WordCounter()
        {
            MatchCollection collection = Regex.Matches(NotepadText.Text, @"[\S]+");
            //return collection.Count;
            toolStripStatusLabel1.Text = $"WordCount: { Convert.ToString(collection.Count)}";
        }
        private void CharCounter()
        {
            int charCounter = NotepadText.Text.Length;
            toolStripStatusLabel2.Text = $"CharCount: {Convert.ToString(charCounter)}";
        }

        private void NotepadText_TextChanged(object sender, EventArgs e)
        {
            WordCounter();
            CharCounter();
        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (NotepadText.WordWrap == false) { NotepadText.WordWrap = true; }
            else if (NotepadText.WordWrap == true) { NotepadText.WordWrap = false; }
        }

        public class DarkMode : ProfessionalColorTable
        {
            public DarkMode()
            {
                base.UseSystemColors = false;
            }
            Color color = Color.FromKnownColor(KnownColor.DarkOrange);
            public override Color MenuItemSelected
            {
                get { return color; }
            }

            public override Color MenuItemBorder
            {
                get { return color; }
            }

            public override Color MenuItemSelectedGradientBegin
            {
                get { return color; }
            }
        }
        public class LightMode : ProfessionalColorTable
        {
            public LightMode()
            {
                base.UseSystemColors = true;
            }
        }
    }
}
