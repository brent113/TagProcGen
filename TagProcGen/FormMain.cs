using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TagProcGen
{
    /// <summary>Form Main</summary>
    public partial class FormMain : Form
    {
        /// <summary>Form Main Constructor</summary>
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            Path.Text = Properties.Settings.Default.SavedPath;
        }

        private void FormMain_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            Path.Text = files[0];
            Properties.Settings.Default.SavedPath = files[0];
            Properties.Settings.Default.Save();
        }

        private void FormMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void Browse_Click(object sender, EventArgs e)
        {
            string Dir = "";
            if (Path.Text.Length > 0)
            {
                Dir = System.IO.Path.GetDirectoryName(Path.Text);
            }
            _OpenFileDialog1.InitialDirectory = Dir;

            if ((int)_OpenFileDialog1.ShowDialog() == (int)DialogResult.OK)
            {
                Path.Text = _OpenFileDialog1.FileName;
                Properties.Settings.Default.SavedPath = _OpenFileDialog1.FileName;
                Properties.Settings.Default.Save();
            }
        }

        private void Gen_Click(object sender, EventArgs e)
        {
            Gen.Enabled = false;

            if (!System.IO.File.Exists(Path.Text))
            {
                MessageBox.Show("File does not exist");
                return;
            }

            GenTags.Generate(Path.Text, new MessageBoxNotifier());

            Gen.Enabled = true;
        }
    }

    /// <summary>MessageBox Notifier</summary>
    public class MessageBoxNotifier : INotifier
    {
        /// <summary>Write to log</summary>
        /// <param name="Log">Log text</param>
        /// <param name="Title">Log Title</param>
        /// <param name="Severity">Log Severity</param>
        public void Log(string Log, string Title, LogSeverity Severity)
        {
            MessageBoxIcon icon;
            switch (Severity)
            {
                case LogSeverity.Info:
                    icon = MessageBoxIcon.Information; break;
                case LogSeverity.Warning:
                    icon = MessageBoxIcon.Warning; break;
                case LogSeverity.Error:
                default:
                    icon = MessageBoxIcon.Error; break;
            }
            MessageBox.Show(Log, Title, MessageBoxButtons.OK, icon);
        }
    }
}
