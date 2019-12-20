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
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            _Path.Text = Properties.Settings.Default.SavedPath;
        }

        private void Main_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            _Path.Text = files[0];
            Properties.Settings.Default.SavedPath = files[0];
            Properties.Settings.Default.Save();
        }

        private void Main_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void _Browse_Click(object sender, EventArgs e)
        {
            string Dir = "";
            if (_Path.Text.Length > 0)
            {
                Dir = System.IO.Path.GetDirectoryName(_Path.Text);
            }
            _OpenFileDialog1.InitialDirectory = Dir;

            if ((int)_OpenFileDialog1.ShowDialog() == (int)DialogResult.OK)
            {
                _Path.Text = _OpenFileDialog1.FileName;
                Properties.Settings.Default.SavedPath = _OpenFileDialog1.FileName;
                Properties.Settings.Default.Save();
            }
        }

        private void _Gen_Click(object sender, EventArgs e)
        {
            _Gen.Enabled = false;

            if (!System.IO.File.Exists(_Path.Text))
            {
                MessageBox.Show("File does not exist");
                return;
            }

            GenTags.Generate(_Path.Text);

            _Gen.Enabled = true;
        }
    }
}
