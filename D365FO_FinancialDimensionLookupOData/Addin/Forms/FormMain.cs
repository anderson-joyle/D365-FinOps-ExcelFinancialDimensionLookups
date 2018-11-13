using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addin
{
    public partial class FormMain : FormBase
    {
        List<string> modelList;

        public FormMain(List<string> modelList)
        {
            InitializeComponent();

            this.modelList = modelList;
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            foreach (string modelName in this.modelList)
            {
                comboBoxModels.Items.Add(modelName);
            }
        }

        public string getModelName()
        {
            return comboBoxModels.SelectedItem.ToString();
        }

        private void buttonAbout_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog();
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
