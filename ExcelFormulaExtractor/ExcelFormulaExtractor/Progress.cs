using System;
using System.Windows.Forms;

namespace ExcelFormulaExtractor
{
    public partial class Progress : Form
    {
        readonly string _label;
        readonly int _total;
        int _completed = 0;

        public Progress(string label, int total)
        {
            _label = label;
            _total = total;
            InitializeComponent();
            updateLabel();
            this.CenterToScreen();
        }

        public void increment()
        {
            // if this method is called from any thread other than
            // the GUI thread, call the method on the correct thread
            if (this.InvokeRequired)
            {
                BeginInvoke(new Action(increment));
                return;
            }

            _completed++;
            updateLabel();
            this.Refresh();
        }

        private void updateLabel()
        {
            this.label1.Text = _label + ": " + _completed + " of " + _total;
        }
    }
}