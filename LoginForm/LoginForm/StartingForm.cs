using System;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class StartingForm : Form
    {
        public StartingForm()
        {
            
            InitializeComponent();
            this.TransparencyKey = BackColor;
            timer1.Start();
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            LoginForm f = new LoginForm();
            f.Show();
            this.Hide();
        }

    }
}
