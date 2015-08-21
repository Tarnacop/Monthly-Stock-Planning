using System;
using System.Diagnostics;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace WindowsFormsApplication1
{
    public partial class LoginForm : Form
    {
        // Constructor
        public LoginForm()
        {
            InitializeComponent();
            groupBox1.Controls.Add(textBox1);
            groupBox1.Controls.Add(textBox2);
        } // END of LoginForm CONSTRUCTOR

        // Creates a linkLabel for contact
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string url;
            if (e.Link.LinkData != null)
                url = e.Link.LinkData.ToString();
            else
                url = linkLabel1.Text.Substring(e.Link.Start, e.Link.Length);

            var si = new ProcessStartInfo("http://google.ro");
            Process.Start(si);
            linkLabel1.LinkVisited = true;
        }// END OF linkLabel1_LinkClicked EVENT

        MySqlConnection connection; // Used to connect to database
        MySqlCommand command; // Used to launch a command to database

        // Event handler for button click
        private void button1_Click(object sender, EventArgs e)
        {
            
            // Check if the text boxes have any text
            if(string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("You have to complete the username field!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                textBox1.Focus();
            }
            else if(string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("You have to complete the password field!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                textBox2.Focus();
            }

            string connectionString = "Server = 127.0.0.1; Port = 3306; Database = db_admin; Uid = root; Pwd=; "; // Details for connection
            //string connectionString = "Server = 192.168.1.113; Port = 3306; Database = db_A030; Uid = user_A030; Pwd = pass_A030";
            string querry = "SELECT userType FROM users WHERE userName = '" + textBox1.Text + "' AND userPass = '" + textBox2.Text +"'"; // The querry for the database
            connection = new MySqlConnection(connectionString); // Create a connection with the database
            command = new MySqlCommand(querry, connection); // Create a command for the database

            connection.Open(); // Connect to the database 
            string userType = command.ExecuteScalar() as string; // Execute the command and return the result as a string
            connection.Close(); // Close the connection to the database 

            // Check if any results have been returned. If there is no result, return an error to the user
            if (string.IsNullOrEmpty(userType))
            {
                MessageBox.Show("Wrong username or password!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Text = "";
                textBox2.Text = "";
            }
            // If there is a result, check if the user is an administrator or an user based on what the database returned and
            // open the form for the administrator or the user
            else if (userType.Equals("A"))
            {
                AdminForm f = new AdminForm();
                f.Show();
                this.Close();
            }
            else if (userType.Equals("U"))
            {
                UserForm f = new UserForm();
                f.Show();
                this.Close();
            }
            
            /*if (textBox1.Text == "")
            {
                MessageBox.Show("You have to complete the username field!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                textBox1.Focus();
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("You have to complete the password field!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                textBox2.Focus();
            }
            else if (textBox1.Text == "admin")
            {
                AdminForm f = new AdminForm();
                f.Show();
                this.Close();
            }
            else
            {
                UserForm f = new UserForm();
                f.Show();
                this.Close();
            }*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
