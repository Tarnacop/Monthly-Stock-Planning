using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;

namespace WindowsFormsApplication1
{
    public partial class UserForm : Form
    {
        public UserForm()
        {
            InitializeComponent();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
        }

        MySqlConnection connection;
        MySqlCommandBuilder commandBuilder;
        MySqlDataAdapter dataAdapter;
        MySqlCommand command;
        DataTable dt;
        string connectionString = "Server = 127.0.0.1; Port = 3306; Database = db_user; Uid = root; Pwd=; ";
        string querry = "";
        string selectedItem;
        private void UserForm_Load(object sender, EventArgs e)
        {
            connection = new MySqlConnection(connectionString);
            connection.Open();
            DataTable schema = connection.GetSchema("Tables");
            foreach (DataRow row in schema.Rows)
            {
                listBox1.Items.Add(row[2].ToString());
            }
            dt.Columns[0].ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                button5.Enabled = true;
                selectedItem = listBox1.SelectedItem.ToString();
                querry = "SELECT * FROM " + listBox1.SelectedItem.ToString();
                connection = new MySqlConnection(connectionString);
                dataAdapter = new MySqlDataAdapter(querry, connectionString);
                dt = new DataTable();
                dataAdapter.Fill(dt);
                if (dataGridView1.DataSource == null)
                {

                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();
                }
                dataGridView1.DataSource = dt;
            }
            catch (NullReferenceException ex)
            {
                MessageBox.Show("Please select one of the databases in the list.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Visible == false && button4.Visible == false)
            {
                button2.Text = "Querry on";
                button2.ForeColor = Color.Green;
                textBox1.Visible = true;
                button4.Visible = true;
            }
            else
            {
                button2.Text = "Querry off";
                button2.ForeColor = Color.Red;
                textBox1.Visible = false;
                button4.Visible = false;
                textBox1.Text = "";
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                commandBuilder = new MySqlCommandBuilder(dataAdapter);
                dataAdapter.Update(dt);
                dt = new DataTable();
                dataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
                MessageBox.Show("Updated succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error while triyng to update!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoginForm f = new LoginForm();
            f.Show();
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button5.Enabled = true;
            if (string.IsNullOrEmpty(textBox1.Text))
                MessageBox.Show("Please enter a querry!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                querry = textBox1.Text;
                if (querry.Contains("SELECT"))
                {
                    try
                    {
                        dataAdapter = new MySqlDataAdapter(querry, connectionString);
                        dt = new DataTable();
                        dataAdapter.Fill(dt);
                        if (dataGridView1.DataSource == null)
                        {

                            dataGridView1.Rows.Clear();
                            dataGridView1.Columns.Clear();
                        }
                        dataGridView1.DataSource = dt;
                        MessageBox.Show("Command executed succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        textBox1.Text = "";
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("There was an error while triyng to querry the database!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    try
                    {
                        command = new MySqlCommand(querry, connection);
                        command.ExecuteNonQuery();
                        MessageBox.Show("Command executed succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show("There was an error while triyng to querry the database!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

            }

        }

        private void contactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm f = new AboutForm();
            f.Show();
        }

        private void supportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var si = new ProcessStartInfo("http://google.ro");
            Process.Start(si);
        }

        private void exportAsXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFileDialog1.Title = "Save the Excel file!";
            saveFileDialog1.FileName = selectedItem;
            saveFileDialog1.Filter = "Excel Document(2003)|*.xls|Excel Document(2007)|*.xlsx";
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {

                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("The DataGridView is empty. Please fill it!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        Workbook excelWorkbook = new Workbook();
                        Worksheet worksheet = new Worksheet("DataGridView");
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            worksheet.Cells[0, i] = new Cell(dataGridView1.Columns[i].HeaderText);
                        }
                        for (int i = 1; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {
                                worksheet.Cells[i, j] = new Cell(dataGridView1.Rows[i - 1].Cells[j].Value.ToString());
                            }
                        }
                        //MessageBox.Show(dataGridView1.Rows[0].Cells[0].Value.ToString());
                        excelWorkbook.Worksheets.Add(worksheet);
                        excelWorkbook.Save(saveFileDialog1.FileName.ToString());
                        MessageBox.Show("File exported succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Can't export the file!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void importXLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Title = "Open an Excel file!";
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "*.xls|*.xls|*.xlsx|*.xlsx";

            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {

                try
                {
                    button5.Enabled = false;
                    Workbook excelWorkbook = Workbook.Load(openFileDialog1.FileName.ToString());
                    Worksheet worksheet = excelWorkbook.Worksheets[0];

                    dataGridView1.DataSource = null;
                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();

                    Row row = worksheet.Cells.GetRow(0);
                    for (int i = row.FirstColIndex; i <= row.LastColIndex; i++)
                    {
                        dataGridView1.ColumnCount = i + 1;
                        dataGridView1.Columns[i].HeaderText = row.GetCell(i).Value.ToString();
                    }

                    for (int i = worksheet.Cells.FirstRowIndex + 1; i <= worksheet.Cells.LastRowIndex; i++)
                    {
                        row = worksheet.Cells.GetRow(i);
                        dataGridView1.Rows.Add();
                        for (int j = row.FirstColIndex; j <= row.LastColIndex; j++)
                        {
                            dataGridView1.Rows[i - 1].Cells[j].Value = row.GetCell(j).Value;
                        }
                    }

                    MessageBox.Show("File imported succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Can't import the file!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void exportAsPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            saveFileDialog1.Title = selectedItem;
            saveFileDialog1.FileName = listBox1.SelectedItem.ToString();
            saveFileDialog1.Filter = "*.pdf|*.pdf";
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {


                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("The DataGridView is empty. Please fill it!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    try
                    {
                        Document doc = new Document(iTextSharp.text.PageSize.A4);
                        PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(saveFileDialog1.FileName.ToString(), FileMode.Create));
                        doc.Open(); // Open document to be written

                        iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(System.Reflection.Assembly.GetExecutingAssembly().Location + "\\..\\..\\..\\Resources\\logoprogram.png");

                        logo.SetAbsolutePosition(doc.PageSize.Width / 2 - 200f, doc.PageSize.Height - 250f);
                        doc.Add(logo);

                        Paragraph paragraphTable = new Paragraph();
                        paragraphTable.SpacingBefore = 200f;


                        iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(dataGridView1.Columns.Count);

                        // Add the headers to the PDF
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            table.AddCell(new Phrase(dataGridView1.Columns[j].HeaderText));
                        }

                        table.HeaderRows = 1;

                        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                        {

                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {

                                table.AddCell(new Phrase(dataGridView1.Rows[i].Cells[j].Value.ToString()));

                            }
                        }
                        paragraphTable.Add(table);

                        doc.Add(paragraphTable);
                        doc.Close();
                        MessageBox.Show("File exported succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Can't export the file!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void exportReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReportForm f = new ReportForm();
            f.Show();
        }
    }
}
