using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;

namespace WindowsFormsApplication1
{
    public partial class AdminForm : Form
    {
        public AdminForm()
        {
            InitializeComponent();
        }

        MySqlConnection connection;
        MySqlDataAdapter dataAdapter;
        MySqlCommandBuilder commandBuilder;
        DataTable dt;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                commandBuilder = new MySqlCommandBuilder(dataAdapter);
                dataAdapter.Update(dt);
                dt = new DataTable();
                dataAdapter.Fill(dt);
                //connection.Close();
                dataGridView1.DataSource = dt;
                MessageBox.Show("Updated succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error while triyng to update!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AdminForm_Load(object sender, EventArgs e)
        {
            string connectionString = "Server = 127.0.0.1; Port = 3306; Database = db_admin; Uid = root; Pwd =; ";
            //string connectionString = "Server = 192.168.1.113; Port = 3306; Database = db_A030; Uid = user_A030; Pwd = pass_A030";
            string querry = "SELECT * FROM users";

            connection = new MySqlConnection(connectionString);

            dataAdapter = new MySqlDataAdapter(querry, connectionString);

            try
            {
                connection.Open();
                //DataSet ds = new DataSet();
                dt = new DataTable();
                dataAdapter.Fill(dt);
                //connection.Close();
                dt.Columns[0].ReadOnly = true;
                dataGridView1.DataSource = dt;
            

            }
            catch (Exception ex)
            {
                MessageBox.Show("There was an error!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            LoginForm f = new LoginForm();
            f.Show();
            this.Close();
        }

        private void exportAsPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            saveFileDialog1.Title = "Save the PDF file!";
            saveFileDialog1.FileName = "users";
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

                        iTextSharp.text.Image PNG = iTextSharp.text.Image.GetInstance(System.Reflection.Assembly.GetExecutingAssembly().Location + "\\..\\..\\..\\Resources\\logoprogram.png");

                        PNG.SetAbsolutePosition(doc.PageSize.Width / 2 - 200f, doc.PageSize.Height - 250f);
                        doc.Add(PNG);

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

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFileDialog1.Title = "Save the Excel file!";
            saveFileDialog1.FileName = "users";
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

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
