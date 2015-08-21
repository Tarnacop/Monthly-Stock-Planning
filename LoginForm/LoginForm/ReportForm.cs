using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;

namespace WindowsFormsApplication1
{
    public partial class ReportForm : Form
    {
        public ReportForm()
        {
            InitializeComponent();
        }

        int value1, value2, value3;

        private void ReportForm_Load(object sender, EventArgs e)
        {
            for (int i = 0; i <= 31; i++)
                comboBox1.Items.Add(i);
            comboBox1.SelectedIndex = 0;
            value1 = 0;

            for (int i = 0; i <= 12; i++)
                comboBox2.Items.Add(i);
            comboBox2.SelectedIndex = 0;
            value2 = 0;

            comboBox3.Items.Add(0);
            for (int i = 1900; i <= 2015; i++)
                comboBox3.Items.Add(i);
            comboBox3.SelectedIndex = 0;
            value3 = 0;

            toolTip1.SetToolTip(groupBox1, "0 means it won't consider that part of a date");
        }

        MySqlConnection connection;
        MySqlDataAdapter dataAdapter;
        string connectionString;
        string querry;

        private void button1_Click(object sender, EventArgs e)
        {
            connectionString = "Server = 127.0.0.1; Port = 3306; Database = db_user; Uid = root; Pwd =; ";
            querry = "SELECT DATE(transaction.date) AS 'Date',"
                + " transaction.transactionID AS 'Transaction Number',"
                + " material.name AS 'Material Name',"
                + " contract.quantity AS 'Quantity',"
                + " supplier.name AS 'Supplier'"
                + " FROM transaction, material, contract, supplier"
                + " WHERE";
            
            if (value1 != 0)
            {
                querry = querry + " DAY(transaction.date) = " + value1 + " AND";
            }

            if (value2 != 0)
            {
                querry = querry + " MONTH(transaction.date) = " + value2 + " AND";
            }

            if (value3 != 0)
            {
                querry = querry + " YEAR(transaction.date) = " + value3 + " AND";
            }

            querry = querry + " transaction.materialID = material.materialID AND"
                + " transaction.contractID = contract.contractID AND"
                + " contract.supplierID = supplier.supplierID";
            connection = new MySqlConnection(connectionString);
            dataAdapter = new MySqlDataAdapter(querry, connection);
            DataTable dt = new DataTable();
            connection.Open();
            dataAdapter.Fill(dt);

            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            saveFileDialog1.Title = "Save report file";
            saveFileDialog1.FileName = "report";
            saveFileDialog1.Filter = "*.pdf|*.pdf";
            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                try
                {
                    Document doc = new Document(PageSize.A4);
                    PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(saveFileDialog1.FileName.ToString(), FileMode.Create));
                    doc.Open();

                    iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(System.Reflection.Assembly.GetExecutingAssembly().Location + "\\..\\..\\..\\Resources\\logoprogram.png");
                    logo.SetAbsolutePosition(doc.PageSize.Width / 2 - 200f, doc.PageSize.Height - 250f);
                    doc.Add(logo);

                    Paragraph dateParagraph = new Paragraph("Report created: " + DateTime.Now.ToString("dd-MM-yyyy H:mm:ss"));
                    dateParagraph.SpacingBefore = 200f;
                    doc.Add(dateParagraph);

                    Paragraph tableParagraph;
                    if (dt.Rows.Count == 0)
                    {
                        tableParagraph = new Paragraph("There are no records in the database");
                        doc.Add(tableParagraph);
                        doc.Close();
                        MessageBox.Show("Report created succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                    else
                    {

                        iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(dt.Columns.Count);

                        // Add the headers to the PDF
                        foreach (DataColumn column in dt.Columns)
                        {
                            table.AddCell(new Phrase(column.ColumnName));
                        }

                        table.HeaderRows = 1;

                        List<string> dates = new List<string>();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DateTime date = DateTime.ParseExact(dt.Rows[i][0].ToString(), "M/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
                            string s = date.ToString("dd-MM-yyyy");
                            dates.Add(s);
                        }
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            table.AddCell(new Phrase(dates.ElementAt(i)));
                            for (int j = 1; j < dt.Columns.Count; j++)
                            {
                                table.AddCell(new Phrase(dt.Rows[i][j].ToString()));
                            }
                        }
                        tableParagraph = new Paragraph();
                        tableParagraph.Add(table);
                        tableParagraph.SpacingBefore = 30f;
                        doc.Add(tableParagraph);

                        int sum = 0;
                        float avg;

                        for (int i = 0; i < dt.Rows.Count; i++) sum = sum + Int32.Parse(dt.Rows[i]["Quantity"].ToString());
                        avg = sum / dt.Rows.Count;

                        Paragraph conclusionParagraph = new Paragraph("The average quantity bought from suppliers based on the date selected is " + avg);
                        conclusionParagraph.SpacingBefore = 30f;
                        doc.Add(conclusionParagraph);

                        doc.Close();
                        MessageBox.Show("Report created succesfully!", "Succes!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There was an error while trying  to create the report!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            value1 = (int)comboBox1.SelectedItem;
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            value2 = (int)comboBox2.SelectedItem;
        }

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            value3 = (int)comboBox3.SelectedItem;
        }
    }
}
