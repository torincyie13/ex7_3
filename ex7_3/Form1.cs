using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.IO;

namespace ex7_3
{
    public partial class frmPubs : Form
    {
        SqlConnection booksConnection;
        SqlCommand publishersCommand;
        SqlDataAdapter publishersAdapter;
        DataTable publishersTable;
        int pageNumber;
        const int recordsPerPage = 6;

        public frmPubs()
        {
            InitializeComponent();
        }


        private void frmPubs_Load(object sender, EventArgs e)
        {
            // Build interface
            Button[] btnLetter = new Button[27];
            int l = 0;
                int w = Convert.ToInt32(this.ClientRectangle.Width / 3);
            int t = 0;
            int h = Convert.ToInt32(this.ClientRectangle.Height / 9);
            // resize form to fit buttons exactly
            this.ClientSize = new Size(3 * w, 9 * h);
            for (int i = 0; i < 27; i++)
            {
                btnLetter[i] = new Button();
                btnLetter[i].Name = i.ToString();
                btnLetter[i].Left = l;
                btnLetter[i].Width = w;
                btnLetter[i].Top = t;
                btnLetter[i].Height = h;
                btnLetter[i].Text = (char)(i + 65) + " Publishers";
                this.Controls.Add(btnLetter[i]);
                btnLetter[i].Click += new EventHandler(this.btnLetter_Click);
                t += h;
                if (i == 8 || i == 17)
                {
                    l += w;
                    t = 0;
                }
            }
            btnLetter[26].Text = "Other Publishers";
            string path = Path.GetFullPath("SQLBooksDB.mdf");
            // connect to books database
            booksConnection = new
                SqlConnection("Data Source=.\\SQLEXPRESS; AttachDBFilename=" + path + ";" +
                "Integrated Security=True; Connect Timeout=30; User Instance=True");

        }

        private void frmPubs_FormClosing(object sender, FormClosingEventArgs e)
        {
            // dispose of object
            booksConnection.Close();
            booksConnection.Dispose();
            publishersCommand.Dispose();
            publishersAdapter.Dispose();
            publishersTable.Dispose();
        }

        private void btnLetter_Click(object sender, EventArgs e)
        {
            Button whichButton = (Button) sender;
            // Retrieve records
            string sql = "SELECT * FROM Publishers ";
            int i = Convert.ToInt32(whichButton.Name);
            if (i >= 0 && i <= 24) // A to Y
            {
                sql += "WHERE Name >= '" + (char) (i + 65) + "' AND Name < '" + (char) (i + 65 + 1) + "'";
            }
            else if (i == 25) // Z
            {
                sql += "WHERE Name >= 'Z'";
            }
            else // Other
            {
                sql += "WHERE Name < 'A'";
            }
            sql += " ORDER BY Name";
            publishersCommand = new SqlCommand(sql, booksConnection);
            // establish data adapter/data table
            publishersAdapter = new SqlDataAdapter();
            publishersAdapter.SelectCommand = publishersCommand;
            publishersTable = new DataTable();
            publishersAdapter.Fill(publishersTable);
            // set up printdocument
            PrintDocument publishersDocument;
            // create the document and name it
            publishersDocument = new PrintDocument();
            publishersDocument.DocumentName = "Publishers Listing";
            // add code handler
            publishersDocument.PrintPage += new PrintPageEventHandler(this.PrintPublishersPage);
            // print document
            pageNumber = 1;
            // dlgPreview.Document = publishersDocument;
            // dlgPreview.ShowDialog();
            dlgPrint.Document = publishersDocument;
            dlgPrint.ShowDialog();

            DialogResult result = dlgPrint.ShowDialog();
            if (result == DialogResult.OK)
            {
                publishersDocument.Print();
            }
            // dispose of object when done printing
            publishersDocument.Dispose();
        }

        private void PrintPublishersPage(object sender, PrintPageEventArgs e)
        {
            // print headings
            Font myFont = new Font("Arial", 18, FontStyle.Bold);
            int y = Convert.ToInt32(e.MarginBounds.Top);
            e.Graphics.DrawString("Book Publishers Listing - " +
                DateTime.Now.ToString(), myFont, Brushes.Black,
                e.MarginBounds.Left, y);
            y += Convert.ToInt32(myFont.GetHeight());
            e.Graphics.DrawString("Page " + pageNumber.ToString(),
                myFont, Brushes.Black, e.MarginBounds.Left, y);
            y += Convert.ToInt32(myFont.GetHeight()) + 10;
            e.Graphics.DrawLine(Pens.Black, e.MarginBounds.Left, y,
                e.MarginBounds.Right, y);
            y += Convert.ToInt32(myFont.GetHeight());
            myFont = new Font("Courier new", 12, FontStyle.Regular);
            int iEnd = recordsPerPage * pageNumber;
            if (iEnd > publishersTable.Rows.Count)
            {
                iEnd = publishersTable.Rows.Count;
                e.HasMorePages = false;
            }
            else
            {
                e.HasMorePages = true;
            }
            for (int i = recordsPerPage * (pageNumber - 1); i < iEnd; i++)
            {
                // display current record
                e.Graphics.DrawString("Publisher: " +
                    publishersTable.Rows[i]["Name"].ToString(), myFont,
                    Brushes.Black, e.MarginBounds.Left, y);
                y += Convert.ToInt32(myFont.GetHeight());
                e.Graphics.DrawString("Address:   " +
                    publishersTable.Rows[i]["Address"].ToString(), myFont,
                    Brushes.Black, e.MarginBounds.Left, y);
                y += Convert.ToInt32(myFont.GetHeight());
                e.Graphics.DrawString("City:   " +
                    publishersTable.Rows[i]["City"].ToString(), myFont,
                    Brushes.Black, e.MarginBounds.Left, y);
                y += Convert.ToInt32(myFont.GetHeight());
                e.Graphics.DrawString("State:  " +
                    publishersTable.Rows[i]["State"].ToString(), myFont,
                    Brushes.Black, e.MarginBounds.Left, y);
                y += Convert.ToInt32(myFont.GetHeight());
                e.Graphics.DrawString("Zip:   " +
                    publishersTable.Rows[i]["Zip"].ToString(), myFont,
                    Brushes.Black, e.MarginBounds.Left, y);
                y += Convert.ToInt32(myFont.GetHeight());
                y += 2 * Convert.ToInt32(myFont.GetHeight());
            }
            if (e.HasMorePages)
                pageNumber++;
            else
                pageNumber = 1;
        }
    }
}
