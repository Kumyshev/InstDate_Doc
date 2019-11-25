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
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace db
{
    public partial class Form1 : Form
    {



        string path = Path.GetFullPath(@"template.docx");
        string name;
        string age;
        string date;


        


        SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=D:\myPROJECTS\+MURAT\demo\db\Database.mdf;Integrated Security=True");
        SqlDataAdapter adapt;
        SqlCommand cmd;

        public Form1()
        {
            InitializeComponent();
            DisplayData();

            
            

        }


        private void DisplayData()
        {
            con.Open();
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("SELECT * FROM [Table]", con);

           
            adapt.Fill(dt);
            dataGridView.DataSource = dt;
            con.Close();
        }
        private void btnSave_Click(object sender, EventArgs e)
        {

            if (dataGridView.CurrentRow != null)

            {
                DataGridViewRow dgvRow = dataGridView.CurrentRow;
                cmd = new SqlCommand("INSERT INTO [Table] (Name,Age,Date) VALUES (@Name,@Age,@Date)", con);
                cmd.Parameters.AddWithValue("@Name", dgvRow.Cells["Name"].Value == DBNull.Value ? "" : dgvRow.Cells["Name"].Value.ToString());
                cmd.Parameters.AddWithValue("@Age", dgvRow.Cells["Age"].Value == DBNull.Value ? "" : dgvRow.Cells["Age"].Value.ToString());
                cmd.Parameters.AddWithValue("@Date", dgvRow.Cells["Date"].Value == DBNull.Value ? "" : dgvRow.Cells["Date"].Value.ToString());
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
                DisplayData();

            }
        }

        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView.CurrentCell.ColumnIndex == 5)
            {
                e.Control.KeyPress -= AllowNumbersOnly;
                e.Control.KeyPress += AllowNumbersOnly;
            }

        }
        private void AllowNumbersOnly(Object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }              
        }

        private void dataGridView_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (dataGridView.CurrentRow.Cells["id"].Value != DBNull.Value)
            {
                cmd = new SqlCommand("DELETE [Table] WHERE id=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView.CurrentRow.Cells["id"].Value));
                cmd.ExecuteNonQuery();
                con.Close();         
            }           
        }

        private void btnDocMake_Click(object sender, EventArgs e)
        {
            string currentPath = Directory.GetParent(path).FullName;
            
            Word.Application WordApp;
            Word.Document WordDoc;
            WordApp = new Word.Application();
            WordDoc = WordApp.Documents.Open(path);

            ReplaceWord("{name}", name, WordDoc);
            ReplaceWord("{age}", age, WordDoc);
            ReplaceWord("{date}", date, WordDoc);
            WordDoc.SaveAs(currentPath + @"/result.docx");
            WordApp.Visible = true;

        }
        private void ReplaceWord(string toReplace, string text, Word.Document WordDoc)
        {
            var range = WordDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: toReplace, ReplaceWith: text);
        }


        private void dataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            name = dataGridView.Rows[e.RowIndex].Cells[1].Value.ToString();
            age = dataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            date = dataGridView.Rows[e.RowIndex].Cells[3].Value.ToString();
        }
    }
}
