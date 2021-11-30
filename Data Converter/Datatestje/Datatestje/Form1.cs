using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Datatestje
{
    #region Form
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #endregion


        //You can only use this program with .xls and .mdb files. Further to use it your self you will need to change some values further in the code.
        // This has all been written out by me further down the code.
        #region Buttons
        private void btnrun_Click(object sender, EventArgs e)
        {
            string EXpath = textBox1.Text;
            string PathConn = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" + EXpath + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(PathConn);

            try
            {
                string EXsheet = textBox3.Text;
                var sqlQuery = "Select * from [" + EXsheet + "$]";
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(sqlQuery, conn);
                DataTable dt = new DataTable();
                myDataAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            catch
            { 
                
            }
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            {
                string EXpath = textBox1.Text;
                string fileNameExcel = @EXpath;
                string ACpath = textBox2.Text;
                string fileNameAccess = @ACpath;
                string EXsheet = textBox3.Text;

                string connectionStringExcel =
                    string.Format("Data Source= {0};Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;", fileNameExcel);

                string ConnectionStringAccess =
                    string.Format("Data Source= {0};Provider=Microsoft.Jet.OLEDB.4.0; Persist security Info = false", fileNameAccess);

                OleDbConnection connExcel = new OleDbConnection(connectionStringExcel);
                OleDbConnection connAccess = new OleDbConnection(ConnectionStringAccess);
                OleDbCommand cmdExcel = connExcel.CreateCommand();
                cmdExcel.CommandType = CommandType.Text;
                cmdExcel.CommandText = "SELECT * from [" + EXsheet + "$]";
                OleDbCommand cmdAccess = connAccess.CreateCommand();
                cmdAccess.CommandType = CommandType.Text;
                

                //Add parameters *
                //Here you need to change the names of the columns.
                //These should be called the same in Excel and Access.
                //Also change - Informatie - To the name of your Access Database Table.
                cmdAccess.CommandText = "INSERT INTO Informatie (Naam, Achternaam, Land, Stad, Huisnummer, Postcode, Telefoonnummer) VALUES(@Naam, @Achternaam, @Land, @Stad, @Huisnummer, @Postcode, @Telefoonnummer)";


                //Add parameters to Access command object **
                //All you need to do here is change the names of the columns.
                //These should be called the same in Excel and Access.
                OleDbParameter param1 = new OleDbParameter("@Naam", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param1);
                OleDbParameter param2 = new OleDbParameter("@Achternaam", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param2);
                OleDbParameter param3 = new OleDbParameter("@Land", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param3);
                OleDbParameter param4 = new OleDbParameter("@Stad", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param4);
                OleDbParameter param5 = new OleDbParameter("@Huisnummer", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param5);
                OleDbParameter param6 = new OleDbParameter("@Postcode", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param6);
                OleDbParameter param7 = new OleDbParameter("@Telefoonnummer", OleDbType.VarChar);
                cmdAccess.Parameters.Add(param7);

                connExcel.Open();
                connAccess.Open();
                OleDbDataReader drExcel = cmdExcel.ExecuteReader();

                while (drExcel.Read())
                {

                    //Assign values to access command parameters ***
                    // To add more columns you just copy what it has and add more. F.E - param8.Value = drExcel[7].ToString();
                    param1.Value = drExcel[0].ToString();
                    param2.Value = drExcel[1].ToString();
                    param3.Value = drExcel[2].ToString();
                    param4.Value = drExcel[3].ToString();
                    param5.Value = drExcel[4].ToString();
                    param6.Value = drExcel[5].ToString();
                    param7.Value = drExcel[6].ToString();

                    cmdAccess.ExecuteNonQuery();
                }

                connAccess.Close();
                connExcel.Close();
                MessageBox.Show("Succesfully uploaded Excel data to Database.");

            }
        }

        private void btnbrowse_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.ShowDialog();
            openfiledialog1.Filter = "allfiles|*.xls";
            textBox1.Text = openfiledialog1.FileName;
        }

        private void btnbrowse2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog1 = new OpenFileDialog();
            openfiledialog1.ShowDialog();
            openfiledialog1.Filter = "allfiles|*.mdb";
            textBox2.Text = openfiledialog1.FileName;
        }
        #endregion


        #region Extra


        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        #endregion
    }
}