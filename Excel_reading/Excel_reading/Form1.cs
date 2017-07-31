using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Excel_reading
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public DataTable ReadCsv(String filename)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection conn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=\"" + Path.GetDirectoryName(filename) + "\";Extended Properties='text;HDR=Yes;FMT=TabDelimited';"))
            {
                using (OleDbCommand command = new OleDbCommand(string.Format("SELECT * from [{0}]", new FileInfo(filename).Name), conn))
                {
                    conn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
                return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV|*.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadCsv(ofd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw;
            }
        }

        // ikinci foksinyon //

        public DataTable ShowTheUserPlayed346Song(String filename)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection conn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=\"" + Path.GetDirectoryName(filename) + "\";Extended Properties='text;HDR=Yes;FMT=TabDelimited';"))
            {
                using (OleDbCommand command = new OleDbCommand(string.Format("SELECT CLIENT_ID, COUNT(*) FROM [{0}] GROUP BY CLIENT_ID HAVING COUNT(CLIENT_ID) = 346", new FileInfo(filename).Name), conn))
                {
                    conn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV|*.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ShowTheUserPlayed346Song(ofd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw;
            }
        }

        // üçüncü fonksiyon//

        public DataTable ShowTheMostPlayedSong(String filename)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection conn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=\"" + Path.GetDirectoryName(filename) + "\";Extended Properties='text;HDR=Yes;FMT=TabDelimited';"))
            {
                using (OleDbCommand command = new OleDbCommand(string.Format("SELECT SONG_ID, COUNT(*) FROM [{0}] GROUP BY SONG_ID ORDER BY COUNT(*) DESC", new FileInfo(filename).Name), conn))
                {
                    conn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV|*.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ShowTheMostPlayedSong(ofd.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw;
            }
        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }        
    }
}
