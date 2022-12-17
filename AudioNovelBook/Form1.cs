using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AudioNovelBook
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer speech = new SpeechSynthesizer();
        OleDbConnection connection;
        public static string dbPath = @"C:\Users\karac\source\repos\AudioNovelBook\AudioNovelBook\bin\Debug\Database1.mdb";
        String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+dbPath;
        public int ID;
        public Form1()
        {
            InitializeComponent();             
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            connection = new OleDbConnection(connectionString);
            readTitleFromDb();
            speech.SelectVoice("Microsoft Stefanos");
        }

        public void readTitleFromDb()
        {
            connection.Open();
            String query = "Select Title from Novels";
            OleDbCommand command = connection.CreateCommand();
            command.CommandText = query;
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add(reader.GetString(0));
            }
            reader.Close();
            connection.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxSelectedItem();
        }
       

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            speech.Rate = trackBar1.Value;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            speech.SpeakAsync(richTextBox1.Text);
            if (speech.State == SynthesizerState.Paused)
            {
                speech.Resume();
            }
            
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            if (speech.State == SynthesizerState.Speaking)
            {
                speech.Pause();
            }
        }

        public void comboBoxSelectedItem()
        {
            richTextBox1.Clear();
            ID = comboBox1.SelectedIndex + 1;
            connection.Open();
            String query = "Select Content from Novels Where ID=" + ID + "";
            OleDbCommand command = connection.CreateCommand();
            command.CommandText = query;
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                richTextBox1.AppendText(reader.GetString(0));
            }
            reader.Close();
            connection.Close();
        }
    }
}
