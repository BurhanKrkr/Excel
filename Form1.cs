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

namespace Excel_işlemleri
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Burhan\Desktop\Kitaplar.xlsx;Extended Properties='Excel 12.0 Xml; HDR=YES;'");

        void veriler()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("Select * from [Sayfa1$]",baglanti);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            veriler();
        }

        private void BtnListele_Click(object sender, EventArgs e)
        {
            veriler();
        }

        private void BtnKaydet_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("insert into [Sayfa1$] (KitapAd,Yazar,Tür,Fiyat) values(@p1,@p2,@p3,@p4)", baglanti);
            komut.Parameters.AddWithValue("@p1", TxtKitapAd.Text);
            komut.Parameters.AddWithValue("@p2", TxtYazar.Text);
            komut.Parameters.AddWithValue("@p3", TxtTur.Text);
            komut.Parameters.AddWithValue("@p4", TxtFiyat.Text);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Sisteme Kaydedildi","Bilgi",MessageBoxButtons.OK,MessageBoxIcon.Information);
            veriler();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int secilen = dataGridView1.SelectedCells[0].RowIndex;
            TxtKitapAd.Text = dataGridView1.Rows[secilen].Cells[0].Value.ToString();
            TxtYazar.Text = dataGridView1.Rows[secilen].Cells[1].Value.ToString();
            TxtTur.Text = dataGridView1.Rows[secilen].Cells[2].Value.ToString();
            TxtFiyat.Text = dataGridView1.Rows[secilen].Cells[3].Value.ToString();
        }

        private void BtnGuncelle_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("update [Sayfa1$] set Yazar=@p2, Tür=@p3, Fiyat=@p4 where KitapAd=@p1", baglanti);
            komut.Parameters.AddWithValue("@p2", TxtYazar.Text);
            komut.Parameters.AddWithValue("@p3", TxtTur.Text);
            komut.Parameters.AddWithValue("@p4", TxtFiyat.Text);
            komut.Parameters.AddWithValue("@p1", TxtKitapAd.Text);
            komut.ExecuteNonQuery();
            baglanti.Close();
            MessageBox.Show("Kitap Bilgisi Güncellendi", "Güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            veriler();
        }
    }
}
