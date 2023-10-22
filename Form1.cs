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
using System.Text.RegularExpressions;

namespace Excel_Düzenleme_Programı
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        public static string excelYol = "";
        public static OleDbConnection baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; 
             Data Source =" + excelYol + ";" +
             "Extended Properties= 'Excel 12.0 Xml;HDR=YES'");

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)//Form Kontrollü Kapatma
        {
            DialogResult result = MessageBox.Show("Uygulamadan Çıkmak İstediğinize Emin misiniz?", "Uygulamadan Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
        private void Form1_Load(object sender, EventArgs e)//Dosya Yolu Olmadığında Form KOMPLE KAPALI Haldedir. 
        {
            //Add_Picture.Enabled = false;
            Sil_Button.Enabled = false;
            TümKayıtlarınNoSil.Enabled = false;
            Mail_Sil.Enabled = false;
            Listele_Button.Enabled = false;
            Ekle_Button.Enabled = false;
            Güncelle_Button.Enabled = false;
            ÖzelSil_button.Enabled = false;
            label15.Visible = false;//İcon Metni Gözükmez

        }

        private void Add_Picture_Click(object sender, EventArgs e)//Dosya Ekleme İconu
        {
            if (string.IsNullOrEmpty(textBox13.Text))
            {
                openFileDialog1.ShowDialog();
                textBox13.Text = openFileDialog1.FileName;
                excelYol = textBox13.Text;

                baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; 
             Data Source =" + excelYol + ";" +
                 "Extended Properties= 'Excel 12.0 Xml;HDR=YES'");
                Listele_metod();
                Btn_acma();
            }
            else if (!string.IsNullOrEmpty(textBox13.Text))
            {
                openFileDialog1.ShowDialog();
                textBox13.Text = openFileDialog1.FileName;
                excelYol = textBox13.Text;

                baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; 
                Data Source =" + excelYol + ";" +
                 "Extended Properties= 'Excel 12.0 Xml;HDR=YES'");
                Listele_metod();
                Btn_acma();
            }
        }

        void Listele_metod() // Listeleme Metodu
        {
            //SADECE ŞUANLIK BAZI SÜTUNLAR GÖZÜKMESİ İÇİN AYARLANDI. DAHA SONRA TÜM TABLO SÜTUNLARINI GÖSTERİCEZ.

            OleDbDataAdapter da = new OleDbDataAdapter("select id, firma_adi, sektor, isim, il, adres, email, ilce, tel, tel2, faks, cep_tel from [sayfa1$]", baglanti); //SADECE ŞUANLIK BAZI SÜTUNLAR GÖZÜKMESİ İÇİN AYARLANDI. DAHA SONRA TÜM TABLO SÜTUNLARINI GÖSTERİCEZ.
            DataTable dt = new DataTable();

            if (da.Fill(dt) > 0)
            {
                dataGridView1.DataSource = dt; // BU KISIMDA, Özel Silme Alanında null referance hatası aldım
            }
            else
            {
                dataGridView1.DataSource = null;
            }

            // Form Listelendiğinde Otomatik Olarak 1.Satır Seçili Olarak Gelir
            int sec = dataGridView1.SelectedCells[0].RowIndex;
            textBox1.Text = dataGridView1.Rows[sec].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[sec].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.Rows[sec].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[sec].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[sec].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[sec].Cells[5].Value.ToString();
            textBox7.Text = dataGridView1.Rows[sec].Cells[6].Value.ToString();
            textBox8.Text = dataGridView1.Rows[sec].Cells[7].Value.ToString();
            textBox9.Text = dataGridView1.Rows[sec].Cells[8].Value.ToString();
            textBox10.Text = dataGridView1.Rows[sec].Cells[9].Value.ToString();
            textBox11.Text = dataGridView1.Rows[sec].Cells[10].Value.ToString();
            textBox12.Text = dataGridView1.Rows[sec].Cells[11].Value.ToString();
        }

        private void Listele_Button_Click(object sender, EventArgs e) // Listeleme İşlemi
        {
            Listele_metod();
        }

        //-----------------------------------------------------------------------------------------------------------------------------------
        // Sadece sayısal verilerin girişine izin veren Kod Parçası.
        // - , (), Boşluk ve Null değer girişine izin vermez ve Boş Değer alamadığı için zorunlu alan olmuş olur.
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)// İD ALANI-Veri Engeli
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)// TELEFON 1 ALANI-Veri Engeli
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)// TELEFON 2 ALANI-Veri Engeli
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        //private void textBox11_KeyPress(object sender, KeyPressEventArgs e)// FAKS ALANI-Veri Engeli
        //{
        //    if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
        //    {
        //        e.Handled = true;
        //    }
        //}
         
        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)// CEP_TEL ALANI-Veri Engeli
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; 
            }
        }
//----------------------------------------------------------------------------------------------------------------------------------------
        private void Ekle_Button_Click(object sender, EventArgs e)//Kayıt Ekleme 
        {
            baglanti.Open();

            int id, tel1, tel2, faks, cep_tel;
            bool isNumeric = int.TryParse(textBox1.Text, out id);

            // ID kısmının, sadece sayısal bir değer olmasını kontrol etme
            if (!isNumeric)
            {
                MessageBox.Show("Lütfen İD Kısmına Rakamlardan Oluşan Sayısal Bir Değer Giriniz", "Uygunsuz Veri Girişi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                baglanti.Close();
                return;
            }


            // Tel, Tel2, Faks, Cep_Tel alanlarının, sadece sayısal bir değer olmasını kontrol etme
            if (!int.TryParse(textBox9.Text, out tel1) || !int.TryParse(textBox10.Text, out tel2) || !int.TryParse(textBox11.Text, out faks) || !int.TryParse(textBox12.Text, out cep_tel))
            {
                MessageBox.Show("Lütfen telefon, faks ve cep telefonu kısımlarına sadece sayısal değerler giriniz", "Uygunsuz Veri Girişi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                baglanti.Close();
                return;
            }

            OleDbCommand ekle = new OleDbCommand("insert into [sayfa1$] (id,firma_adi,sektor,isim,il,adres,email,ilce,tel,tel2,faks,cep_tel) values (@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12)", baglanti);
            //progressBar1.Value = 100;
            ekle.Parameters.AddWithValue("@p1", id);
            ekle.Parameters.AddWithValue("@p2", textBox2.Text);
            ekle.Parameters.AddWithValue("@p3", textBox3.Text);
            ekle.Parameters.AddWithValue("@p4", textBox4.Text);
            ekle.Parameters.AddWithValue("@p5", textBox5.Text);
            ekle.Parameters.AddWithValue("@p6", textBox6.Text);
            ekle.Parameters.AddWithValue("@p7", textBox7.Text);
            ekle.Parameters.AddWithValue("@p8", textBox8.Text);
            ekle.Parameters.AddWithValue("@p9", tel1);
            ekle.Parameters.AddWithValue("@p10", tel2);
            ekle.Parameters.AddWithValue("@p11", faks);
            ekle.Parameters.AddWithValue("@p12", cep_tel);
            ekle.ExecuteNonQuery();
            MessageBox.Show("Ekleme İşlemi Başarıyla Gerçekleşmiştir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            baglanti.Close();
            Listele_metod();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)//Satırdaki Kayıtları Seçme 
        {
            int sec = dataGridView1.SelectedCells[0].RowIndex;

            textBox1.Text = dataGridView1.Rows[sec].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[sec].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.Rows[sec].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[sec].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[sec].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[sec].Cells[5].Value.ToString();
            textBox7.Text = dataGridView1.Rows[sec].Cells[6].Value.ToString();
            textBox8.Text = dataGridView1.Rows[sec].Cells[7].Value.ToString();
            textBox9.Text = dataGridView1.Rows[sec].Cells[8].Value.ToString();
            textBox10.Text = dataGridView1.Rows[sec].Cells[9].Value.ToString();
            textBox11.Text = dataGridView1.Rows[sec].Cells[10].Value.ToString();
            textBox12.Text = dataGridView1.Rows[sec].Cells[11].Value.ToString();
        }
        //private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) SATIRDAKİ KAYITLARI SEÇMENİN 2. YOLU
        //{
        //    if (e.RowIndex >= 0)
        //    {
        //        DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
        //        textBox1.Text = row.Cells["id"].Value.ToString();
        //        textBox2.Text = row.Cells["firma_adi"].Value.ToString();
        //        textBox3.Text = row.Cells["sektor"].Value.ToString();
        //        textBox4.Text = row.Cells["isim"].Value.ToString();
        //        textBox5.Text = row.Cells["il"].Value.ToString();
        //        textBox6.Text = row.Cells["adres"].Value.ToString();
        //        textBox7.Text = row.Cells["email"].Value.ToString();
        //        textBox8.Text = row.Cells["ilce"].Value.ToString();
        //        textBox9.Text = row.Cells["tel"].Value.ToString();
        //        textBox10.Text = row.Cells["tel2"].Value.ToString();
        //        textBox11.Text = row.Cells["faks"].Value.ToString();
        //        textBox12.Text = row.Cells["cep_tel"].Value.ToString();
        //    }
        //}
       

        //private void Güncelle_Button_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+excelYol;
        //        using (OleDbConnection connection = new OleDbConnection(connectionString))
        //        {
        //            connection.Open();
        //            OleDbCommand command = new OleDbCommand("UPDATE sayfa1 SET firma_adi=?, sektor=?, isim=?, il=?, adres=?, email=?, ilce=?, tel=?, tel2=?, faks=?, cep_tel=? WHERE id=?", connection);
        //            command.Parameters.AddWithValue("@p2", textBox2.Text);
        //            command.Parameters.AddWithValue("@p3", textBox3.Text);
        //            command.Parameters.AddWithValue("@p4", textBox4.Text);
        //            command.Parameters.AddWithValue("@p5", textBox5.Text);
        //            command.Parameters.AddWithValue("@p6", textBox6.Text);
        //            command.Parameters.AddWithValue("@p7", textBox7.Text);
        //            command.Parameters.AddWithValue("@p8", textBox8.Text);
        //            command.Parameters.AddWithValue("@p9", textBox9.Text);
        //            command.Parameters.AddWithValue("@p10", textBox10.Text);
        //            command.Parameters.AddWithValue("@p11", textBox11.Text);
        //            command.Parameters.AddWithValue("@p12", textBox12.Text);
        //            command.Parameters.AddWithValue("@p1", textBox1.Text);
        //            int rowsUpdated = command.ExecuteNonQuery();
        //            if (rowsUpdated > 0)
        //            {
        //                MessageBox.Show("Tablo Verileriniz Güncellenmiştir.", "Güncelleme Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                Listele_metod();
        //            }
        //            else
        //            {
        //                MessageBox.Show("Güncelleme işlemi başarısız oldu.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Bir hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}


        private void Güncelle_Button_Click(object sender, EventArgs e)//Kayıt Güncelleme
        {
            baglanti.Open();
            OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set firma_adi=@p2, sektor=@p3, isim=@p4, il=@p5, adres=@p6, email=@p7, ilce=@p8, tel=@p9, tel2=@p10, faks=@p11, cep_tel=@p12 where id=@p1", baglanti);

            güncelle.Parameters.AddWithValue("@p2", textBox2.Text);
            güncelle.Parameters.AddWithValue("@p3", textBox3.Text);
            güncelle.Parameters.AddWithValue("@p4", textBox4.Text);
            güncelle.Parameters.AddWithValue("@p5", textBox5.Text);
            güncelle.Parameters.AddWithValue("@p6", textBox6.Text);
            güncelle.Parameters.AddWithValue("@P7", textBox7.Text);
            güncelle.Parameters.AddWithValue("@p8", textBox8.Text);
            güncelle.Parameters.AddWithValue("@p9", textBox9.Text);
            güncelle.Parameters.AddWithValue("@p10", textBox10.Text);
            güncelle.Parameters.AddWithValue("@p11", textBox11.Text);
            güncelle.Parameters.AddWithValue("@p12", textBox12.Text);
            güncelle.Parameters.AddWithValue("@p1", textBox1.Text);

            var sonuc = güncelle.ExecuteNonQuery();
            dataGridView1.DataSource = null;
            Listele_metod();
            MessageBox.Show("Tablo Verileriniz Güncellenmistir.", "Güncelleme Bilgisi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            baglanti.Close();
        }

        private void Sil_Button_Click(object sender, EventArgs e) //SEÇİLİ OLAN SATIRDAKİ Numara Verilerini Temizler
        {
            Btn_kapatma(); 
            string degisecek1 = "-";
            string degismis1 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek1, degismis1);
            textBox10.Text = textBox10.Text.Replace(degisecek1, degismis1);
            textBox11.Text = textBox11.Text.Replace(degisecek1, degismis1);
            textBox12.Text = textBox12.Text.Replace(degisecek1, degismis1);

            string degisecek2 = "()";
            string degismis2 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek2, degismis2);
            textBox10.Text = textBox10.Text.Replace(degisecek2, degismis2);
            textBox11.Text = textBox11.Text.Replace(degisecek2, degismis2);
            textBox12.Text = textBox12.Text.Replace(degisecek2, degismis2);

            string degisecek3 = "(   )";
            string degismis3 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek3, degismis3);
            textBox10.Text = textBox10.Text.Replace(degisecek3, degismis3);
            textBox11.Text = textBox11.Text.Replace(degisecek3, degismis3);
            textBox12.Text = textBox12.Text.Replace(degisecek3, degismis3);

            string degisecek4 = "(";
            string degismis4 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek4, degismis4);
            textBox10.Text = textBox10.Text.Replace(degisecek4, degismis4);
            textBox11.Text = textBox11.Text.Replace(degisecek4, degismis4);
            textBox12.Text = textBox12.Text.Replace(degisecek4, degismis4);

            string degisecek5 = ")";
            string degismis5 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek5, degismis5);
            textBox10.Text = textBox10.Text.Replace(degisecek5, degismis5);
            textBox11.Text = textBox11.Text.Replace(degisecek5, degismis5);
            textBox12.Text = textBox12.Text.Replace(degisecek5, degismis5);

            string degisecek6 = "    ";
            string degismis6 = "";
            textBox9.Text = textBox9.Text.Replace(degisecek6, degismis6);
            textBox10.Text = textBox10.Text.Replace(degisecek6, degismis6);
            textBox11.Text = textBox11.Text.Replace(degisecek6, degismis6);
            textBox12.Text = textBox12.Text.Replace(degisecek6, degismis6);
            //  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --  --
            baglanti.Open();
            OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set firma_adi=@p2, sektor=@p3, isim=@p4, il=@p5, adres=@p6, email=@p7, ilce=@p8, tel=@p9, tel2=@p10, faks=@p11, cep_tel=@p12 where id=@p1", baglanti);

            güncelle.Parameters.AddWithValue("@p2", textBox2.Text);
            güncelle.Parameters.AddWithValue("@p3", textBox3.Text);
            güncelle.Parameters.AddWithValue("@p4", textBox4.Text);
            güncelle.Parameters.AddWithValue("@p5", textBox5.Text);
            güncelle.Parameters.AddWithValue("@p6", textBox6.Text);
            güncelle.Parameters.AddWithValue("@P7", textBox7.Text);
            güncelle.Parameters.AddWithValue("@p8", textBox8.Text);
            güncelle.Parameters.AddWithValue("@p9", textBox9.Text);
            güncelle.Parameters.AddWithValue("@p10", textBox10.Text);
            güncelle.Parameters.AddWithValue("@p11", textBox11.Text);
            güncelle.Parameters.AddWithValue("@p12", textBox12.Text);
            güncelle.Parameters.AddWithValue("@p1", textBox1.Text);

            güncelle.ExecuteNonQuery();
            Listele_metod();
            baglanti.Close();

            MessageBox.Show("Telefon Alanına Girilen Uygunsuz Veriler Temizlenmiştir." +
                " İyi Çalışmalar Efendim.", "Temizleme Mesajı",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            Btn_acma();
        }

        private void TümKayıtlarınNoSil_Click(object sender, EventArgs e) //TÜM KAYITLARIN NUMARA VERİLERİNİ TEMİZLEYEN  BackGroundWorker YAPISI
        {
            Btn_kapatma();
            try
            {
                backgroundWorker1.RunWorkerAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "/n/n" + ex.ToString());
            }
        }

    //    private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
    //    {
    //      try
    //      {
    //            progressBar1.Maximum = dataGridView1.Rows.Count + 1;

    //            foreach (DataGridViewRow row in dataGridView1.Rows)
    //            {
    //                progressBar1.Value = row.Index + 1;

    //                string id = row.Cells[0].Value.ToString();
    //                string firma_adi = row.Cells[1].Value.ToString();
    //                string sektor = row.Cells[2].Value.ToString();
    //                string isim = row.Cells[3].Value.ToString();
    //                string il = row.Cells[4].Value.ToString();
    //                string adres = row.Cells[5].Value.ToString();
    //                string email = row.Cells[6].Value.ToString();
    //                string ilce = row.Cells[7].Value.ToString();
    //                string tel = row.Cells[8].Value.ToString();
    //                string tel2 = row.Cells[9].Value.ToString();
    //                string faks = row.Cells[10].Value.ToString();
    //                string cep_tel = row.Cells[11].Value.ToString();
                    
    //                tel = tel.Replace("-", "");
    //                tel2 = tel2.Replace("-", "");
    //                faks = faks.Replace("-", "");
    //                cep_tel = cep_tel.Replace("-", "");

    //                tel = tel.Replace("(", "");
    //                tel2 = tel2.Replace("(", "");
    //                faks = faks.Replace("(", "");
    //                cep_tel = cep_tel.Replace("(", "");

    //                tel = tel.Replace(")", "");
    //                tel2 = tel2.Replace(")", "");
    //                faks = faks.Replace(")", "");
    //                cep_tel = cep_tel.Replace(")", "");

    //                tel = tel.Replace(" ", "");
    //                tel2 = tel2.Replace(" ", "");
    //                faks = faks.Replace(" ", "");
    //                cep_tel = cep_tel.Replace(" ", "");

    //                baglanti.Open();

    //                OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set firma_adi=@p2, sektor=@p3, isim=@p4, il=@p5, adres=@p6, email=@p7, ilce=@p8, tel=@p9, tel2=@p10, faks=@p11, cep_tel=@p12 where id=@p1", baglanti);

    //                güncelle.Parameters.AddWithValue("@p1", id);
    //                güncelle.Parameters.AddWithValue("@p2", firma_adi);
    //                güncelle.Parameters.AddWithValue("@p3", sektor);
    //                güncelle.Parameters.AddWithValue("@p4", isim);
    //                güncelle.Parameters.AddWithValue("@p5", il);
    //                güncelle.Parameters.AddWithValue("@p6", adres);
    //                güncelle.Parameters.AddWithValue("@p7", email);
    //                güncelle.Parameters.AddWithValue("@p8", ilce);
    //                güncelle.Parameters.AddWithValue("@p9", tel);
    //                güncelle.Parameters.AddWithValue("@10", tel2);
    //                güncelle.Parameters.AddWithValue("@11", faks);
    //                güncelle.Parameters.AddWithValue("@12", cep_tel);

    //                güncelle.ExecuteNonQuery();
    //                baglanti.Close();
    //            }
    //    }
    //        catch
    //        {
    //            MessageBox.Show("Güncelleme Başarısız");
    //        }
    //}
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) // TÜM KAYITLARIN NUMARA VERİLERİNİ TEMİZLEYEN  BackGroundWorker YAPISI
        {

            progressBar1.Maximum = dataGridView1.Rows.Count + 1;
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                progressBar1.Value = dataGridView1.Rows.Count + 1;
                progressBar1.Value = item.Index;

                    textBox1.Text = item.Cells[0].Value.ToString();
                    textBox2.Text = item.Cells[1].Value.ToString();
                    textBox3.Text = item.Cells[2].Value.ToString();
                    textBox4.Text = item.Cells[3].Value.ToString();
                    textBox5.Text = item.Cells[4].Value.ToString();
                    textBox6.Text = item.Cells[5].Value.ToString();
                    textBox7.Text = item.Cells[6].Value.ToString();
                    textBox8.Text = item.Cells[7].Value.ToString();
                    textBox9.Text = item.Cells[8].Value.ToString();
                    textBox10.Text = item.Cells[9].Value.ToString();
                    textBox11.Text = item.Cells[10].Value.ToString();
                    textBox12.Text = item.Cells[11].Value.ToString();

                    string degisecek1 = "-";
                    string degismis1 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek1, degismis1);
                    textBox10.Text = textBox10.Text.Replace(degisecek1, degismis1);
                    textBox11.Text = textBox11.Text.Replace(degisecek1, degismis1);
                    textBox12.Text = textBox12.Text.Replace(degisecek1, degismis1);

                    string degisecek2 = "()";
                    string degismis2 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek2, degismis2);
                    textBox10.Text = textBox10.Text.Replace(degisecek2, degismis2);
                    textBox11.Text = textBox11.Text.Replace(degisecek2, degismis2);
                    textBox12.Text = textBox12.Text.Replace(degisecek2, degismis2);

                    string degisecek3 = "(   )";
                    string degismis3 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek3, degismis3);
                    textBox10.Text = textBox10.Text.Replace(degisecek3, degismis3);
                    textBox11.Text = textBox11.Text.Replace(degisecek3, degismis3);
                    textBox12.Text = textBox12.Text.Replace(degisecek3, degismis3);

                    string degisecek4 = "(";
                    string degismis4 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek4, degismis4);
                    textBox10.Text = textBox10.Text.Replace(degisecek4, degismis4);
                    textBox11.Text = textBox11.Text.Replace(degisecek4, degismis4);
                    textBox12.Text = textBox12.Text.Replace(degisecek4, degismis4);

                    string degisecek5 = ")";
                    string degismis5 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek5, degismis5);
                    textBox10.Text = textBox10.Text.Replace(degisecek5, degismis5);
                    textBox11.Text = textBox11.Text.Replace(degisecek5, degismis5);
                    textBox12.Text = textBox12.Text.Replace(degisecek5, degismis5);

                    string degisecek6 = "    ";
                    string degismis6 = "";
                    textBox9.Text = textBox9.Text.Replace(degisecek6, degismis6);
                    textBox10.Text = textBox10.Text.Replace(degisecek6, degismis6);
                    textBox11.Text = textBox11.Text.Replace(degisecek6, degismis6);
                    textBox12.Text = textBox12.Text.Replace(degisecek6, degismis6);

                    baglanti.Open();
                    OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set firma_adi=@p2, sektor=@p3, isim=@p4, il=@p5, adres=@p6, email=@p7, ilce=@p8, tel=@p9, tel2=@p10, faks=@p11, cep_tel=@p12 where id=@p1", baglanti);

                    güncelle.Parameters.AddWithValue("@p2", textBox2.Text);
                    güncelle.Parameters.AddWithValue("@p3", textBox3.Text);
                    güncelle.Parameters.AddWithValue("@p4", textBox4.Text);
                    güncelle.Parameters.AddWithValue("@p5", textBox5.Text);
                    güncelle.Parameters.AddWithValue("@p6", textBox6.Text);
                    güncelle.Parameters.AddWithValue("@P7", textBox7.Text);
                    güncelle.Parameters.AddWithValue("@p8", textBox8.Text);
                    güncelle.Parameters.AddWithValue("@p9", textBox9.Text);
                    güncelle.Parameters.AddWithValue("@p10", textBox10.Text);
                    güncelle.Parameters.AddWithValue("@p11", textBox11.Text);
                    güncelle.Parameters.AddWithValue("@p12", textBox12.Text);
                    güncelle.Parameters.AddWithValue("@p1", textBox1.Text);

                    güncelle.ExecuteNonQuery();
                    baglanti.Close();
                
            }
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                MessageBox.Show("Bütün Telefon Verileri Temizlenmiştir", "Temizleme Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Btn_acma();
                Listele_metod();
                progressBar1.Value = 0;
            }

            private void ÖzelSil_button_Click(object sender, EventArgs e)//CheckedListBox'dan Sildirme İşlemi
            { 
                backgroundWorker2.RunWorkerAsync();
                Btn_kapatma();
            }
            private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)//checkedListBox1' de Sadece Tek Bir Seçim Yaptıran Kod Parçacığı
            {
                // seçili olan öğe değiştiriliyor
                if (e.NewValue == CheckState.Checked)
                {
                    // diğer öğelerin durumu false yapılıyor
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (i != e.Index)
                        {
                            checkedListBox1.SetItemChecked(i, false);
                        }
                    }
                }
            }
            private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)//CheckedListBox'dan Sildirme İşlemi
            {
                if (checkedListBox1.SelectedIndex == 0) // TELEFON 1 ALANI 
                {
                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        progressBar1.Maximum = dataGridView1.Rows.Count + 1;
                        progressBar1.Value = item.Index;

                        if (item.Cells[8].Value != null && !string.IsNullOrEmpty(item.Cells[8].Value.ToString()))
                        {
                            string strVal = item.Cells[8].Value.ToString();
                            if (!string.IsNullOrEmpty(strVal))
                            {
                                string degisecek1 = "-";
                                string degismis1 = "";
                                item.Cells[8].Value = strVal.Replace(degisecek1, degismis1);

                                string degisecek2 = "()";
                                string degismis2 = "";
                                item.Cells[8].Value = item.Cells[8].Value.ToString().Replace(degisecek2, degismis2);

                                string degisecek3 = "(";
                                string degismis3 = "";
                                item.Cells[8].Value = item.Cells[8].Value.ToString().Replace(degisecek3, degismis3);

                                string degisecek4 = ")";
                                string degismis4 = "";
                                item.Cells[8].Value = item.Cells[8].Value.ToString().Replace(degisecek4, degismis4);

                                string degisecek5 = "    ";
                                string degismis5 = "";
                                item.Cells[8].Value = item.Cells[8].Value.ToString().Replace(degisecek5, degismis5);

                                baglanti.Open();
                                OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set tel=@p9 where id=@p1", baglanti);
                                güncelle.Parameters.AddWithValue("@p9", item.Cells[8].Value);
                                güncelle.Parameters.AddWithValue("@p1", textBox1.Text);
                                güncelle.ExecuteNonQuery();
                                baglanti.Close();
                            }
                        }
                    }
                }

                else if (checkedListBox1.SelectedIndex == 1)//TELEFON 2 ALANI 
                {
                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        progressBar1.Maximum = dataGridView1.Rows.Count + 1;
                        progressBar1.Value = item.Index;
                        if (item.Cells[9].Value != null && !string.IsNullOrEmpty(item.Cells[9].Value.ToString()))
                        {
                            string strVal = item.Cells[9].Value.ToString();
                            if (!string.IsNullOrEmpty(strVal))
                            {
                                string degisecek1 = "-";
                                string degismis1 = "";
                                item.Cells[9].Value = strVal.Replace(degisecek1, degismis1);

                                string degisecek2 = "()";
                                string degismis2 = "";
                                item.Cells[9].Value = item.Cells[9].Value.ToString().Replace(degisecek2, degismis2);

                                string degisecek3 = "(";
                                string degismis3 = "";
                                item.Cells[9].Value = item.Cells[9].Value.ToString().Replace(degisecek3, degismis3);

                                string degisecek4 = ")";
                                string degismis4 = "";
                                item.Cells[9].Value = item.Cells[9].Value.ToString().Replace(degisecek4, degismis4);

                                string degisecek5 = "    ";
                                string degismis5 = "";
                                item.Cells[9].Value = item.Cells[9].Value.ToString().Replace(degisecek5, degismis5);

                                baglanti.Open();
                                OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set tel2=@p10 where id=@p1", baglanti);
                                güncelle.Parameters.AddWithValue("@p10", item.Cells[9].Value);
                                güncelle.Parameters.AddWithValue("@p1", textBox1.Text);
                                baglanti.Close();

                            }
                        }
                    }
                }

                else if (checkedListBox1.SelectedIndex == 2)// FAKS ALANI
                {
                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        progressBar1.Maximum = dataGridView1.Rows.Count + 1;
                        progressBar1.Value = item.Index;

                        if (item.Cells[9].Value != null && !string.IsNullOrEmpty(item.Cells[9].Value.ToString()))
                        {
                            string strVal = item.Cells[10].Value.ToString();
                            if (!string.IsNullOrEmpty(strVal))
                            {
                                string degisecek1 = "-";
                                string degismis1 = "";
                                item.Cells[10].Value = strVal.Replace(degisecek1, degismis1);

                                string degisecek2 = "()";
                                string degismis2 = "";
                                item.Cells[10].Value = item.Cells[10].Value.ToString().Replace(degisecek2, degismis2);

                                string degisecek3 = "(";
                                string degismis3 = "";
                                item.Cells[10].Value = item.Cells[10].Value.ToString().Replace(degisecek3, degismis3);

                                string degisecek4 = ")";
                                string degismis4 = "";
                                item.Cells[10].Value = item.Cells[10].Value.ToString().Replace(degisecek4, degismis4);

                                string degisecek5 = "    ";
                                string degismis5 = "";
                                item.Cells[10].Value = item.Cells[10].Value.ToString().Replace(degisecek5, degismis5);

                                baglanti.Open();
                                OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set faks=@p11 where id=@p1", baglanti);
                                güncelle.Parameters.AddWithValue("@p11", item.Cells[10].Value);
                                güncelle.Parameters.AddWithValue("@p1", textBox1.Text);
                                baglanti.Close();
                            }
                        }
                    }
                }

                else if (checkedListBox1.SelectedIndex == 3)// CEP TELEFONU ALANI
                {
                    foreach (DataGridViewRow item in dataGridView1.Rows)
                    {
                        progressBar1.Maximum = dataGridView1.Rows.Count + 1;
                        progressBar1.Value = item.Index;

                        if (item.Cells[9].Value != null && !string.IsNullOrEmpty(item.Cells[9].Value.ToString()))
                        {
                            string strVal = item.Cells[11].Value.ToString();
                            if (!string.IsNullOrEmpty(strVal))
                            {
                                string degisecek1 = "-";
                                string degismis1 = "";
                                item.Cells[11].Value = strVal.Replace(degisecek1, degismis1);

                                string degisecek2 = "()";
                                string degismis2 = "";
                                item.Cells[11].Value = item.Cells[11].Value.ToString().Replace(degisecek2, degismis2);

                                string degisecek3 = "(";
                                string degismis3 = "";
                                item.Cells[11].Value = item.Cells[11].Value.ToString().Replace(degisecek3, degismis3);

                                string degisecek4 = ")";
                                string degismis4 = "";
                                item.Cells[11].Value = item.Cells[11].Value.ToString().Replace(degisecek4, degismis4);

                                string degisecek5 = "    ";
                                string degismis5 = "";
                                item.Cells[11].Value = item.Cells[11].Value.ToString().Replace(degisecek5, degismis5);

                                baglanti.Open();
                                OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set cep_tel=@p12 where id=@p1", baglanti);
                                güncelle.Parameters.AddWithValue("@p12", item.Cells[11].Value);
                                güncelle.Parameters.AddWithValue("@p1", textBox1.Text);
                                baglanti.Close();
                            }
                        }
                    } 
                }

                else
                {
                    MessageBox.Show("Lütfen Tek Bir Seçim Yapınız !", "Hatalı Seçim", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                MessageBox.Show("Seçmiş Olduğunuz Alanın Telefon Verileri Temizlenmiştir", "Temizleme Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Btn_acma();
                Listele_metod();
                progressBar1.Value = 0;
            }

            private void Mail_Sil_Click_1(object sender, EventArgs e) //Email Silme İşlemi
            {
                backgroundWorker3.RunWorkerAsync();
                Btn_kapatma();
            }
            private bool IsValidEmail(string email)// Email Kontrolü
            {
                string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
                Regex regex = new Regex(pattern);
                return regex.IsMatch(email);
        }//@"Provider = Microsoft.ACE.OLEDB.12.0; 
        //Data Source = " + excelYol + ";" +
                //"Extended Properties= 'Excel 12.0 Xml;HDR=YES'"

            private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e) //Email Silme İşlemi
            {
            if (textBox7 != null && textBox1 != null)
            {
                using (OleDbConnection baglanti = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; 
        Data Source = " + excelYol + ";" +
                "Extended Properties= 'Excel 12.0 Xml;HDR=YES'"))
                {
                    baglanti.Open();

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        progressBar1.Maximum = dataGridView1.Rows.Count + 1;
                        progressBar1.Value = row.Index;
                         
                        if (row.Cells[6].Value != null)
                        {
                            string email = row.Cells[6].Value.ToString();

                            if (!IsValidEmail(email))
                            {
                                row.Cells[6].Value = "";
                            }
                        }

                        if (row.Cells[6].Value == null || string.IsNullOrEmpty(row.Cells[6].Value.ToString()))
                        {
                            row.Cells[6].Value = DBNull.Value;
                        }

                        string updateQuery = "UPDATE [sayfa1$] SET email = ? WHERE id = ?";

                        using (OleDbCommand güncelle = new OleDbCommand(updateQuery, baglanti))
                        {
                            güncelle.Parameters.AddWithValue("@p6", row.Cells[6].Value);
                            güncelle.Parameters.AddWithValue("@p1", textBox1.Text);

                            güncelle.ExecuteNonQuery();
                        }
                    }

                    baglanti.Close();
                }
            }

            //// Tüm satırları döngüye aldırdım
            //foreach (DataGridViewRow row in dataGridView1.Rows)
            //{
            //    progressBar1.Maximum = dataGridView1.Rows.Count + 1;
            //    progressBar1.Value = row.Index;

            //    // Satırdaki hücrede değer varsa e-posta adresi olarak kabul et
            //    if (row.Cells[6].Value != null)
            //    {
            //        string email = row.Cells[6].Value.ToString();
            //        // E-posta adresi geçersizse hücreyi temizle (sil)
            //        if (!IsValidEmail(email))
            //        {
            //            //string bosdgr = "";
            //            //textBox7.Text = textBox7.Text.Replace(email, bosdgr);
            //            row.Cells[6].Value = "";
            //        }
            //    }

            //    // Güncelleme işlemi yapılır
            //    if (textBox7 != null && textBox1 != null)
            //    {
            //        baglanti.Open(); // external table is not in the expected format. HATASI !
            //        OleDbCommand güncelle = new OleDbCommand("update [sayfa1$] set email=@p6 where id=@p1", baglanti);

            //        //güncelle.Parameters.AddWithValue("@p6", row.Cells[6].Value);
            //        if (row.Cells[6].Value == null || string.IsNullOrEmpty(row.Cells[6].Value.ToString()))
            //        {
            //            güncelle.Parameters.AddWithValue("@p6", DBNull.Value);
            //        }
            //        else
            //        {
            //            güncelle.Parameters.AddWithValue("@p6", row.Cells[6].Value);
            //        }
            //        güncelle.Parameters.AddWithValue("@p1", textBox1.Text);
            //        güncelle.ExecuteNonQuery();
            //        baglanti.Close();

            //    }
            //}
        }
            private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                MessageBox.Show("Formata Uygun Olmayan Mailler Tablonuzdan Temizlenmiştir", "Mail Silme Mesajı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar1.Value = 0;
                Listele_metod();
                Btn_acma();
            }

            void Btn_kapatma() //Bir İşlem Dönerken Diğer İşlem Butonlarını Kapatıyor.
            {
                //Add_Picture.Enabled = false;
                Sil_Button.Enabled = false;
                TümKayıtlarınNoSil.Enabled = false;
                Mail_Sil.Enabled = false;
                Listele_Button.Enabled = false;
                Ekle_Button.Enabled = false;
                Güncelle_Button.Enabled = false;
                ÖzelSil_button.Enabled = false;
            }
            void Btn_acma() //Kapatılmış Butonları İşlem Sonunda Tekrardan Açıyor. 
            {
                //Add_Picture.Enabled = true;
                Sil_Button.Enabled = true;
                TümKayıtlarınNoSil.Enabled = true;
                Mail_Sil.Enabled = true;
                Listele_Button.Enabled = true;
                Ekle_Button.Enabled = true;
                Güncelle_Button.Enabled = true;
                ÖzelSil_button.Enabled = true;
            }

        private void pictureBox1_Click(object sender, EventArgs e)//Bilsoft Linki
        {
            System.Diagnostics.Process.Start("https://www.bilsoft.com");
        }

        // Placeholder Kısmı
        private void textBox9_Enter(object sender, EventArgs e) //Placeholder Kısmı TextBox9
        {
            if (textBox9.Text == "Örn; 541 327 1267")
            {
                textBox9.ForeColor = Color.FromArgb(0, 0, 192);
                textBox9.Text = "";
            }
        }
        private void textBox9_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox9.Text))
            {
                textBox9.ForeColor = SystemColors.GrayText;
                textBox9.Text = "Örn; 541 327 1267";
            }
        }//----------------------------------------------------------------------

        private void textBox10_Enter(object sender, EventArgs e)//Placeholder Kısmı TextBox10
        {
            if (textBox10.Text == "Örn; 541 327 1267")
            {
                textBox10.ForeColor = Color.FromArgb(0, 0, 192);
                textBox10.Text = "";
            }
        } 

        private void textBox10_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox10.Text))
            {
                textBox10.ForeColor = SystemColors.GrayText;
                textBox10.Text = "Örn; 541 327 1267";
            }
        }//----------------------------------------------------------------------

        private void textBox12_Enter(object sender, EventArgs e)//Placeholder Kısmı TextBox12
        {
            if (textBox12.Text == "Örn; 541 327 1267")
            {
                textBox12.ForeColor = Color.FromArgb(0, 0, 192);
                textBox12.Text = "";
            }
        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox12.Text))
            {
                textBox12.ForeColor = SystemColors.GrayText;
                textBox12.Text = "Örn; 541 327 1267";
            }
        }//----------------------------------------------------------------------

        private void Kopya_iconu_Click(object sender, EventArgs e)//Kopyalama İconu
        {
            if (!string.IsNullOrEmpty(textBox13.Text))
            {
                Clipboard.SetText(textBox13.Text);
                MessageBox.Show("Dosya yolu kopyalanmıştır.", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }  
        }

        private void Kopya_iconu_MouseEnter(object sender, EventArgs e)//Kopyala İconu Metin Görüntülenme
        {
            label15.Visible = true;//İcon Metni Gözükür
        }

        private void Kopya_iconu_MouseLeave(object sender, EventArgs e)//Kopyala İconu Metin Görüntülenme
        {
            label15.Visible = false;//İcon Metni Gözükmez
        }

    }
    }

