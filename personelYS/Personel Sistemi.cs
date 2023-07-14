using Microsoft.VisualBasic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Text;

namespace personelYS
{
    public partial class Form1 : Form
    {
        //access veritaban� ba�lant� dizesi
        string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=personelys.accdb;";
        public Form1()
        {
            InitializeComponent();
            FillComboBoxWithCities();

            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.None;

            gnclebtn.Visible = true; //form a��l���nda gnclebtn aktif, t�kland���nda deaktif olacak
            gncleyp.Visible = false; //kay�t se�ilip gnclebtn t�kland���nda gncleyp aktif olacak

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void FillComboBoxWithCities()
        {
            string[] cities = {
                "Adana", "Ad�yaman", "Afyon", "A�r�", "Amasya", "Ankara", "Antalya", "Artvin", "Ayd�n", "Bal�kesir",
                "Bilecik", "Bing�l", "Bitlis", "Bolu", "Burdur", "Bursa", "�anakkale", "�ank�r�", "�orum", "Denizli",
                "Diyarbak�r", "Edirne", "Elaz��", "Erzincan", "Erzurum", "Eski�ehir", "Gaziantep", "Giresun", "G�m��hane",
                "Hakkari", "Hatay", "Isparta", "Mersin", "�stanbul", "�zmir", "Kars", "Kastamonu", "Kayseri", "K�rklareli",
                "K�r�ehir", "Kocaeli", "Konya", "K�tahya", "Malatya", "Manisa", "Kahramanmara�", "Mardin", "Mu�la", "Mu�",
                "Nev�ehir", "Ni�de", "Ordu", "Rize", "Sakarya", "Samsun", "Siirt", "Sinop", "Sivas", "Tekirda�",
                "Tokat", "Trabzon", "Tunceli", "�anl�urfa", "U�ak", "Van", "Yozgat", "Zonguldak", "Aksaray", "Bayburt",
                "Karaman", "K�r�kkale", "Batman", "��rnak", "Bart�n", "Ardahan", "I�d�r", "Yalova", "Karab�k", "Kilis", "Osmaniye",
                "D�zce"
            };

            shrcbox.Items.AddRange(cities);
            //shrcbox combobox itemine yukarda yaz�lan �ehirleri listeler.
        }

        private void lstlebtn_Click(object sender, EventArgs e)
        {
            // fonksiyon her kay�t i�in ger�ekle�tiriyor

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                //sql sorgusundan elde edilen veriler veritablosuna eklendi
                OleDbCommand command = new OleDbCommand("SELECT perTC AS TC, perAdi AS Ad�, perSoyadi AS Soyad�, perCinsiyet AS Cinsiyeti, perSehir AS �ehir, perMaas AS Maa��, perUnvan AS Unvan�, perIkramiye AS Ikramiyesi FROM personelBilgi", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Veri tablosu dataGridView1 veri kayna�� olarak ayarland�
                dataGridView1.DataSource = dataTable;
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                connection.Close();

            }

        }
        List<string> secilenTCKayitlar = new List<string>(); // secilenTCKayitlar ad�nda bo� bir string listesi olu�turuluyor.
        List<string> secilenAdSoyadKayitlar = new List<string>(); // secilenAdSoyadKayitlar ad�nda bo� bir string listesi olu�turuluyor.
        List<string> secilenUnvanKayitlar = new List<string>(); // secilenUnvanKayitlar ad�nda bo� bir string listesi olu�turuluyor.

        private void kydtbtn_Click(object sender, EventArgs e)
        {
            // bo� giri� kontrol� yap�l�yor...
            if (string.IsNullOrWhiteSpace(tctb.Text) ||
            string.IsNullOrWhiteSpace(aditb.Text) ||
            string.IsNullOrWhiteSpace(saditb.Text) ||
            !erb.Checked && !krb.Checked ||
            shrcbox.SelectedItem == null ||
            string.IsNullOrWhiteSpace(maastb.Text) ||
            string.IsNullOrWhiteSpace(unvtb.Text))

            {
                // Hata mesaj� g�sterilir ve i�leme devam edilmiyor
                MessageBox.Show("L�tfen t�m alanlar� doldurun.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (tctb.Text.Length != 11)
            {
                // tc no 11 haneli de�ilse hata mesaj� veriyor.
                MessageBox.Show("TC kimlik numaras� 11 haneli olmal�d�r.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Veritaban�na ekleme sorgusu
                string insertQuery = "INSERT INTO personelBilgi (perTC, perAdi, perSoyadi, perCinsiyet, perSehir, perMaas, perUnvan) VALUES (@Alan1, @Alan2, @Alan3, @Alan4, @Alan5, @Alan6, @Alan7)";

                OleDbCommand command = new OleDbCommand(insertQuery, connection);

                // Parametreler atan�yor
                command.Parameters.AddWithValue("@Alan1", tctb.Text);
                command.Parameters.AddWithValue("@Alan2", aditb.Text);
                command.Parameters.AddWithValue("@Alan3", saditb.Text);
                command.Parameters.AddWithValue("@Alan4", GetCinsiyet());
                command.Parameters.AddWithValue("@Alan5", shrcbox.SelectedItem);
                command.Parameters.AddWithValue("@Alan6", maastb.Text);
                command.Parameters.AddWithValue("@Alan7", unvtb.Text);

                try
                {
                    // Komut �al��t�r�l�yor ve kay�t ekleniyor
                    command.ExecuteNonQuery();
                    MessageBox.Show("Kay�t ba�ar�yla eklendi.");
                    lstlebtn_Click(sender, e); // datagridviewdeki verilerin listesini yenilemek i�in lstlebtn �a��r�l�yor

                    //giri� alanlar� temizleniyor
                    tctb.Text = "";
                    aditb.Text = "";
                    saditb.Text = "";
                    erb.Checked = false;
                    krb.Checked = false;
                    shrcbox.SelectedItem = null;
                    maastb.Text = "";
                    unvtb.Text = "";
                }
                catch (OleDbException ex)
                {
                    // ayn� TC ile kay�tl� personel varsa hata mesaj� g�steriliyor
                    if (ex.ErrorCode == -2147467259)
                    {
                        MessageBox.Show("Bu TC kimlik numaras�yla kay�tl� personel zaten mevcut.");
                    }
                    else
                    {
                        // Di�er hata durumlar�nda hata mesaj�
                        MessageBox.Show("Kaydetme s�ras�nda bir hata olu�tu: " + ex.Message);
                    }
                }

                connection.Close();
            }
        }

        string GetCinsiyet()
        {
            //cinsiyet se�imi i�in konulan radiobuttonlara g�re de�er d�nd�r�l�yor.
            if (erb.Checked)
                return "E";
            else if (krb.Checked)
                return "K";
            else
                return string.Empty;
        }

        private void silbtn_Click(object sender, EventArgs e)
        {
            // Ge�erli bir kay�t se�ilip se�ilmedi�i kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("L�tfen ge�erli bir kay�t se�in.");
                return;
            }

            // Silme i�lemi onay� dialog penceresi ile al�n�yor
            DialogResult result = MessageBox.Show("Se�ili kay�tlar� silmek istiyor musunuz?", "Kay�t Silme", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        int selectedRowIndex = row.Index;
                        string tc = dataGridView1.Rows[selectedRowIndex].Cells["TC"].Value.ToString();

                        // �lgili personelin mesai kay�tlar� siliniyor
                        string deleteMesaiQuery = "DELETE FROM perMesai WHERE perTC = @TC";
                        OleDbCommand deleteMesaiCommand = new OleDbCommand(deleteMesaiQuery, connection);
                        deleteMesaiCommand.Parameters.AddWithValue("@TC", tc);
                        deleteMesaiCommand.ExecuteNonQuery();

                        // �lgili personelin genel kayd� siliniyor
                        string deleteQuery = "DELETE FROM personelBilgi WHERE perTC = @TC";
                        OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection);
                        deleteCommand.Parameters.AddWithValue("@TC", tc);
                        deleteCommand.ExecuteNonQuery();

                        // Se�ili sat�r DataGridView'den kald�r�l�yor
                        dataGridView1.Rows.RemoveAt(selectedRowIndex);
                    }

                    connection.Close();
                }
                
                MessageBox.Show("Kay�tlar ba�ar�yla silindi.");
            }
        }

        private void tmzlebtn_Click(object sender, EventArgs e)
        {
            //datagridviewde listeli kay�tlar temizleniyor. (veritaban�ndan de�il!)
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            // e�er giri� alanlar� dolu ise oray� da temizlemektedir.
            tctb.Text = "";
            aditb.Text = "";
            saditb.Text = "";
            erb.Checked = false;
            krb.Checked = false;
            shrcbox.SelectedItem = null;
            maastb.Text = "";
            unvtb.Text = "";
        }

        private void gnclebtn_Click(object sender, EventArgs e)
        {
            // Ge�erli bir kay�t se�ilip se�ilmedi�i kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("L�tfen ge�erli bir kay�t se�in.");
                return;
            }

            // Se�ilen sat�r�n verileri al�n�yor ve ilgili alanlara yerle�tiriliyor
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                tctb.Text = selectedRow.Cells["TC"].Value.ToString();
                aditb.Text = selectedRow.Cells["Ad�"].Value.ToString();
                saditb.Text = selectedRow.Cells["Soyad�"].Value.ToString();

                string cinsiyet = selectedRow.Cells["Cinsiyeti"].Value.ToString();
                if (cinsiyet == "E")
                    erb.Checked = true;
                else if (cinsiyet == "K")
                    krb.Checked = true;

                shrcbox.SelectedItem = selectedRow.Cells["�ehir"].Value.ToString();
                maastb.Text = selectedRow.Cells["Maa��"].Value.ToString();
                unvtb.Text = selectedRow.Cells["Unvan�"].Value.ToString();

                // G�ncelleme i�lemleri i�in ilgili butonlar�n g�r�n�rl��� ayarlan�yor
                gnclebtn.Visible = false;
                gncleyp.Visible = true; //de�i�iklik yap�l�p bu butona t�klanmas� i�in
            }
            else
            {
                // kay�t se�ilmeden buton t�klan�rsa verilen hata mesaj�
                MessageBox.Show("L�tfen g�ncellenecek bir kay�t se�in.");
            }
        }

        private void gncleyp_Click(object sender, EventArgs e)
        {
            // Ge�erli bir kay�t se�ilip se�ilmedi�i kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("L�tfen ge�erli bir kay�t se�in.");
                return;
            }

            // Se�ilen sat�r�n verileri al�n�yor ve g�ncelleme i�lemi ger�ekle�tiriliyor
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                string tc = selectedRow.Cells["TC"].Value.ToString();

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Personel kayd� g�ncelleniyor
                    string updateQueryPersonel = "UPDATE personelBilgi SET perTC = @Alan1, perAdi = @Alan2, perSoyadi = @Alan3, perCinsiyet = @Alan4, perSehir = @Alan5, perMaas = @Alan6, perUnvan = @Alan7 WHERE perTC = @TC";
                    // Mesai kay�tlar� g�ncelleniyor (bu eklenmezse ilii�kili alanlar oldu�u i�in g�ncelleme yap�lamaz)
                    string updateQueryMesai = "UPDATE perMesai SET perTC = @Alan1 WHERE perTC = @TC";

                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = updateQueryPersonel;
                        command.Parameters.AddWithValue("@Alan1", tctb.Text);
                        command.Parameters.AddWithValue("@Alan2", aditb.Text);
                        command.Parameters.AddWithValue("@Alan3", saditb.Text);
                        string cinsiyetValue = (erb.Checked) ? "E" : "K";
                        command.Parameters.AddWithValue("@Alan4", cinsiyetValue);
                        command.Parameters.AddWithValue("@Alan5", shrcbox.SelectedItem);

                        decimal maas;
                        if (decimal.TryParse(maastb.Text, out maas))
                        {
                            command.Parameters.AddWithValue("@Alan6", maas);
                        }
                        else
                        {
                            // say� girili alana say� girilmezse verilen hata mesaj�
                            MessageBox.Show("Maa� alan� i�in ge�erli bir say� giriniz.");
                            return;
                        }

                        command.Parameters.AddWithValue("@Alan7", unvtb.Text);
                        command.Parameters.AddWithValue("@TC", tc);

                        try
                        {
                            // Personel kayd� g�ncelleniyor
                            command.ExecuteNonQuery();
                            MessageBox.Show("Kay�t ba�ar�yla g�ncellendi.");
                            lstlebtn_Click(sender, e);

                            using (OleDbCommand command2 = new OleDbCommand())
                            {
                                command2.Connection = connection;
                                command2.CommandText = updateQueryMesai;
                                command2.Parameters.AddWithValue("@Alan1", tctb.Text);
                                command2.Parameters.AddWithValue("@TC", tc);
                                command2.ExecuteNonQuery();
                            }
                        }
                        catch (OleDbException ex)
                        {
                            MessageBox.Show("G�ncelleme s�ras�nda bir hata olu�tu: " + ex.Message);
                        }
                    }

                    connection.Close();
                }

                // Alanlar temizleniyor
                tctb.Text = "";
                aditb.Text = "";
                saditb.Text = "";
                erb.Checked = false;
                krb.Checked = false;
                shrcbox.SelectedItem = null;
                maastb.Text = "";
                unvtb.Text = "";

                // Butonlar eski haline getiriliyor
                gnclebtn.Visible = true;
                gncleyp.Visible = false;
            }
            else
            {
                // herhangi bir kay�t se�ilmemi�se bu mesaj geliyor
                MessageBox.Show("L�tfen g�ncellenecek bir kay�t se�in.");
            }
        }

        private void clsmasaatibtn_Click(object sender, EventArgs e)
        {
            // Ge�erli bir kay�t se�ilip se�ilmedi�i kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("L�tfen ge�erli bir kay�t se�in.");
                return;
            }
            // Kullan�c�dan �al��ma saati bilgisi ayr� bir giri� kutusunda isteniyor
            string calismaSaati = Interaction.InputBox("�al��ma Saati Girin:", "Mesai Giri�i", "");

            // Ge�erli bir �al��ma saati girilip girilmedi�i kontrol ediliyor
            if (string.IsNullOrEmpty(calismaSaati))
            {
                MessageBox.Show("L�tfen ge�erli bir �al��ma saati girin.");
                return;
            }

            // Bug�n�n tarihi al�n�yor
            DateTime tarih = DateTime.Now.Date;

            // Se�ili sat�rlar�n �zerinde d�n�lerek mesai giri�i yap�l�yor
            foreach (DataGridViewRow selectedRow in dataGridView1.SelectedRows)
            {
                string selectedPerTC = selectedRow.Cells["TC"].Value.ToString();

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Ayn� ki�i ve tarih i�in daha �nce mesai giri�i yap�lm�� m� kontrol ediliyor
                    string query = "SELECT COUNT(*) FROM perMesai WHERE perTC = @perTC AND tarih = @tarih";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@perTC", selectedPerTC);
                        command.Parameters.AddWithValue("@tarih", tarih);

                        int existingRecordsCount = (int)command.ExecuteScalar();

                        if (existingRecordsCount > 0)
                        {
                            MessageBox.Show("Bu ki�iye ayn� g�n i�inde birden fazla mesai giri�i yap�lamaz!");
                            continue;
                        }
                    }

                    // Mesai giri�i veritaban�na kaydediliyor
                    query = "INSERT INTO perMesai (perTC, tarih, calismaSaati) VALUES (@perTC, @tarih, @calismaSaati)";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@perTC", selectedPerTC);
                        command.Parameters.AddWithValue("@tarih", tarih);
                        command.Parameters.AddWithValue("@calismaSaati", calismaSaati);

                        command.ExecuteNonQuery();

                        MessageBox.Show("Mesai giri�i ba�ar�yla kaydedildi.");
                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //button1 form aray�z�nde sa� en �stte olan �arp� i�aretidir. t�klan�rsa uygulama kapat�l�r
            Application.Exit();
        }

        private void ikramiyeHsplaBtn_Click(object sender, EventArgs e)
        {
            // Ge�erli bir kay�t se�ilip se�ilmedi�i kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("L�tfen ge�erli bir kay�t se�in.");
                return;
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Se�ilen sat�rlar �zerinde d�n�lerek ikramiye hesaplamas� yap�l�yor
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        string perTC = row.Cells["TC"].Value.ToString();

                        int toplamCalismaSaati = 0;
                        int mesaiGirilenGunSayisi = 0;

                        // Se�ilen personele ait mesai kay�tlar� al�n�yor
                        string perMesaiQuery = "SELECT calismaSaati FROM perMesai WHERE perTC = @perTC";
                        using (OleDbCommand perMesaiCommand = new OleDbCommand(perMesaiQuery, connection))
                        {
                            perMesaiCommand.Parameters.AddWithValue("@perTC", perTC);
                            using (OleDbDataReader perMesaiReader = perMesaiCommand.ExecuteReader())
                            {
                                while (perMesaiReader.Read())
                                {
                                    int calismaSaati;
                                    if (int.TryParse(perMesaiReader["calismaSaati"].ToString(), out calismaSaati))
                                    {
                                        toplamCalismaSaati += calismaSaati;
                                        mesaiGirilenGunSayisi++;
                                    }
                                }
                            }
                        }

                        string perAdi = "";
                        string perSoyadi = "";
                        double perIkramiye = 0;
                        // Se�ilen personele ait ki�isel bilgiler al�n�yor
                        string personelBilgiQuery = "SELECT perAdi, perSoyadi, perIkramiye FROM personelBilgi WHERE perTC = @perTC";
                        using (OleDbCommand personelBilgiCommand = new OleDbCommand(personelBilgiQuery, connection))
                        {
                            personelBilgiCommand.Parameters.AddWithValue("@perTC", perTC);
                            using (OleDbDataReader personelBilgiReader = personelBilgiCommand.ExecuteReader())
                            {
                                if (personelBilgiReader.Read())
                                {
                                    perAdi = personelBilgiReader["perAdi"].ToString();
                                    perSoyadi = personelBilgiReader["perSoyadi"].ToString();
                                    double.TryParse(personelBilgiReader["perIkramiye"].ToString(), out perIkramiye);
                                }
                            }
                        }

                        // �kramiye zaten hesaplanm��sa hata mesaj� g�sterilip di�er kay�tlara ge�iliyor
                        if (perIkramiye > 0)
                        {
                            MessageBox.Show("Hata: " + perTC + " - " + perAdi + " " + perSoyadi + " i�in ikramiye zaten hesapland�!");
                            continue;
                        }

                        // 30 g�nden az mesai girilmi�se hata mesaj� g�sterilip di�er kay�tlara ge�iliyor
                        if (mesaiGirilenGunSayisi < 30)
                        {
                            MessageBox.Show("Hata: " + perTC + " - " + perAdi + " " + perSoyadi + " i�in 30 g�nl�k mesai girilmemi�tir!");
                            continue;
                        }

                        double ikramiye = 0;
                        
                        // g�nl�k min 8 saat dersek ayl�k minimum 240 saat
                        // Toplam �al��ma saati baz�nda ikramiye hesaplan�yor
                        if (toplamCalismaSaati > 300) //300 e ula�anlara 6000 ikramiye
                        {
                            ikramiye = 6000;
                        }
                        else if (toplamCalismaSaati > 250) //250 ve 300 aras�nda olanlara 4000 ikramiye
                        {
                            ikramiye = 4000;
                        }
                        // �kramiye hesaplanm��sa, veritaban�nda g�ncelleme yap�l�yor ve bilgi mesaj� g�steriliyor
                        if (ikramiye > 0)
                        {
                            string updateQuery = "UPDATE personelBilgi SET perIkramiye = @ikramiye WHERE perTC = @perTC";
                            using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                            {
                                updateCommand.Parameters.AddWithValue("@ikramiye", ikramiye);
                                updateCommand.Parameters.AddWithValue("@perTC", perTC);
                                updateCommand.ExecuteNonQuery();
                            }

                            MessageBox.Show(perTC + " - " + perAdi + " " + perSoyadi + " i�in ikramiye hesapland� ve Personel bilgisine eklendi!");
                        }
                    }
                    // Sonu�lar� g�r�nt�lemek i�in listeleme i�lemi yeniden yap�l�yor
                    lstlebtn_Click(sender, e);
                }

                connection.Close();
            }

        }

        private void iststkbtn_Click(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Toplam personel say�s� istatisti�i
                int toplamPersonel = 0;
                string toplamPersonelQuery = "SELECT COUNT(*) FROM personelBilgi";
                using (OleDbCommand toplamPersonelCommand = new OleDbCommand(toplamPersonelQuery, connection))
                {
                    toplamPersonel = (int)toplamPersonelCommand.ExecuteScalar();
                }

                // Cinsiyet istatistikleri
                Dictionary<string, int> cinsiyetIstatistikleri = new Dictionary<string, int>();
                string cinsiyetIstatistikleriQuery = "SELECT perCinsiyet, COUNT(*) FROM personelBilgi GROUP BY perCinsiyet";
                using (OleDbCommand cinsiyetIstatistikleriCommand = new OleDbCommand(cinsiyetIstatistikleriQuery, connection))
                {
                    using (OleDbDataReader reader = cinsiyetIstatistikleriCommand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string cinsiyet = reader.GetString(0);
                            int count = reader.GetInt32(1);
                            cinsiyetIstatistikleri.Add(cinsiyet, count);
                        }
                    }
                }

                // �ehir istatistikleri
                Dictionary<string, int> sehirIstatistikleri = new Dictionary<string, int>();
                string sehirIstatistikleriQuery = "SELECT perSehir, COUNT(*) FROM personelBilgi GROUP BY perSehir";
                using (OleDbCommand sehirIstatistikleriCommand = new OleDbCommand(sehirIstatistikleriQuery, connection))
                {
                    using (OleDbDataReader reader = sehirIstatistikleriCommand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string sehir = reader.GetString(0);
                            int count = reader.GetInt32(1);
                            sehirIstatistikleri.Add(sehir, count);
                        }
                    }
                }

                // Unvan istatistikleri
                Dictionary<string, int> unvanIstatistikleri = new Dictionary<string, int>();
                string unvanIstatistikleriQuery = "SELECT perUnvan, COUNT(*) FROM personelBilgi GROUP BY perUnvan";
                using (OleDbCommand unvanIstatistikleriCommand = new OleDbCommand(unvanIstatistikleriQuery, connection))
                {
                    using (OleDbDataReader reader = unvanIstatistikleriCommand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string unvan = reader.GetString(0);
                            int count = reader.GetInt32(1);
                            unvanIstatistikleri.Add(unvan, count);
                        }
                    }
                }

                connection.Close();

                // �statistikleri g�ster
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"Toplam Personel Say�s�: {toplamPersonel}");
                sb.AppendLine();
                sb.AppendLine("Cinsiyet �statistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(cinsiyetIstatistikleri));
                sb.AppendLine();
                sb.AppendLine("�ehir �statistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(sehirIstatistikleri));
                sb.AppendLine();
                sb.AppendLine("Unvan �statistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(unvanIstatistikleri));

                MessageBox.Show(sb.ToString(), "�statistikler");
            }
        }

        private string GetFormattedIstatistikler(Dictionary<string, int> istatistikler)
        {
            StringBuilder sb = new StringBuilder();
            // �statistikler s�zl���ndeki her bir �ift i�in d�ng� olu�turulur
            foreach (var entry in istatistikler)
            {
                sb.AppendLine($"{entry.Key}: {entry.Value}");
            }
            return sb.ToString();

        }
    }
}