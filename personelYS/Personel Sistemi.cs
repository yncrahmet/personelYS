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
        //access veritabaný baðlantý dizesi
        string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=personelys.accdb;";
        public Form1()
        {
            InitializeComponent();
            FillComboBoxWithCities();

            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.None;

            gnclebtn.Visible = true; //form açýlýþýnda gnclebtn aktif, týklandýðýnda deaktif olacak
            gncleyp.Visible = false; //kayýt seçilip gnclebtn týklandýðýnda gncleyp aktif olacak

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void FillComboBoxWithCities()
        {
            string[] cities = {
                "Adana", "Adýyaman", "Afyon", "Aðrý", "Amasya", "Ankara", "Antalya", "Artvin", "Aydýn", "Balýkesir",
                "Bilecik", "Bingöl", "Bitlis", "Bolu", "Burdur", "Bursa", "Çanakkale", "Çankýrý", "Çorum", "Denizli",
                "Diyarbakýr", "Edirne", "Elazýð", "Erzincan", "Erzurum", "Eskiþehir", "Gaziantep", "Giresun", "Gümüþhane",
                "Hakkari", "Hatay", "Isparta", "Mersin", "Ýstanbul", "Ýzmir", "Kars", "Kastamonu", "Kayseri", "Kýrklareli",
                "Kýrþehir", "Kocaeli", "Konya", "Kütahya", "Malatya", "Manisa", "Kahramanmaraþ", "Mardin", "Muðla", "Muþ",
                "Nevþehir", "Niðde", "Ordu", "Rize", "Sakarya", "Samsun", "Siirt", "Sinop", "Sivas", "Tekirdað",
                "Tokat", "Trabzon", "Tunceli", "Þanlýurfa", "Uþak", "Van", "Yozgat", "Zonguldak", "Aksaray", "Bayburt",
                "Karaman", "Kýrýkkale", "Batman", "Þýrnak", "Bartýn", "Ardahan", "Iðdýr", "Yalova", "Karabük", "Kilis", "Osmaniye",
                "Düzce"
            };

            shrcbox.Items.AddRange(cities);
            //shrcbox combobox itemine yukarda yazýlan þehirleri listeler.
        }

        private void lstlebtn_Click(object sender, EventArgs e)
        {
            // fonksiyon her kayýt için gerçekleþtiriyor

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                //sql sorgusundan elde edilen veriler veritablosuna eklendi
                OleDbCommand command = new OleDbCommand("SELECT perTC AS TC, perAdi AS Adý, perSoyadi AS Soyadý, perCinsiyet AS Cinsiyeti, perSehir AS Þehir, perMaas AS Maaþý, perUnvan AS Unvaný, perIkramiye AS Ikramiyesi FROM personelBilgi", connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                // Veri tablosu dataGridView1 veri kaynaðý olarak ayarlandý
                dataGridView1.DataSource = dataTable;
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                connection.Close();

            }

        }
        List<string> secilenTCKayitlar = new List<string>(); // secilenTCKayitlar adýnda boþ bir string listesi oluþturuluyor.
        List<string> secilenAdSoyadKayitlar = new List<string>(); // secilenAdSoyadKayitlar adýnda boþ bir string listesi oluþturuluyor.
        List<string> secilenUnvanKayitlar = new List<string>(); // secilenUnvanKayitlar adýnda boþ bir string listesi oluþturuluyor.

        private void kydtbtn_Click(object sender, EventArgs e)
        {
            // boþ giriþ kontrolü yapýlýyor...
            if (string.IsNullOrWhiteSpace(tctb.Text) ||
            string.IsNullOrWhiteSpace(aditb.Text) ||
            string.IsNullOrWhiteSpace(saditb.Text) ||
            !erb.Checked && !krb.Checked ||
            shrcbox.SelectedItem == null ||
            string.IsNullOrWhiteSpace(maastb.Text) ||
            string.IsNullOrWhiteSpace(unvtb.Text))

            {
                // Hata mesajý gösterilir ve iþleme devam edilmiyor
                MessageBox.Show("Lütfen tüm alanlarý doldurun.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (tctb.Text.Length != 11)
            {
                // tc no 11 haneli deðilse hata mesajý veriyor.
                MessageBox.Show("TC kimlik numarasý 11 haneli olmalýdýr.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Veritabanýna ekleme sorgusu
                string insertQuery = "INSERT INTO personelBilgi (perTC, perAdi, perSoyadi, perCinsiyet, perSehir, perMaas, perUnvan) VALUES (@Alan1, @Alan2, @Alan3, @Alan4, @Alan5, @Alan6, @Alan7)";

                OleDbCommand command = new OleDbCommand(insertQuery, connection);

                // Parametreler atanýyor
                command.Parameters.AddWithValue("@Alan1", tctb.Text);
                command.Parameters.AddWithValue("@Alan2", aditb.Text);
                command.Parameters.AddWithValue("@Alan3", saditb.Text);
                command.Parameters.AddWithValue("@Alan4", GetCinsiyet());
                command.Parameters.AddWithValue("@Alan5", shrcbox.SelectedItem);
                command.Parameters.AddWithValue("@Alan6", maastb.Text);
                command.Parameters.AddWithValue("@Alan7", unvtb.Text);

                try
                {
                    // Komut çalýþtýrýlýyor ve kayýt ekleniyor
                    command.ExecuteNonQuery();
                    MessageBox.Show("Kayýt baþarýyla eklendi.");
                    lstlebtn_Click(sender, e); // datagridviewdeki verilerin listesini yenilemek için lstlebtn çaðýrýlýyor

                    //giriþ alanlarý temizleniyor
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
                    // ayný TC ile kayýtlý personel varsa hata mesajý gösteriliyor
                    if (ex.ErrorCode == -2147467259)
                    {
                        MessageBox.Show("Bu TC kimlik numarasýyla kayýtlý personel zaten mevcut.");
                    }
                    else
                    {
                        // Diðer hata durumlarýnda hata mesajý
                        MessageBox.Show("Kaydetme sýrasýnda bir hata oluþtu: " + ex.Message);
                    }
                }

                connection.Close();
            }
        }

        string GetCinsiyet()
        {
            //cinsiyet seçimi için konulan radiobuttonlara göre deðer döndürülüyor.
            if (erb.Checked)
                return "E";
            else if (krb.Checked)
                return "K";
            else
                return string.Empty;
        }

        private void silbtn_Click(object sender, EventArgs e)
        {
            // Geçerli bir kayýt seçilip seçilmediði kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("Lütfen geçerli bir kayýt seçin.");
                return;
            }

            // Silme iþlemi onayý dialog penceresi ile alýnýyor
            DialogResult result = MessageBox.Show("Seçili kayýtlarý silmek istiyor musunuz?", "Kayýt Silme", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        int selectedRowIndex = row.Index;
                        string tc = dataGridView1.Rows[selectedRowIndex].Cells["TC"].Value.ToString();

                        // Ýlgili personelin mesai kayýtlarý siliniyor
                        string deleteMesaiQuery = "DELETE FROM perMesai WHERE perTC = @TC";
                        OleDbCommand deleteMesaiCommand = new OleDbCommand(deleteMesaiQuery, connection);
                        deleteMesaiCommand.Parameters.AddWithValue("@TC", tc);
                        deleteMesaiCommand.ExecuteNonQuery();

                        // Ýlgili personelin genel kaydý siliniyor
                        string deleteQuery = "DELETE FROM personelBilgi WHERE perTC = @TC";
                        OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection);
                        deleteCommand.Parameters.AddWithValue("@TC", tc);
                        deleteCommand.ExecuteNonQuery();

                        // Seçili satýr DataGridView'den kaldýrýlýyor
                        dataGridView1.Rows.RemoveAt(selectedRowIndex);
                    }

                    connection.Close();
                }
                
                MessageBox.Show("Kayýtlar baþarýyla silindi.");
            }
        }

        private void tmzlebtn_Click(object sender, EventArgs e)
        {
            //datagridviewde listeli kayýtlar temizleniyor. (veritabanýndan deðil!)
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            // eðer giriþ alanlarý dolu ise orayý da temizlemektedir.
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
            // Geçerli bir kayýt seçilip seçilmediði kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("Lütfen geçerli bir kayýt seçin.");
                return;
            }

            // Seçilen satýrýn verileri alýnýyor ve ilgili alanlara yerleþtiriliyor
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                tctb.Text = selectedRow.Cells["TC"].Value.ToString();
                aditb.Text = selectedRow.Cells["Adý"].Value.ToString();
                saditb.Text = selectedRow.Cells["Soyadý"].Value.ToString();

                string cinsiyet = selectedRow.Cells["Cinsiyeti"].Value.ToString();
                if (cinsiyet == "E")
                    erb.Checked = true;
                else if (cinsiyet == "K")
                    krb.Checked = true;

                shrcbox.SelectedItem = selectedRow.Cells["Þehir"].Value.ToString();
                maastb.Text = selectedRow.Cells["Maaþý"].Value.ToString();
                unvtb.Text = selectedRow.Cells["Unvaný"].Value.ToString();

                // Güncelleme iþlemleri için ilgili butonlarýn görünürlüðü ayarlanýyor
                gnclebtn.Visible = false;
                gncleyp.Visible = true; //deðiþiklik yapýlýp bu butona týklanmasý için
            }
            else
            {
                // kayýt seçilmeden buton týklanýrsa verilen hata mesajý
                MessageBox.Show("Lütfen güncellenecek bir kayýt seçin.");
            }
        }

        private void gncleyp_Click(object sender, EventArgs e)
        {
            // Geçerli bir kayýt seçilip seçilmediði kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("Lütfen geçerli bir kayýt seçin.");
                return;
            }

            // Seçilen satýrýn verileri alýnýyor ve güncelleme iþlemi gerçekleþtiriliyor
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                string tc = selectedRow.Cells["TC"].Value.ToString();

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Personel kaydý güncelleniyor
                    string updateQueryPersonel = "UPDATE personelBilgi SET perTC = @Alan1, perAdi = @Alan2, perSoyadi = @Alan3, perCinsiyet = @Alan4, perSehir = @Alan5, perMaas = @Alan6, perUnvan = @Alan7 WHERE perTC = @TC";
                    // Mesai kayýtlarý güncelleniyor (bu eklenmezse iliiþkili alanlar olduðu için güncelleme yapýlamaz)
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
                            // sayý girili alana sayý girilmezse verilen hata mesajý
                            MessageBox.Show("Maaþ alaný için geçerli bir sayý giriniz.");
                            return;
                        }

                        command.Parameters.AddWithValue("@Alan7", unvtb.Text);
                        command.Parameters.AddWithValue("@TC", tc);

                        try
                        {
                            // Personel kaydý güncelleniyor
                            command.ExecuteNonQuery();
                            MessageBox.Show("Kayýt baþarýyla güncellendi.");
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
                            MessageBox.Show("Güncelleme sýrasýnda bir hata oluþtu: " + ex.Message);
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
                // herhangi bir kayýt seçilmemiþse bu mesaj geliyor
                MessageBox.Show("Lütfen güncellenecek bir kayýt seçin.");
            }
        }

        private void clsmasaatibtn_Click(object sender, EventArgs e)
        {
            // Geçerli bir kayýt seçilip seçilmediði kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("Lütfen geçerli bir kayýt seçin.");
                return;
            }
            // Kullanýcýdan çalýþma saati bilgisi ayrý bir giriþ kutusunda isteniyor
            string calismaSaati = Interaction.InputBox("Çalýþma Saati Girin:", "Mesai Giriþi", "");

            // Geçerli bir çalýþma saati girilip girilmediði kontrol ediliyor
            if (string.IsNullOrEmpty(calismaSaati))
            {
                MessageBox.Show("Lütfen geçerli bir çalýþma saati girin.");
                return;
            }

            // Bugünün tarihi alýnýyor
            DateTime tarih = DateTime.Now.Date;

            // Seçili satýrlarýn üzerinde dönülerek mesai giriþi yapýlýyor
            foreach (DataGridViewRow selectedRow in dataGridView1.SelectedRows)
            {
                string selectedPerTC = selectedRow.Cells["TC"].Value.ToString();

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Ayný kiþi ve tarih için daha önce mesai giriþi yapýlmýþ mý kontrol ediliyor
                    string query = "SELECT COUNT(*) FROM perMesai WHERE perTC = @perTC AND tarih = @tarih";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@perTC", selectedPerTC);
                        command.Parameters.AddWithValue("@tarih", tarih);

                        int existingRecordsCount = (int)command.ExecuteScalar();

                        if (existingRecordsCount > 0)
                        {
                            MessageBox.Show("Bu kiþiye ayný gün içinde birden fazla mesai giriþi yapýlamaz!");
                            continue;
                        }
                    }

                    // Mesai giriþi veritabanýna kaydediliyor
                    query = "INSERT INTO perMesai (perTC, tarih, calismaSaati) VALUES (@perTC, @tarih, @calismaSaati)";
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@perTC", selectedPerTC);
                        command.Parameters.AddWithValue("@tarih", tarih);
                        command.Parameters.AddWithValue("@calismaSaati", calismaSaati);

                        command.ExecuteNonQuery();

                        MessageBox.Show("Mesai giriþi baþarýyla kaydedildi.");
                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //button1 form arayüzünde sað en üstte olan çarpý iþaretidir. týklanýrsa uygulama kapatýlýr
            Application.Exit();
        }

        private void ikramiyeHsplaBtn_Click(object sender, EventArgs e)
        {
            // Geçerli bir kayýt seçilip seçilmediði kontrol ediliyor
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows[0].IsNewRow)
            {
                MessageBox.Show("Lütfen geçerli bir kayýt seçin.");
                return;
            }

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Seçilen satýrlar üzerinde dönülerek ikramiye hesaplamasý yapýlýyor
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        string perTC = row.Cells["TC"].Value.ToString();

                        int toplamCalismaSaati = 0;
                        int mesaiGirilenGunSayisi = 0;

                        // Seçilen personele ait mesai kayýtlarý alýnýyor
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
                        // Seçilen personele ait kiþisel bilgiler alýnýyor
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

                        // Ýkramiye zaten hesaplanmýþsa hata mesajý gösterilip diðer kayýtlara geçiliyor
                        if (perIkramiye > 0)
                        {
                            MessageBox.Show("Hata: " + perTC + " - " + perAdi + " " + perSoyadi + " için ikramiye zaten hesaplandý!");
                            continue;
                        }

                        // 30 günden az mesai girilmiþse hata mesajý gösterilip diðer kayýtlara geçiliyor
                        if (mesaiGirilenGunSayisi < 30)
                        {
                            MessageBox.Show("Hata: " + perTC + " - " + perAdi + " " + perSoyadi + " için 30 günlük mesai girilmemiþtir!");
                            continue;
                        }

                        double ikramiye = 0;
                        
                        // günlük min 8 saat dersek aylýk minimum 240 saat
                        // Toplam çalýþma saati bazýnda ikramiye hesaplanýyor
                        if (toplamCalismaSaati > 300) //300 e ulaþanlara 6000 ikramiye
                        {
                            ikramiye = 6000;
                        }
                        else if (toplamCalismaSaati > 250) //250 ve 300 arasýnda olanlara 4000 ikramiye
                        {
                            ikramiye = 4000;
                        }
                        // Ýkramiye hesaplanmýþsa, veritabanýnda güncelleme yapýlýyor ve bilgi mesajý gösteriliyor
                        if (ikramiye > 0)
                        {
                            string updateQuery = "UPDATE personelBilgi SET perIkramiye = @ikramiye WHERE perTC = @perTC";
                            using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                            {
                                updateCommand.Parameters.AddWithValue("@ikramiye", ikramiye);
                                updateCommand.Parameters.AddWithValue("@perTC", perTC);
                                updateCommand.ExecuteNonQuery();
                            }

                            MessageBox.Show(perTC + " - " + perAdi + " " + perSoyadi + " için ikramiye hesaplandý ve Personel bilgisine eklendi!");
                        }
                    }
                    // Sonuçlarý görüntülemek için listeleme iþlemi yeniden yapýlýyor
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

                // Toplam personel sayýsý istatistiði
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

                // Þehir istatistikleri
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

                // Ýstatistikleri göster
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"Toplam Personel Sayýsý: {toplamPersonel}");
                sb.AppendLine();
                sb.AppendLine("Cinsiyet Ýstatistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(cinsiyetIstatistikleri));
                sb.AppendLine();
                sb.AppendLine("Þehir Ýstatistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(sehirIstatistikleri));
                sb.AppendLine();
                sb.AppendLine("Unvan Ýstatistikleri:");
                sb.AppendLine(GetFormattedIstatistikler(unvanIstatistikleri));

                MessageBox.Show(sb.ToString(), "Ýstatistikler");
            }
        }

        private string GetFormattedIstatistikler(Dictionary<string, int> istatistikler)
        {
            StringBuilder sb = new StringBuilder();
            // Ýstatistikler sözlüðündeki her bir çift için döngü oluþturulur
            foreach (var entry in istatistikler)
            {
                sb.AppendLine($"{entry.Key}: {entry.Value}");
            }
            return sb.ToString();

        }
    }
}