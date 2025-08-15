using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using gizliMail.gizliDb; // Bilgileri Değiştir

//KENDİME NOT:
//YORUM SATIRLARINDA BELİRTİLEN GİZLİ BİLGİLERİ DEĞİŞTİR 

//PACKAGE YÖNETİCİSİNDEN GEREKEN KÜTÜPHANELERİ YÜKLE

//Scaffold-DbContext "Server=XX.X.X.XXX\GİZLİSQLSERVER;Database= GizliDbAdı;User Id=UserIDAdı;Password= SifreAdı; TrustServerCertificate=True " Microsoft.EntityFrameworkCore.SqlServer -OutputDir GizliDb -Force

internal class Program
{
    private static async Task Main(string[] args)
    {
        using var db = new GizliDbContext(); //Bilgileri Değiştir

        var rows = FetchUygunsuzlukRowsWithNames(db);
        Console.WriteLine($"DB'den çekilen kayıt: {rows.Count}");

        string excelPath = ExportUygunsuzluklarToExcel(rows);
        Console.WriteLine($"Excel oluşturuldu: {excelPath}");

        var myEmail = "gizli@gizli.com"; //Bilgileri Değiştir
        await SendEmailWithAttachmentAsync(
            recipientEmails: new List<string> { myEmail },
            subject: $"Uygunsuzluklar Raporu - {DateTime.Now:yyyy-MM-dd}",
            htmlBody: BuildHtmlBodyForReport(),
            attachmentFilePaths: new List<string> { excelPath }
        );

        Console.WriteLine("Mail gönderildi.");
    }
    public class UygunsuzlukExcelRow
    {
        public DateTime? Tarih { get; set; }
        public string? UygunsuzluguYapanPersonel { get; set; }
        public string? UygunsuzlukAnaNeden { get; set; }
        public string? UygunsuzlukTehlike { get; set; }
        public string? UygunsuzlukAciklama { get; set; }
        public string? AlinmasiGerekenOnlemler { get; set; }
        public string? SorumluFirma { get; set; }
        public string? TespitYapan { get; set; }
        public DateTime? HedefTarihi { get; set; }
        public string? Sorumlular { get; set; }
        public string? UygunsuzlukKapatan { get; set; }
        public DateTime? UygunsuzlukKapatmaZamani { get; set; }
        public string? YapilanUygulama { get; set; }
        public bool? OnayaGonderildiMi { get; set; }
        public string? SonucuOnaylayan { get; set; }
        public DateTime? SonucuKapatmaZamani { get; set; }
        public bool? Sonuc { get; set; }
        public string? Grup { get; set; }
    }

    private static List<UygunsuzlukExcelRow> FetchUygunsuzlukRowsWithNames(GizliContext db) //Bilgileri Değiştir
    {
        var list = db.Uygunsuzluklars
            .Where(x => x.OnayaGonderildiMi == false)
            .Include(x => x.UygunsuzluguYapanPersonel)
            .Include(x => x.UygunsuzlukAnaNeden)
            .Include(x => x.UygunsuzlukTehlike)
            .Include(x => x.SorumluFirma)
            .Include(x => x.TespitYapanUser)
            .Include(x => x.UygunsuzlukKapatanUser)
            .Include(x => x.SonucuOnaylayanUser)
            .AsNoTracking()
            .ToList();

        return list.Select(x => new UygunsuzlukExcelRow
        {
            Tarih = x.Tarih,
            UygunsuzluguYapanPersonel = GetPersonelName(x.UygunsuzluguYapanPersonel),
            UygunsuzlukAnaNeden = GetDisplayName(x.UygunsuzlukAnaNeden),
            UygunsuzlukTehlike = GetDisplayName(x.UygunsuzlukTehlike),
            UygunsuzlukAciklama = x.UygunsuzlukAciklama,
            AlinmasiGerekenOnlemler = x.AlinmasiGerekenOnlemler,
            SorumluFirma = GetDisplayName(x.SorumluFirma),
            TespitYapan = GetDisplayName(x.TespitYapanUser),
            HedefTarihi = x.HedefTarihi,
            Sorumlular = x.Sorumlular,
            UygunsuzlukKapatan = GetDisplayName(x.UygunsuzlukKapatanUser),
            UygunsuzlukKapatmaZamani = GetProperty<DateTime?>(x, "UygunsuzlukKapatmaZamani", "UygungsuzlukKapatmaZamani"),
            YapilanUygulama = x.YapilanUygulama,
            OnayaGonderildiMi = x.OnayaGonderildiMi,
            SonucuOnaylayan = GetDisplayName(x.SonucuOnaylayanUser),
            SonucuKapatmaZamani = x.SonucuKapatmaZamani,
            Sonuc = x.Sonuc,
            Grup = x.Grup
        }).ToList();
    }



    private static string? GetDisplayName(object? entity)
    {
        if (entity == null) return null;
        var t = entity.GetType();


        var adi = t.GetProperty("Adi", BindingFlags.Public | BindingFlags.Instance)?.GetValue(entity)?.ToString();
        var soyadi = t.GetProperty("Soyadi", BindingFlags.Public | BindingFlags.Instance)?.GetValue(entity)?.ToString();
        if (!string.IsNullOrWhiteSpace(adi) || !string.IsNullOrWhiteSpace(soyadi))
            return $"{adi} {soyadi}".Trim();

        string[] candidates =
        {
            "AdSoyad","AdıSoyadı","AdiSoyadi","FullName","DisplayName","Name","CompanyName","Unvan","Title"
        };
        foreach (var name in candidates)
        {
            var p = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (p != null)
            {
                var v = p.GetValue(entity);
                if (v != null) return v.ToString();
            }
        }
        return entity.ToString();
    }

    private static T? GetProperty<T>(object obj, params string[] names)
    {
        var t = obj.GetType();
        foreach (var n in names)
        {
            var p = t.GetProperty(n, BindingFlags.Public | BindingFlags.Instance);
            if (p != null && typeof(T).IsAssignableFrom((Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType)))
            {
                return (T?)p.GetValue(obj);
            }
        }
        return default;
    }

    private static string ExportUygunsuzluklarToExcel(List<UygunsuzlukExcelRow> rows)
    {
        string fileName = $"Uygunsuzluklar_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        string fullPath = Path.Combine(Path.GetTempPath(), fileName);

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Uygunsuzluklar");

        var props = typeof(UygunsuzlukExcelRow)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .ToList();

        for (int c = 0; c < props.Count; c++)
            ws.Cell(1, c + 1).Value = props[c].Name;

        for (int r = 0; r < rows.Count; r++)
        {
            var item = rows[r];
            for (int c = 0; c < props.Count; c++)
            {
                object? val = props[c].GetValue(item, null);
                var cell = ws.Cell(r + 2, c + 1);
                SetCellValue(cell, val);
            }
        }

        ws.RangeUsed()?.SetAutoFilter();
        ws.Columns().AdjustToContents();
        wb.SaveAs(fullPath);
        return fullPath;
    }

    private static void SetCellValue(IXLCell cell, object? val)
    {
        if (val is null) { cell.Clear(); return; }

        switch (val)
        {
            case string s: cell.Value = s; break;
            case bool b: cell.Value = b; break;
            case DateTime dt:
                cell.Value = dt;
                cell.Style.DateFormat.Format = "yyyy-mm-dd HH:mm";
                break;
            case TimeSpan ts: cell.Value = ts; break;
            case byte or sbyte or short or ushort or int or uint or long or ulong:
                cell.Value = Convert.ToDouble(val); break;
            case float or double or decimal:
                cell.Value = Convert.ToDouble(val); break;
            case Guid g: cell.Value = g.ToString(); break;
            default:
                if (val.GetType().IsEnum) cell.Value = val.ToString();
                else cell.Value = val.ToString() ?? string.Empty;
                break;
        }
    }

    public static async Task SendEmailWithAttachmentAsync(
        List<string> recipientEmails,
        string subject,
        string htmlBody,
        List<string>? attachmentFilePaths = null)
    {
        //BİLGİLERİ DEĞİŞTİR--------------------------------------------------------------
        string smtpServer = "x.x.x.xxx";
        int smtpPort = 25;
        string senderEmail = "gizli@gizli.com";
        string senderPassword = "gizli_xxx";
        //--------------------------------------------------------------------------------
        var tasks = new List<Task>();

        foreach (var recipient in recipientEmails.Distinct())
        {
            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    using var mail = new MailMessage
                    {
                        From = new MailAddress(senderEmail, "SEÇ | Gizli Şirket"), //Bilgileri Değiştir
                        Subject = subject,
                        Body = htmlBody,
                        IsBodyHtml = true,
                        BodyEncoding = Encoding.UTF8,
                        Priority = MailPriority.High
                    };
                    mail.To.Add(recipient);

                    // Ek(ler) - MemoryStream ile
                    if (attachmentFilePaths != null)
                    {
                        foreach (var path in attachmentFilePaths.Where(File.Exists))
                        {
                            long size = new FileInfo(path).Length;
                            Console.WriteLine($"Ek dosya: {Path.GetFileName(path)}, Boyut: {size} bayt");

                            byte[] fileBytes = File.ReadAllBytes(path);
                            var stream = new MemoryStream(fileBytes); 
                            var attachment = new Attachment(stream, Path.GetFileName(path));
                            mail.Attachments.Add(attachment);
                        }
                    }

                    using var smtp = new SmtpClient(smtpServer, smtpPort)
                    {
                        EnableSsl = false,
                        Credentials = new NetworkCredential("info.sec", senderPassword, "gizli.local") // Bilgileri Değiştir
                    };

                    await smtp.SendMailAsync(mail);
                    Console.WriteLine($"Gönderildi: {recipient}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Hata [{recipient}]: {ex}");
                }
            }));
        }

        await Task.WhenAll(tasks);
    }

    // ---------- E-posta gövdesi ----------
    private static string BuildHtmlBodyForReport()
    {
        return @$"<!DOCTYPE html>
<html lang='tr'>
<head>
<meta charset='UTF-8'>
<style>
body {{ font-family: Arial, sans-serif; line-height: 1.6; color: rgb(0,62,81); background-color: #F3F3F3; padding: 20px; }}
.email-container {{ max-width: 680px; margin: 0 auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); }}
.email-header {{ text-align: center; margin-bottom: 16px; }}
.email-header h1 {{ font-size: 20px; margin: 0; color: rgb(0,62,81); }}
.email-content p {{ margin: 8px 0; }}
.email-footer {{ text-align:center; color:#777; font-size:12px; margin-top: 20px; }}
</style>
</head>
<body>
  <div class='email-container'>
    <div class='email-header'>
      <h1>Uygunsuzluklar Raporu</h1>
    </div>
    <div class='email-content'>
      <p>Merhaba Bilgesu,</p>
      <p>Uygunsuzluklar tablosunun güncel dökümü ektedir (.xlsx).</p>
      <p>İyi çalışmalar.</p>
    </div>
    <div class='email-footer'>
      <p>© {DateTime.Now.Year} GİZLİ Holding A.Ş.</p> 
    </div>
  </div>
</body>
</html>";
    }
}
