using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALFAsposeHelper
{
	/// <summary>
	/// Proje detaylarinin tutuldugu class
	/// </summary>
	public class Info
	{
		/// <summary>
		/// veri yukleme tipi
		/// </summary>
		public  enum EnumUploadType
		{
			/// <summary>
			/// update
			/// </summary>
			Update,
			/// <summary>
			/// import
			/// </summary>
			Import
		}
		/// <summary>
		/// veri indirme tipi
		/// </summary>
		public enum EnumDownloadType
		{
			/// <summary>
			/// update
			/// </summary>
			Update,
			/// <summary>
			/// import
			/// </summary>
			Import
		}
		/// <summary>
		/// veri yukleme tipi
		/// </summary>
		public EnumUploadType UploadType { get; set; }
		/// <summary>
		/// veri indirme tipi
		/// </summary>
		public EnumDownloadType DownloadType { get; set; }
		/// <summary>
		/// log sistemi icin authorization sorgusunun bulundugu connection adi
		/// </summary>
		public string AuthConnectionName { get; set; }
		/// <summary>
		/// log sistemi icin authorization sorgusunun adi
		/// </summary>
		public string AuthQueryName { get; set; }
		/// <summary>
		/// Visual Studyoda calistirmak icin gerekli auth querysinin tamami(integration managerdaki orjinal hali)
		/// </summary>
		public string AuthQuery { get; set; } = null;
		/// <summary>
		/// Integratin managerdaki mantikla calisan query parametreleri
		/// </summary>
		public List<KeyValuePair<string, string>> AuthParameters { get; set; }
		/// <summary>
		/// tek form ise formun degilse ana formun adi => "FrmPersonnel"
		/// </summary>
		public string MainFormName { get; set; }
		/// <summary>
		/// Eğer işlem Sql tablosu İşlemi değilse ilgili Projenin adı => "Pon002Personnel"
		/// </summary>
		public string ProjectName { get; set; }
		/// <summary>
		/// Sistemde Lookup Mevcutsa Lookup Valuelarinin tutuldugu Tablonun Adi => "TbPon000LookupValues"
		/// </summary>
		public string MtlTableName { get; set; }
		/// <summary>
		/// Eger Lookupta MultiLang varsa dilleri icermeyen column adi, degilse veriyi tutan columnun adi => "VALUE1"
		/// </summary>
		public string MtlColumnName { get; set; }
		/// <summary>
		/// Update Yapmak Icin gerekli olan Unique Columun veri tipi => typeof(string)
		/// </summary>
		public Type UniqueColumnType { get; set; }
		/// <summary>
		/// <para>Lookuptan beslenmeyen cmblerin ve dropdownlarin iceriklerini excele yazmamis icin gerekli sorgulari saklayan dictionary, </para>
		/// <para>key olarak ilgili kolonun adini value olarak da 2 kolon iceren bir sql sorgusu tutar, sorgudan </para>
		/// <para>donen veriler ilgili kolonun alabilecegi veriler ve IDleridir, ilk kolon text ikinci kolon value/ID olmalidir</para>
		/// <para>org =>  { "cmbGroupChief", "SELECT DISTINCT VALUE, ID AS LOOKUP FROM TbPon015ForemanList " }</para>
		/// </summary>
		public Dictionary<string, string> DictCmbFieldQueries { get; set; }
		
		/// <summary>
		/// Artik kullanilmiyor
		/// </summary>
		public Dictionary<string, string> DictCmbDictionaryQueries { get; set; }
		/// <summary>
		/// Artik kullanilmiyor ama kullanildigi zaman bu sekilde kullaniliyordur =>"PONDM/ExcelTemplates/DataBulkUpdate.xlsm"
		/// </summary>
		public string TemplatePath { get; set; }
		/// <summary>
		/// Indirilecek Dosyanin uzantisiyla birlikte Adi"Data_Bulk_Update.xlsm"
		/// </summary>
		public string DownloadName { get; set; } 
		/// <summary>
		/// Update icin veri secilmedigi zaman kullaniciya gosterilecek Uyari mesaji default olarak "Lutfen once secim yapiniz" yazilidir
		/// </summary>
		public string NotEnoughObjectMessage { get; set; } = "Lutfen once secim yapiniz";
		/// <summary>
		/// <para>Updatein gerceklestirilebilmesi adina unique bir kolon gerekli ID/Employee ID gibi, bu kolonun adi, </para>
		/// <para>tutuldugu form, excelde gosterilecek Labeli ve secilen verileri getirecek sorgusu tuple olarak tutulur</para>
		/// <para>new Tuple{string, string, string,string}("txtEmployeeId", "FrmPersonnel", "Employee ID / Sicil No", @"SELECT txtEmployeeId</para>
		/// <para>FROM E_Pon002Personnel_FrmPersonnel FRM WITH(NOLOCK)</para>
		///	<para>INNER JOIN DOCUMENTS D WITH(NOLOCK) ON FRM.ID = D.ID</para>
		///	<para>WHERE txtEmployeeId in ({ 0}) AND D.DELETED = 0</para>
		///	<para>ORDER BY FRM.txtEmployeeId");</para>
		/// </summary>
		public Tuple<string, string,string, string> UniqueColumnInfo { get; set; } 
		/// <summary>
		/// Excelde gosterilecek ama yuklendiginde update edilmeyecek veriler icin bu liste kullanilir, icerigi uniquecolumnInfo ile aynidir
		/// </summary>
		public List<Tuple<string, string, string, string>> PreDefinedColums { get; set; } = null;
		/// <summary>
		/// Genel olarak attachment kategorisi default tutuluyor ama olasi durumda farkli bir kategori istenilirse bu field kullanilabilir
		/// </summary>
		public string AttachmentCategory { get; set; } = "default";
		/// <summary>
		/// <para>Update yapilirken bazi Verilerde hata olustuysa gosterilecek mesaji belirler default olarak,</para>
		/// <para>"Güncelleme Tamamlandı ama Bazı Hatalar Mevcutö Lütfen İndirilmiş Zip Dosyasını Kontrol Ediniz!" yazmaktadir </para>
		/// </summary>
		public string FailedUpdate { get; set; } = "Güncelleme Tamamlandı ama Bazı Hatalar Mevcutö Lütfen İndirilmiş Zip Dosyasını Kontrol Ediniz!";
		/// <summary>
		/// Update Basarili sekilde tamamlanmissa bu mesaj gosterilir, default olarak "Güncelleme Başarılı Şekilde Tamamlandı" yazmaktadir
		/// </summary>
		public string SuccesfullUpdate { get; set; } = "Güncelleme Başarılı Şekilde Tamamlandı";
		/// <summary>
		/// Hatadan dolayi her hangi bir sekilde surec iptal edildiyse bu mesaj  gozukur, "Sistemsel Bir Hata Oluştu, Lütfen Formu Yenileyip Tekrar Deneyiniz!" default valuesudur
		/// </summary>
		public string ErrorOnUpdate { get; set; } = "Sistemsel Bir Hata Oluştu, Lütfen Formu Yenileyip Tekrar Deneyiniz!";
		/// <summary>
		/// Artik Kullanilmiyor
		/// </summary>
		public string MtlUpdateQuery { get; set; } = "EXEC  sp016GetMtlıd @LANGVALUE = {0}_TEXT, @FORM = {1} , @COLUMNNAME = '{2}' ,@FORMID = @@@@@@ ";
		/// <summary>
		/// Log sisteminde ana formda guncellenip modal formda guncellenmesini istemedigimiz veriler varsa onlar icin kullanilir
		/// </summary>
		public HashSet<string> UpdateExceptions { get; set; } = new HashSet<string>() { "ID", "docGlobalId"
			//, "txtEmployeeId"
			, "chbIsMain"
			//, "txtProjectId", "txtLocationStatusId", "cmbResponsibleStaff_TEXT", "cmbGroupChief_TEXT" 
		};

		
	}
}
