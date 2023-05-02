using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Aspose.Cells;
using System.IO;
using System.Globalization;
using eBAControls;
using eBAPI.Connection;
using eBADB;
using eBAControls.eBABaseForm;
using eBAPI.DocumentManagement;

namespace ALFAsposeHelper
{
	/// <summary>
	/// Log sistemindeki formlara update ve ya import yapmak icin kullanilan class
	/// </summary>
	public class MultipleFormsUpload
	{

		static Dictionary<string, Dictionary<string, List<string>>> dictColumns;//FORMLARIN İSİMLERİNİ TUTULDUKLARI TXTLERİN İSİMLERİNİ VE FORMLARDAKİ KOLONLARIN İSİMLERİNİ TUTAN KÜTÜPHANE
		static Dictionary<string, Object> dictProjectRegion; //Proje bölgesini sözleşme ekranına taşımak için kullanılıyor.
		/// <summary>
		/// formun unique idsi ve globalAttribute classini kullanan constructor, globalAttribute.Forms null olmamali!
		/// </summary>
		/// <param name="globalId"></param>
		/// <param name="GlobalAttributes"></param>
		public MultipleFormsUpload(string globalId, GlobalAttributes GlobalAttributes)
		{
			if (dictColumns == null) dictColumns = new Dictionary<string, Dictionary<string, List<string>>>();

			Dictionary<string, List<string>> temp = new Dictionary<string, List<string>>();
			foreach (var form in GlobalAttributes.Forms)
			{
				temp.Add(form.Item3, new List<string>());
			}
			dictColumns.Add(globalId, temp);

			if (dictProjectRegion == null) dictProjectRegion = new Dictionary<string, object>();
			dictProjectRegion.Add(globalId, DBNull.Value);

		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes,
			eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, null, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="AdditionalUpdate">Update ya da import sirasinda sql tarafinda harici bir islem yapilmasi gerekirse bu parametreyi kullanarak kendi methodunuzu olusturabilirsiniz</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, AdditionalUpdate, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="Validations">Mutlaka girilmesi gereken alanlar, olmamasi gereken eslesmeler gibi projeye ozel kontroller icin method olusturup bu parametreyi kullanabilirsiniz</param>
		/// <param name="parameters">Validations methodu icin extra parametrelere gerek duyarsaniz object array olarak atama yapip methodunuzda ilgili castler ile bu parametreden nesnelerinizi cekebilirsiniz</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, string, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, null, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="Validations">Mutlaka girilmesi gereken alanlar, olmamasi gereken eslesmeler gibi projeye ozel kontroller icin method olusturup bu parametreyi kullanabilirsiniz</param>
		/// <param name="parameters">Validations methodu icin extra parametrelere gerek duyarsaniz object array olarak atama yapip methodunuzda ilgili castler ile bu parametreden nesnelerinizi cekebilirsiniz</param>
		/// <param name="AdditionalUpdate">Update ya da import sirasinda sql tarafinda harici bir islem yapilmasi gerekirse bu parametreyi kullanarak kendi methodunuzu olusturabilirsiniz</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, string, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, AdditionalUpdate, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="SecondValidation">extrem durumlar haricinde kullanilmasina gerek olmayacak bir parametre, genel olarak validations parametresi ile istenilen validasyonlar yapilabilir</param>
		/// <param name="secondParameters">secondValidation methodunun parametresi</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] secondParameters,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, null, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="SecondValidation">extrem durumlar haricinde kullanilmasina gerek olmayacak bir parametre, genel olarak validations parametresi ile istenilen validasyonlar yapilabilir</param>
		/// <param name="secondParameters">secondValidation methodunun parametresi</param>
		/// <param name="AdditionalUpdate">Update ya da import sirasinda sql tarafinda harici bir islem yapilmasi gerekirse bu parametreyi kullanarak kendi methodunuzu olusturabilirsiniz</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] secondParameters,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, AdditionalUpdate, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="Validations">Mutlaka girilmesi gereken alanlar, olmamasi gereken eslesmeler gibi projeye ozel kontroller icin method olusturup bu parametreyi kullanabilirsiniz</param>
		/// <param name="SecondValidation">extrem durumlar haricinde kullanilmasina gerek olmayacak bir parametre, genel olarak validations parametresi ile istenilen validasyonlar yapilabilir</param>
		/// <param name="parameters">Validations methodu icin extra parametrelere gerek duyarsaniz object array olarak atama yapip methodunuzda ilgili castler ile bu parametreden nesnelerinizi cekebilirsiniz</param>
		/// <param name="secondParameters">secondValidation methodunun parametresi</param>
		/// <param name="AdditionalUpdate">Update ya da import sirasinda sql tarafinda harici bir islem yapilmasi gerekirse bu parametreyi kullanarak kendi methodunuzu olusturabilirsiniz</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, string, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			bool multiLanguage)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, AdditionalUpdate, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
		/// <param name="logonUser">yukleme yapan kullanicinin eba kullanici IDsi ornegin; "afidan"</param>
		/// <param name="lang">ebada kullanilan dil ornegin; "Turkish","English"</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateServerConnection">dllde hata verdigi icin ebadaki CreateServerConnection methodu parametre olarak verilir</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="Validations">Mutlaka girilmesi gereken alanlar, olmamasi gereken eslesmeler gibi projeye ozel kontroller icin method olusturup bu parametreyi kullanabilirsiniz</param>
		/// <param name="SecondValidation">extrem durumlar haricinde kullanilmasina gerek olmayacak bir parametre, genel olarak validations parametresi ile istenilen validasyonlar yapilabilir</param>
		/// <param name="parameters">Validations methodu icin extra parametrelere gerek duyarsaniz object array olarak atama yapip methodunuzda ilgili castler ile bu parametreden nesnelerinizi cekebilirsiniz</param>
		/// <param name="secondParameters">secondValidation methodunun parametresi</param>
		/// <param name="multiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Upload(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, string, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			bool multiLanguage
			)
		{
			Execute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, null, multiLanguage);
		}

		/// <summary>
		/// Visual Studyoda test etmek harici kullanilmamali!!!
		/// </summary>
		/// <param name="info"></param>
		/// <param name="globalAttributes"></param>
		/// <param name="attUpload"></param>
		/// <param name="logonUser"></param>
		/// <param name="language"></param>
		/// <param name="globalId"></param>
		/// <param name="CreateServerConnection"></param>
		/// <param name="CreateDatabaseProvider"></param>
		/// <param name="ShowMessageBox"></param>
		/// <param name="WriteToResponse"></param>
		/// <param name="Validations"></param>
		/// <param name="SecondValidation"></param>
		/// <param name="parameters"></param>
		/// <param name="secondParameters"></param>
		/// <param name="AdditionalUpdate"></param>
		/// <param name="multiLanguage"></param>
		/// <param name="excelStream"></param>
		public void Execute(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string language,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, string, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			bool multiLanguage,
			Stream excelStream = null)
		{// lang == TR,EN,RU
			InfoLog logRecord = new InfoLog();
			string langCode = "EN";
			switch (language)
			{
				case "Turkish":
					langCode = "TR";
					break;
				case "Russian":
					langCode = "RU";
					break;
				default:
					langCode = "EN";
					break;
			}
			try
			{
				if (attUpload != null) attUpload.ReadOnly = true; //Bir kez dosya yüklendikten sonra bir daha yüklenemesin diye readonly yapılıyor.
				List<Tuple<string, string, Type>> columns = Fields.TakeFields(info, "upload", null, CreateDatabaseProvider, globalAttributes);
				List<string> columnNames = columns.Select(x => (x.Item1)).ToList();
				DataTable excelTable = new DataTable();
				DataTable excelTableTemp = new DataTable(); // Excelden gelen verileri typeları ile çekmek için geçiçi bir template table oluşturuluyor
				eBAConnection ebacon = CreateServerConnection();
				Stream data = null;
				ebacon.Open();
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				Connection.Open();
				DMCategoryContentCollection flist = null;
				FileSystem fs = ebacon.FileSystem;
				if (attUpload != null)
				{
					DMFile file = fs.GetFile(attUpload.Filename);
					flist = file.GetAttachments(info.AttachmentCategory);
					if (flist.Count == 0) return;
				}
				List<KeyValuePair<string[], List<string>>> formColumns = new List<KeyValuePair<string[], List<string>>>();
				foreach (Tuple<string, string, string> form in globalAttributes.Forms)
				{
					formColumns.Add(new KeyValuePair<string[], List<string>>(new string[] { form.Item1, form.Item2 }, dictColumns[globalId][form.Item3]));
				}
				if (attUpload != null)
				{
					DMFileContent att = flist[0];
					data = fs.CreateFileAttachmentContentDownloadStream(attUpload.Filename, info.AttachmentCategory, att.ContentName);
				}
				else
				{
					data = excelStream;
				}
				Workbook mybook = new Workbook(data);
				//mybook.Worksheets[0].Cells.DeleteRow(1);
				//mybook.Worksheets[0].Cells.DeleteRow(1);
				int satirsay = mybook.Worksheets[0].Cells.MaxDataRow; //ilk satırı başlık olarak alıyor
				int kolonsay = mybook.Worksheets[0].Cells.MaxDataColumn;
				ExportTableOptions options = new ExportTableOptions
				{
					ExportColumnName = true
				};
				//EXCELDEN GELEN TABLOYU DOĞRU DATATYPELAR İLE OLUŞTURMAK İÇİN KULLANILAN GEÇİCİ DATATABLE
				excelTableTemp = mybook.Worksheets[0].Cells.ExportDataTable(0, 0, satirsay + 1, kolonsay + 1, options);
				excelTable.Columns.Add(info.UniqueColumnInfo.Item1, info.UniqueColumnType);
				var columnsTemp = new List<Tuple<string, string, Type>>();
				foreach (DataColumn cl in excelTableTemp.Columns)
				{   //KULLANILACAK TABLODAKİ ALANLARIN TYPELARINI BELİRLEYEN DÖNGÜ
					foreach (Tuple<string, string, Type> column in columns)
					{
						string columnName = column.Item1;
						string formName = column.Item2;
						Type columnType = column.Item3;
						if (columnName.Equals(cl.ColumnName))
						{
							columnsTemp.Add(column);
							//excelTable.Columns.Add(cl.ColumnName, !(column.Value.Equals(typeof(DateTime))) ? typeof(string) : typeof(DateTime));
							if (columnType.Equals(typeof(DateTime))) excelTable.Columns.Add(cl.ColumnName, typeof(DateTime));
							else if (columnType.Equals(typeof(Decimal))) excelTable.Columns.Add(cl.ColumnName, typeof(Decimal));
							else excelTable.Columns.Add(cl.ColumnName, typeof(string));
						}
					}
				}
				columns = columnsTemp;
				options.DataTable = excelTable;
				options.SkipErrorValue = true;
				Worksheet worksheet = mybook.Worksheets[0];
				//ALFDebugHelper.Log(88, excelTable.Columns, excelTableTemp.Columns);
				try
				{
					excelTable = worksheet.Cells.ExportDataTable(0, 0, satirsay + 1, kolonsay + 1, options);
				}
				catch (IndexOutOfRangeException) // Yetkisiz alanlar var ise sistem index out of range hatası verecektir bu durumda ilerlenmemesi gerekir.
				{
					throw new Exception("Yetkinizin olmadığı alanlar mevcuttur, lütfen yeni bir template oluşturup tekrar deneyiniz!");//, eBAMessageBoxType.Warning);

				}
				int index = 0;
				for (int i = 0; i < excelTable.Columns.Count; i++)
				{
					{
						index++;
					}
					foreach (KeyValuePair<string[], List<string>> formcolumn in formColumns)
					{//FORMLARDAKİ ALANLARIN İSİMLERİNİ LİSTEYE EKLEYEN DÖNGÜ
						if (formcolumn.Key[0] == excelTableTemp.Rows[0][i].ToString())
						{
							if (globalAttributes.RestrictedFields.Contains(excelTable.Columns[i].ColumnName)) continue;
							formcolumn.Value.Add((excelTable.Columns[i].ColumnName));
							break;
						}
					}
				}
				Dictionary<string, string> dictLabels = new Dictionary<string, string>();
				foreach (DataColumn col in excelTableTemp.Columns)
				{
					dictLabels[col.ColumnName] = excelTableTemp.Rows[1][col.ColumnName].ToString();
				}
				excelTable.Rows.RemoveAt(0);
				excelTable.Rows.RemoveAt(0);
				excelTableTemp.Rows.RemoveAt(0);
				excelTableTemp.Rows.RemoveAt(0);
				//VALİDASYON KONTROLLERİ
				List<string> idsToDelete = new List<string>();
				Pondera Pondera = new Pondera();
				Dictionary<string, Dictionary<string, string>> textIdDict = Fields.TakeMtlDictionary(info, columns, langCode, Connection, multiLanguage);
				Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)Fields.TakeMtlValues(info, columns, langCode, CreateDatabaseProvider, Fields.AddDictionary, new Dictionary<string, DataTable>(), "", multiLanguage);
				idsToDelete = Fields.CellValidations(globalAttributes, excelTable, excelTableTemp, logRecord, dictFields);
				idsToDelete = idsToDelete.Distinct().ToList();
				foreach (DataRow row in excelTable.Rows)
				{
					if (idsToDelete.Contains(row[0].ToString()))
					{
						row.Delete();
					}
				}
				excelTable.AcceptChanges();
				if (Validations != null)
				{
					idsToDelete = Validations(info, excelTable, excelTableTemp, CreateDatabaseProvider, columns, langCode, logRecord, logonUser, textIdDict, parameters);
					idsToDelete = idsToDelete.Distinct().ToList();
					foreach (DataRow row in excelTable.Rows)
					{
						if (idsToDelete.Contains(row[0].ToString()))
						{
							row.Delete();
						}
					}
					excelTable.AcceptChanges();
				}

				try
				{
					//ALFDebugHelper.Log(12121, excelTable,excelTable.Columns);
					UpdateInfo(info, globalAttributes, logRecord, excelTable, formColumns, columns, globalId, logonUser, ebacon, Connection, SecondValidation, secondParameters, AdditionalUpdate, textIdDict, dictLabels);
				}
				catch (Exception ex)
				{//HATA ALINSA DAHİ DEĞİŞEN ALANLARI GÖSTERMEK VE HATANIN HANGİ PERSONELDE ALINDIĞINI GÖRMEK İÇİN 
				 //DOSYA İNDİRTİLİYORdictErrorLog[docGlobalId.Text]
					if (logRecord.ChangedFieldsNote != "")
					{
						string[] employeeIds = logRecord.ChangedFieldsNote.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
						string lastId = employeeIds[employeeIds.Length - 1].Split(':')[0].Trim();
						logRecord.ErrorLog = "Hatadan alınan ID: " + lastId + Environment.NewLine;
					}
					logRecord.DownloadLogZip(globalId, WriteToResponse);
					throw ex;
				}

				if (string.IsNullOrEmpty(logRecord.UsedForms) && string.IsNullOrEmpty(logRecord.ErrorLog))
				{
					ShowMessageBox(info.SuccesfullUpdate, eBAMessageBoxType.Information);
				}
				else
				{
					ShowMessageBox(info.FailedUpdate, eBAMessageBoxType.Information);

				}
				logRecord.DownloadLogZip(globalId, WriteToResponse); //Process ile ilgili notlar indirilir.
			}
			catch (KeyNotFoundException)
			{
				ShowMessageBox(info.ErrorOnUpdate, eBAMessageBoxType.Error);
				return;
			}
		}

		/// <summary>
		/// Gerekli validasyonların yapıldığı ve personel bazlı update işlemlerinin yapıldığı main method
		/// </summary>
		public void UpdateInfo(Info info,
			GlobalAttributes globalAttributes,
			InfoLog logRecord,
			DataTable dtExcel,
			List<KeyValuePair<string[], List<string>>> formColumns,
			List<Tuple<string, string, Type>> realColumns,
			string globalId,
			string logonUser,
			eBAConnection ebacon,
			SqlConnection Connection,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] secondParameters,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			Dictionary<string, Dictionary<string, string>> dictTextIds,
			Dictionary<string, string> dictLabels)
		{
			string uniqueIds = "";
			string uniqueColumnName = info.UniqueColumnInfo.Item1;
			//EXCELDEN GENEL PERSONNELLERİN TÜM VERİLERİNİ TEK SEFERDE ALMAK İÇİN IDLERİNİ BİRLEŞTİREN DÖNGÜ
			foreach (DataRow dr in dtExcel.Rows)
			{
				uniqueIds += "'" + dr[0].ToString() + "',";
			}
			uniqueIds += "'111'";
			DataTable dtForms = new DataTable();
			var comm = string.Format(@"SELECT FRM.*
                                         FROM E_{0}_{1} FRM WITH(NOLOCK)
                                        INNER JOIN DOCUMENTS D WITH(NOLOCK) ON FRM.ID = D.ID
                                        WHERE {2} in ({3}) AND D.DELETED = 0
                                        ", info.ProjectName, info.MainFormName, uniqueColumnName, uniqueIds);
			dtForms = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
			//PERSONNELLERİN MAİN MODAL FORMLARINI TUTAN DİCTİONARY
			Dictionary<string, DataTable> mainModalDict = new Dictionary<string, DataTable>();
			//LOG TABLOSUNDA YER ALAN KOLONLARIN İSİMLERİNİ TUTAN DİCTİONARY
			Dictionary<string, List<string>> logColumnsDict = new Dictionary<string, List<string>>();
			foreach (KeyValuePair<string[], List<string>> columns in formColumns)
			{
				string formName = columns.Key[0];
				if (string.Equals(formName, info.UniqueColumnInfo.Item2)) continue;
				string modalFormIds = "";
				foreach (DataRow dr in dtForms.Rows)
				{
					modalFormIds += "'" + dr[columns.Key[1]].ToString() + "',";
				}
				modalFormIds += "'111'";
				mainModalDict.Add(formName, GetMainModalForms(logRecord, modalFormIds, "E_" + info.ProjectName + "_" + formName, Connection, logonUser));
				logColumnsDict.Add(formName, GetLogColumns(info, logRecord, formName, Connection, logonUser));
			}
			for (int i = 0; i < dtExcel.Rows.Count; i++)
			{//EĞER KULLANICI IDSİ SİSTEMDE KAYITLI DEĞİLSE İŞLEMLERİN ES GEÇMESİ İÇİN OLUŞTURULAN KOŞUL
				if (!dtForms.AsEnumerable()
					.Any(arow => (dtExcel.Columns[uniqueColumnName].DataType == typeof(string) && dtExcel.Rows[i][uniqueColumnName].ToString() == arow.Field<String>(uniqueColumnName))
					|| ((dtExcel.Columns[uniqueColumnName].DataType == typeof(Int32) || dtExcel.Columns[uniqueColumnName].DataType == typeof(double)) && dtExcel.Rows[i][uniqueColumnName].ToString() == arow.Field<Int32>(uniqueColumnName).ToString())))
				{
					logRecord.ErrorLog += dtExcel.Rows[i][uniqueColumnName].ToString() + " ID'li Satıra Ait Bilgi Sistemde Mevcut değildir." + Environment.NewLine;
					continue;
				}
				DataColumnCollection excelColumns = dtExcel.Columns;
				bool validation = false;
				if (SecondValidation != null)
				{
					validation = SecondValidation(info, i, excelColumns, logRecord, dtExcel, secondParameters);
				}

				if (validation) continue;
				logRecord.ChangedFieldsNote += Environment.NewLine + dtExcel.Rows[i][uniqueColumnName].ToString() + " : ";
				foreach (KeyValuePair<string[], List<string>> columns in formColumns)
				{
					string formName = columns.Key[0];
					if (string.Equals(formName, info.MainFormName)) continue;
					//PERSONNELE AİT MODAL FORMU GETİREN DATATABLE EXPRESSİON
					DataRow drMainModal = mainModalDict[columns.Key[0]]
						.AsEnumerable()
						.FirstOrDefault(r => r.Field<string>(uniqueColumnName) == dtExcel.Rows[i][uniqueColumnName].ToString());
					//EĞER O FORMDA UPDATE EDİLECEK KOLON BULUNMUYORSA GERİ KALAN İŞLEMLERİN ES GEÇİŞLMESİ GEREKLİ
					if (drMainModal == null)//BİR İHTİMAL MODAL FORMU OLUŞMAMIŞ OLABİLİR BU DURUMDA O FORMU ES GEÇMESİ İÇİN KOYULAN KOŞUL
					{
						logRecord.MissingForms += Environment.NewLine + dtExcel.Rows[i][0].ToString() + "--" + formName;
						continue;
					}
					if (!(columns.Value.Count > 0)) continue;
					if (IsFormActive(info, logRecord, drMainModal, Connection, formName, logonUser)) continue;
					//ÖNCELİKLE EXCELDEN GELEN TÜM BİLGİLERLE İLGİLİ ANA MODAL FORM GÜNCELLENİR.
					try { dictProjectRegion[globalId] = dtExcel.Rows[i]["mtlProjectRegion"]; } catch { }
					DataTable dtMainModal2 = UpdateMainModal(info, logRecord, globalAttributes, columns, drMainModal, dtExcel.Rows[i], Connection, globalId, logonUser, AdditionalUpdate, dictTextIds);
					//BİR DEĞİŞİM OLUP OLMADIĞINI KONTROL ETMEK İÇİN AŞAĞIDAKİ METHOD KULLANILIR, DEĞİŞİM YOKSA DÖNGÜ DEVAM EDİLİR
					if (!IsChangedNew(globalAttributes, drMainModal, dtMainModal2.Rows[0], columns)) continue;
					//DEĞİŞİM VAR İSE YENİ BİR MODAL FORM OLUŞTURULUR
					int uvId = ebacon.WorkflowManager.CreateDocument(info.ProjectName, columns.Key[0]).DocumentId;
					//İLGİLİ DEĞİŞİMLE İLGİLİ LOG DÜŞÜLÜR VE SON OLARAK DA YENİ OLUŞTURULAN FORM GÜNCELLENİR
					InsertLogRow(logRecord, logColumnsDict[formName], uvId, drMainModal, dtMainModal2.Rows[0], columns, Connection, realColumns, logonUser, info, globalAttributes, dictLabels);
					UpdateNewForm(uvId, dtMainModal2, formName, Connection, info);
				}
			}
			Connection.Close();
		}

		/// <summary>
		/// Tüm personellerin ilgili Main modal formlarını sql üzerinden getiren method.
		/// </summary>
		/// <returns>Satır bazlı personellerin ilgili formlarını tutan Datatable </returns>
		private DataTable GetMainModalForms(InfoLog logRecord, string formId, string formName, SqlConnection Connection, string logonUser)
		{//TÜM PERSONELLERE AİT İLGİL MODAL FORMUN VERİLERİNİ GETİREN SORGU
			DataTable dt;
			var comm = string.Format(@"SELECT  *     
                                          FROM {0}
                                         WHERE ID in ({1})
                                        ", formName, formId);

			dt = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
			return dt;
		}

		/// <summary>
		/// Log tablosunda değişen alanların kontrol edilmesi ve tablodaki kolonların doldurulması için ilgili kolon isimlerini getiren method
		/// </summary>
		/// <returns></returns>
		public List<string> GetLogColumns(Info info, InfoLog logRecord, string formName, SqlConnection Connection, string logonUser)
		{
			DataTable dt;
			var comm = string.Format(@"SELECT TOP(1) * FROM E_{1}_{0}_tblLog ", formName, info.ProjectName);
			dt = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
			List<string> logColumns = new List<string>();
			foreach (DataColumn column in dt.Columns)
			{
				logColumns.Add(column.ColumnName);
			}
			return logColumns;
		}

		/// <summary>
		/// Formun başka bir kullanıcıda açık olup olmadığını kontrol eden method
		/// </summary>
		/// <returns></returns>
		public bool IsFormActive(Info info, InfoLog logRecord, DataRow dr, SqlConnection Connection, string formName, string logonUser)
		{
			DataTable dt;
			var comm = string.Format(@"SELECT Top(1) [CHECKOUTDATE] AS 'DATE' ,USERID
									  FROM  [dbo].[CHECKOUTS] 
									  WHERE FILENAME LIKE '%/{0}.%'
									  ORDER BY CHECKOUTDATE DESC", dr["ID"].ToString());

			dt = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
			if (dt.Rows.Count == 0) return false;
			else
			{
				DateTime checkOut = Convert.ToDateTime(dt.Rows[0]["Date"]);
				var hours = (DateTime.Now - checkOut).TotalHours;
				if (hours >= 3) return false;
				else
				{
					logRecord.UsedForms += Environment.NewLine + dr[info.UniqueColumnInfo.Item1].ToString() + " ID'li satıra Ait " + formName + " formu şu anda " + dt.Rows[0]["USERID"].ToString() + " userId'li personelin ekranında açık" +
						" ya da uygun bir biçimde kapatılmamıştır.(Açılma Tarihi :" + checkOut.ToString() + ")";
				}
			}
			return true;
		}

		/// <summary>
		/// Personel ve Form bazlı main formları güncelleyen method.
		/// </summary>
		/// <returns>Formların güncellenmiş halini geri döner</returns>
		private DataTable UpdateMainModal(Info info, InfoLog logRecord,
			GlobalAttributes globalAttributes,
			KeyValuePair<string[], List<string>> columns,
			DataRow drMain,
			DataRow drUpdate,
			SqlConnection Connection,
			string globalId,
			string logonUser,
			Func<KeyValuePair<string[], List<string>>, DataRow, string, string, string> AdditionalUpdate,
			Dictionary<string, Dictionary<string, string>> dictTextIds)
		{
			string mainFormId = drMain["ID"].ToString();
			string setQuery = "";
			foreach (string column in columns.Value)
			{//DEĞİŞMEMESİ GEREKEN ALANLAR VARSA BU ALANLAR QUERYE EKLENMEZ
				if (column == "ID" || column == "docGlobalId" || globalAttributes.RestrictedFields.Contains(column)) continue;
				else if (column.StartsWith("mtl") || column.StartsWith("cmb"))
				{
					//Bazı alanlar dropdownList olarak tutuluyor ve bu alanlarda seçim yapılmadıysa idlerinin -1 olması gerekli bu sebeple null değerlerde -1 atanıyor,
					//comboboxlarda idlerin -1 olması bir hataya yol açmadığı için onlar için ayrı bir validasyon kullanılmıyor.
					if (string.Equals(drMain[column].ToString(), "-1", StringComparison.CurrentCultureIgnoreCase))
						setQuery += (drUpdate[column] == DBNull.Value) ?
							(column + "_TEXT = '' ," + column + "='-1',") :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + dictTextIds[column][drUpdate[column].ToString()] + "',";
					else
					{
						setQuery += (drUpdate[column] == DBNull.Value) ?
							column + "_TEXT = NULL ," + column + " = NULL ," :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + dictTextIds[column][drUpdate[column].ToString()] + "',";
					}

				}

				else if (column.Contains("Date") || globalAttributes.DateColumns.Contains(column))
				{
					setQuery += (drUpdate[column] == DBNull.Value) ?
						column + " = NULL ," :
						column + " = convert(datetime,'" + ((DateTime)drUpdate[column]).ToString("dd/MM/yyyy") + "',103),";
				}
				else
				{
					if (drUpdate[column].GetType() == typeof(decimal))
					{
						setQuery += (drUpdate[column] == DBNull.Value) ?
							column + " = NULL ," :
							column + " = N'" + ((decimal)drUpdate[column]).ToString(new CultureInfo("en-US")) + "',";
					}
					else
					{
						setQuery += (drUpdate[column] == DBNull.Value) ?
							column + " = NULL ," :
							column + " = N'" + drUpdate[column].ToString() + "',";
					}
				}
			}
			if (AdditionalUpdate != null)
			{
				setQuery += AdditionalUpdate(columns, drUpdate, setQuery, globalId);
			}

			if (setQuery[setQuery.Length - 1] == ',')
			{

				setQuery = setQuery.Remove(setQuery.Length - 1);

			}
			DataTable dt;
			var comm = string.Format(@"UPDATE E_{3}_{2}
                                        SET {0}
                                        /*chbIsMain = 7*/
                                        WHERE ID = {1}
										/*UPDATE E_{3}_{2}
                                       SET chbIsMain = 
                                        WHERE ID = {1}1*/

                                        SELECT *
                                        FROM E_{3}_{2}
                                        WHERE ID = {1}
                                        ", setQuery, mainFormId, columns.Key[0], info.ProjectName);
			dt = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
			return dt;
		}


		bool IsChangedNew(
		   GlobalAttributes globalAttributes, DataRow drMain, DataRow drExcel, KeyValuePair<string[], List<string>> columns)
		{//EĞER VERİLERDEN BİRİSİ BİLE AYNI DEĞİLSE TRUE DÖNÜLEN BİR METHOD
			foreach (string column in columns.Value)
			{
				//Neden burada olduğunu hatırlamıyorum ama bir nedenden ötürü sonradan eklenmiş bir kod, silinmemeli Pondera için!
				if (drMain.Table.Columns.Contains("txtProjectRegion"))
				{
					if (string.Compare(drExcel["txtProjectRegion"].ToString(), drMain["txtProjectRegion"].ToString()) != 0)
					{
						return true;
					}
				}
				if (column.Contains("Date") || globalAttributes.DateColumns.Contains(column))
				{// GELEN VERİLER NULL OLABİLİR YA DA NULL ATANMASI GEREKEBİLİR BU SEBEPLE DATETİME? KULLANILDI
					DateTime? dateExcel = (drExcel[column] == System.DBNull.Value) ? (DateTime?)null : Convert.ToDateTime(drExcel[column]);
					DateTime? dateMain = (drMain[column] == System.DBNull.Value) ? (DateTime?)null : Convert.ToDateTime(drMain[column]);
					if (!DateTime.Equals(dateExcel, dateMain)) return true;
				}
				else if (column.StartsWith("txt"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
						return true;
				}
				else if (column.StartsWith("mtl") || column.StartsWith("cmb"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
						return true;
				}
			}
			return false;
		}
		void InsertLogRow(InfoLog logRecord,
		   List<string> logColumns,
		   int newModalId,
		   DataRow drMain,
		   DataRow drExcel,
		   KeyValuePair<string[], List<string>> columns,
		   SqlConnection Connection,
		   List<Tuple<string, string, Type>> realColumns,
		   string logonUser
		   , Info info,
		   GlobalAttributes globalAttributes,
		   Dictionary<string, string> dictLabels)
		{

			//List<KeyValuePair<KeyValuePair<string, string>, Type>> realColumns = TakeFields("upload");
			string changedFields = "";
			string qrColumn = "";
			string qrValue = "";
			string formName = columns.Key[0];
			foreach (Tuple<string, string, Type> columnsInfo in realColumns)
			{
				string column = columnsInfo.Item1;
				if (formName != columnsInfo.Item2) continue;
				if (column.Contains("Date") || globalAttributes.DateColumns.Contains(column))
				{
					DateTime? dateExcel = (drExcel[column] == System.DBNull.Value) ?
						(DateTime?)null :
						Convert.ToDateTime(drExcel[column]);
					DateTime? dateMain = (drMain[column] == System.DBNull.Value) ?
						(DateTime?)null :
						Convert.ToDateTime(drMain[column]);
					if (!DateTime.Equals(dateExcel, dateMain))
					{
						//changedFields += column.Substring(3) + "--";
						changedFields += dictLabels[column] + "--";
						//logRecord.changedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
						logRecord.ChangedFieldsNote += dictLabels[column] + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}
					if (logColumns.Contains(column))
					{
						qrColumn += column + ",";
						qrValue += "N'" + dateExcel.ToString() + "',";
					}
				}
				else if (column.StartsWith("txt"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
					{
						//changedFields += column.Substring(3) + "--";
						changedFields += dictLabels[column] + "--";
						//logRecord.changedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
						logRecord.ChangedFieldsNote += dictLabels[column] + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}
					if (logColumns.Contains(column))
					{
						qrColumn += column + ",";
						qrValue += "N'" + drExcel[column].ToString() + "',";
					}
				}
				else if (column.StartsWith("mtl") || column.StartsWith("cmb"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
					{
						//changedFields += column.Substring(3) + "--";
						changedFields += dictLabels[column] + "--";
						//logRecord.changedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
						logRecord.ChangedFieldsNote += dictLabels[column] + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}
					if (logColumns.Contains(column))
					{
						qrColumn += column + ",";
						qrValue += "N'" + drExcel[column].ToString() + "',";
						qrColumn += column + "_TEXT,";
						qrValue += "N'" + drExcel[column + "_TEXT"].ToString() + "',";
					}
				}
			}
			if (formName.Equals("MdlPersonnelEmployeeInfo"))
				foreach (string field in globalAttributes.RestrictedFields)
				{
					if (logColumns.Contains(field))
					{
						qrColumn += field + ",";
						qrValue += "N'" + drExcel[field].ToString() + "',";
						qrColumn += field + "_TEXT,";
						qrValue += "N'" + drExcel[field + "_TEXT"].ToString() + "',";
					}
				}
			if (string.IsNullOrEmpty(changedFields))
				return;
			var comm = string.Format(@"INSERT INTO E_{7}_{0}_tblLog (FORMID,ORDERID,CHECKED,{1} LogDate, LogUser,ModalFormId,ChangedFields)
                                        VALUES('{5}',ISNULL((SELECT TOP(1) ORDERID
                                          FROM E_{7}_{0}_tblLog
                                          WHERE FORMID = {5}
                                          ORDER BY ORDERID DESC)+1,0),0,{2} FORMAT(GETDATE(),'dd/MM/yyyy hh:mm:ss tt'),'{6}(Excel)','{3}','{4}')
                                        ", formName,
										qrColumn,
										qrValue,
										newModalId.ToString("D10"),
										changedFields,
										drMain["ID"].ToString(),
										logonUser,
										info.ProjectName);

			SqlCalls.RunQueriesNoReturn(comm, Connection);
		}
		/// <summary>
		/// oluşturulan yeni ara formu, update edilen ana forma göre güncelleyen method,chbIsmain diger projelerde de olmali
		/// </summary>
		private void UpdateNewForm(int newFormId, DataTable dtMain, string formName, SqlConnection Connection, Info info)
		{
			DataColumnCollection columns = dtMain.Columns;
			DataRow drMain = dtMain.Rows[0];
			string mainFormId = drMain["ID"].ToString();
			string setQuery = "";
			foreach (DataColumn column in columns)
			{
				string columnName = column.ColumnName;
				if (info.UpdateExceptions.Contains(columnName))
					continue;
				setQuery += columnName + " = MN." + columnName + ",";
			}
			if (setQuery[setQuery.Length - 1] == ',')
			{

				setQuery = setQuery.Remove(setQuery.Length - 1);

			}
			var comm = string.Format(@"UPDATE E_{4}_{3}
                                        SET {0}
                                        /*chbIsMain = 0*/
                                        FROM (SELECT *
                                                FROM E_{4}_{3}
                                                WHERE ID = {1}) MN
                                                WHERE E_{4}_{3}.ID = {2}
                                        ", setQuery, mainFormId, newFormId, formName, info.ProjectName);
			SqlCalls.RunQueriesNoReturn(comm, Connection);
		}

	}
}
