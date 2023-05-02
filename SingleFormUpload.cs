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
{/// <summary>
 /// Ebadaki Tek bir forma update ve ya import icin kullanilan class
 /// </summary>
	public class SingleFormUpload
	{/// <summary>
	 /// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
	 /// </summary>
	 /// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
	 /// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
	 /// <param name="attUpload">attachment nesnesi</param>
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
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, null, null, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, null, null, multiLanguage);
			}
		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			GlobalAttributes globalAttributes,
			eBAAttachments attUpload,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, null, null, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, null, null, multiLanguage);
			}
		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			GlobalAttributes globalAttributes,
			eBAAttachments attUpload,
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			GlobalAttributes globalAttributes,
			eBAAttachments attUpload,
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
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, null, null, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, null, null, multiLanguage);
			}
		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
		}
		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
			}
		}

		/// <summary>
		/// Excel Uzerinden Ebadaki tek bir forma update ya da import icin kullanilir
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="attUpload">attachment nesnesi</param>
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
			string lang,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			bool multiLanguage
			)
		{
			if (info.UploadType == Info.EnumUploadType.Import)
			{
				ExecuteImport(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, null, null, multiLanguage);
			}
			else
			{
				ExecuteUpdate(info, globalAttributes, attUpload, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, null, null, multiLanguage);
			}
		}
		/// <summary>
		/// sadece test ederken visual studyoda kullanilmalidir
		/// </summary>
		/// <param name="info"></param>
		/// <param name="globalAttributes"></param>
		/// <param name="attUpload"></param>
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
		public void ExecuteUpdate(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string language,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
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
				List<Tuple<string, string, Type>> columns = Fields.TakeFields(info, null, CreateDatabaseProvider, globalAttributes);
				List<string> columnNames = columns.Select(x => (x.Item1)).ToList();
				DataTable excelTable = new DataTable();
				DataTable excelTableTemp = new DataTable(); // Excelden gelen verileri typeları ile çekmek için geçiçi bir template table oluşturuluyor
				eBAConnection ebacon = null;
				if (attUpload != null)
				{
					ebacon = CreateServerConnection();
					ebacon.Open();
				}
				Stream data = null;
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				Connection.Open();
				DMCategoryContentCollection flist = null;
				FileSystem fs = null;
				if (attUpload != null)
				{
					fs = ebacon.FileSystem;
					DMFile file = fs.GetFile(attUpload.Filename);
					flist = file.GetAttachments(info.AttachmentCategory);
					if (flist.Count == 0) return;
				}
				List<string> formColumns = new List<string>();
				//attUpload.LoadAttachments();
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
					throw new Exception("Hatali Bir Template Yuklediniz, lütfen yeni bir template oluşturup tekrar deneyiniz!");//, eBAMessageBoxType.Warning);

				}
				int index = 0;
				for (int i = 0; i < excelTable.Columns.Count; i++)
				{
					{
						index++;
					}
					if (globalAttributes.RestrictedFields.Contains(excelTable.Columns[i].ColumnName)) continue;
					formColumns.Add((excelTable.Columns[i].ColumnName));
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
				Dictionary<string, Dictionary<string, string>> textIdDict = Fields.TakeMtlDictionary(info, columns, langCode, Connection, multiLanguage);
				Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)Fields.TakeMtlValues(info, columns, langCode, CreateDatabaseProvider, Fields.AddDictionary, new Dictionary<string, DataTable>(), "", multiLanguage);
				idsToDelete = Fields.CellValidations(globalAttributes, excelTable, excelTableTemp, logRecord, dictFields);
				idsToDelete = idsToDelete.Distinct().ToList();
				excelTable.AcceptChanges();
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
					idsToDelete = Validations(info, excelTable, excelTableTemp, CreateDatabaseProvider, columns, langCode, logRecord, textIdDict, parameters);
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
					UpdateInfo(info, globalAttributes, logRecord, excelTable, formColumns, globalId, Connection, SecondValidation, secondParameters, AdditionalUpdate, AdditionalQuery, textIdDict);
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
		void UpdateInfo(Info info,
		   GlobalAttributes globalAttributes,
		   InfoLog logRecord,
		   DataTable dtExcel,
		   List<string> formColumns,
		   string globalId,
		   SqlConnection Connection,
		   Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
		   object[] secondParameters,
		   Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
			   Func<List<string>, DataRow, string, string> AdditionalQuery,
		   Dictionary<string, Dictionary<string, string>> dictTextIds)
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
                                        WHERE FRM.{2} in ({3}) AND D.DELETED = 0
                                        ", info.ProjectName, info.MainFormName, uniqueColumnName, uniqueIds);
			dtForms = SqlCalls.RunQueries(comm, Connection);

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
				string formName = info.MainFormName;
				Type type = info.UniqueColumnType;
				//PERSONNELE AİT MODAL FORMU GETİREN DATATABLE EXPRESSİON
				DataRow drMainForm = dtForms
					.AsEnumerable()
					.FirstOrDefault(r => (type == typeof(int) ? r.Field<int>(uniqueColumnName).ToString() : r.Field<string>(uniqueColumnName)) == dtExcel.Rows[i][uniqueColumnName].ToString());
				//EĞER O FORMDA UPDATE EDİLECEK KOLON BULUNMUYORSA GERİ KALAN İŞLEMLERİN ES GEÇİŞLMESİ GEREKLİ

				if (!(formColumns.Count > 0)) continue;
				if (IsFormActive(info, logRecord, drMainForm, Connection, formName)) continue;
				//ÖNCELİKLE EXCELDEN GELEN TÜM BİLGİLERLE İLGİLİ ANA MODAL FORM GÜNCELLENİR.
				DataTable dtMainModal2 = UpdateForm(info, globalAttributes, formColumns, drMainForm, dtExcel.Rows[i], Connection, globalId, AdditionalUpdate, AdditionalQuery, dictTextIds);


				//int uvId = ebacon.WorkflowManager.CreateDocument(info.ProjectName, columns.Key[0]).DocumentId;


				//İLGİLİ DEĞİŞİMLE İLGİLİ LOG DÜŞÜLÜR VE SON OLARAK DA YENİ OLUŞTURULAN FORM GÜNCELLENİR
				InsertLogRow(logRecord, drMainForm, dtMainModal2.Rows[0], formColumns, globalAttributes);

			}
			Connection.Close();
		}
		/// <summary>
		/// sadece test etmek icin kullanilmalidir
		/// </summary>
		/// <param name="info"></param>
		/// <param name="globalAttributes"></param>
		/// <param name="attUpload"></param>
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
		public void ExecuteImport(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string language,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, DataTable, Func<eBADBProvider>, List<Tuple<string, string, Type>>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
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
				List<Tuple<string, string, Type>> columns = Fields.TakeFields(info, null, CreateDatabaseProvider, globalAttributes);
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
				List<string> formColumns = new List<string>();
				//attUpload.LoadAttachments();
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
				//excelTable.Columns.Add(info.UniqueColumnInfo.Item1, info.UniqueColumnType);
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
					throw new Exception("Hatali Bir Template Yuklediniz, lütfen yeni bir template oluşturup tekrar deneyiniz!");//, eBAMessageBoxType.Warning);

				}
				int index = 0;
				for (int i = 0; i < excelTable.Columns.Count; i++)
				{
					{
						index++;
					}
					if (globalAttributes.RestrictedFields.Contains(excelTable.Columns[i].ColumnName)) continue;
					formColumns.Add((excelTable.Columns[i].ColumnName));
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
				IEnumerable<int> idsToDeleteInt = new List<int>();
				Dictionary<string, Dictionary<string, string>> textIdDict = Fields.TakeMtlDictionary(info, columns, langCode, Connection, multiLanguage);
				Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)Fields.TakeMtlValues(info, columns, langCode, CreateDatabaseProvider, Fields.AddDictionary, new Dictionary<string, DataTable>(), "", multiLanguage);
				idsToDelete = Fields.CellValidations(globalAttributes, excelTable, excelTableTemp, logRecord, dictFields, false);
				idsToDelete = idsToDelete.Distinct().ToList();
				idsToDeleteInt = idsToDelete.Select(x => (Int32.Parse(x)));
				excelTable.AcceptChanges();
				int k = 0;
				foreach (DataRow row in excelTable.Rows)
				{
					if (idsToDeleteInt.Contains(k))
					{
						row.Delete();
					}
					k++;
				}
				excelTable.AcceptChanges();

				excelTable.AcceptChanges();
				if (Validations != null)
				{
					idsToDelete = Validations(info, excelTable, excelTableTemp, CreateDatabaseProvider, columns, langCode, logRecord, textIdDict, parameters);
					idsToDelete = idsToDelete.Distinct().ToList();
					idsToDeleteInt = idsToDelete.Select(x => (Int32.Parse(x)));
					excelTable.AcceptChanges();
					k = 0;
					foreach (DataRow row in excelTable.Rows)
					{
						if (idsToDeleteInt.Contains(k))
						{
							row.Delete();
						}
						k++;
					}
					excelTable.AcceptChanges();
				}

				try
				{
					//ALFDebugHelper.Log(12121, excelTable,excelTable.Columns);
					InsertInfo(info, globalAttributes, logRecord, excelTable, formColumns, globalId, ebacon, Connection, SecondValidation, secondParameters, AdditionalUpdate, AdditionalQuery, textIdDict);
				}
				catch (Exception ex)
				{//HATA ALINSA DAHİ DEĞİŞEN ALANLARI GÖSTERMEK VE HATANIN HANGİ PERSONELDE ALINDIĞINI GÖRMEK İÇİN 
				 //DOSYA İNDİRTİLİYORdictErrorLog[docGlobalId.Text]
				 //HATA ALINSA DAHİ DEĞİŞEN ALANLARI GÖSTERMEK VE HATANIN HANGİ PERSONELDE ALINDIĞINI GÖRMEK İÇİN 
				 //DOSYA İNDİRTİLİYORdictErrorLog[docGlobalId.Text]
					if (logRecord.ChangedFieldsNote != "")
					{
						string[] employeeIds = logRecord.ChangedFieldsNote.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
						string lastId = employeeIds[employeeIds.Length - 1].Split('.')[0].Trim();
						logRecord.ErrorLog = "Hatadan alınan Satir : " + lastId + Environment.NewLine;
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

		void InsertInfo(Info info,
		   GlobalAttributes globalAttributes,
		   InfoLog logRecord,
		   DataTable dtExcel,
		   List<string> formColumns,
		   string globalId,
		   eBAConnection ebacon,
		   SqlConnection Connection,
		   Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
		   object[] secondParameters,
		   Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
			   Func<List<string>, DataRow, string, string> AdditionalQuery,
		   Dictionary<string, Dictionary<string, string>> dictTextIds)
		{
			//string uniqueColumnName = info.UniqueColumnInfo.Item1;
			//EXCELDEN GENEL PERSONNELLERİN TÜM VERİLERİNİ TEK SEFERDE ALMAK İÇİN IDLERİNİ BİRLEŞTİREN DÖNGÜ

			for (int i = 0; i < dtExcel.Rows.Count; i++)
			{//EĞER KULLANICI IDSİ SİSTEMDE KAYITLI DEĞİLSE İŞLEMLERİN ES GEÇMESİ İÇİN OLUŞTURULAN KOŞUL

				DataColumnCollection excelColumns = dtExcel.Columns;
				bool validation = false;
				if (SecondValidation != null)
				{
					validation = SecondValidation(info, i, excelColumns, logRecord, dtExcel, secondParameters);
				}

				if (validation) continue;
				string formName = info.MainFormName;
				int formId = ebacon.WorkflowManager.CreateDocument(info.ProjectName, formName).DocumentId;

				InsertForm(info, globalAttributes, formColumns, formId, dtExcel.Rows[i], Connection, globalId, AdditionalUpdate, AdditionalQuery, dictTextIds);


				logRecord.ChangedFieldsNote += Environment.NewLine + (i + 4).ToString() + ". Satir Basariyla Ice Aktarildi ; ";

			}
			Connection.Close();
		}
		void InsertForm(Info info,
			   GlobalAttributes globalAttributes,
				List<string> columns,
			   int formId,
			   DataRow drUpdate,
			   SqlConnection Connection,
			   string globalId,
			   Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
			   Func<List<string>, DataRow, string, string> AdditionalQuery,
			   Dictionary<string, Dictionary<string, string>> dictTextIds)
		{
			string mainFormId = formId.ToString();
			string setQuery = "";
			foreach (string column in columns)
			{//DEĞİŞMEMESİ GEREKEN ALANLAR VARSA BU ALANLAR QUERYE EKLENMEZ
				if (column == "ID" || column == "docGlobalId" || globalAttributes.RestrictedFields.Contains(column)) continue;
				else if (column.StartsWith("mtl") || column.StartsWith("cmb"))
				{
					//Bazı alanlar dropdownList olarak tutuluyor ve bu alanlarda seçim yapılmadıysa idlerinin -1 olması gerekli bu sebeple null değerlerde -1 atanıyor,
					//comboboxlarda idlerin -1 olması bir hataya yol açmadığı için onlar için ayrı bir validasyon kullanılmıyor.
					/*
					if (string.Equals(drMain[column].ToString(), "-1", StringComparison.CurrentCultureIgnoreCase))
						setQuery += (drUpdate[column] == DBNull.Value) ?
							(column + "_TEXT = '' ," + column + "='-1',") :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + dictTextIds[column][drUpdate[column].ToString()] + "',";
					else
					*/
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
				setQuery = AdditionalUpdate(columns, drUpdate, setQuery, mainFormId);
			}

			if (setQuery[setQuery.Length - 1] == ',')
			{

				setQuery = setQuery.Remove(setQuery.Length - 1);

			}
			var comm = string.Format(@"UPDATE E_{3}_{2}
                                        SET {0}
                                        WHERE ID = {1}
										{4}
										/*
                                        SELECT *
                                        FROM E_{3}_{2}
                                        WHERE ID = {1} */
                                        ", setQuery, mainFormId, info.MainFormName, info.ProjectName, AdditionalQuery != null ? AdditionalQuery(columns, drUpdate, mainFormId) : " ");
			//SqlCommand Command = new SqlCommand();
			//SqlDataAdapter Adapter = new SqlDataAdapter();
			SqlCalls.RunQueriesNoReturn(comm, Connection);
			//return dt;
		}

		bool IsFormActive(Info info, InfoLog logRecord, DataRow dr, SqlConnection Connection, string formName)
		{
			DataTable dt;
			var comm = string.Format(@"SELECT Top(1) [CHECKOUTDATE] AS 'DATE' ,USERID
									  FROM  [dbo].[CHECKOUTS] 
									  WHERE FILENAME LIKE '%/{0}.%'
									  ORDER BY CHECKOUTDATE DESC", dr["ID"].ToString());

			dt = SqlCalls.RunQueries(comm, Connection);
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
		DataTable UpdateForm(Info info,
			   GlobalAttributes globalAttributes,
				List<string> columns,
			   DataRow drMain,
			   DataRow drUpdate,
			   SqlConnection Connection,
			   string globalId,
			   Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
			   Func<List<string>, DataRow, string, string> AdditionalQuery,
			   Dictionary<string, Dictionary<string, string>> dictTextIds)
		{
			string mainFormId = drMain["ID"].ToString();
			string setQuery = "";
			foreach (string column in columns)
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
				setQuery = AdditionalUpdate(columns, drUpdate, setQuery, globalId);
			}

			if (setQuery[setQuery.Length - 1] == ',')
			{

				setQuery = setQuery.Remove(setQuery.Length - 1);

			}
			DataTable dt;
			var comm = string.Format(@"UPDATE E_{3}_{2}
                                        SET {0}
                                        WHERE ID = {1}
										{4}
                                        SELECT *
                                        FROM E_{3}_{2}
                                        WHERE ID = {1}
                                        ", setQuery, mainFormId, info.MainFormName, info.ProjectName, AdditionalQuery != null ? AdditionalQuery(columns, drUpdate, mainFormId) : " ");
			dt = SqlCalls.RunQueries(comm, Connection);
			return dt;
		}

		void InsertLogRow(InfoLog logRecord,
			   DataRow drMain,
			   DataRow drExcel,
			   List<string> realColumns,
			   GlobalAttributes globalAttributes)
		{
			foreach (string columnsInfo in realColumns)
			{
				string column = columnsInfo;
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
						logRecord.ChangedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}

				}
				else if (column.StartsWith("txt"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
					{
						logRecord.ChangedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}
				}
				else if (column.StartsWith("mtl") || column.StartsWith("cmb"))
				{
					if (string.Compare(drExcel[column].ToString(), drMain[column].ToString()) != 0)
					{
						logRecord.ChangedFieldsNote += column.Substring(3) + "(" + drMain[column].ToString() + "->" + drExcel[column].ToString() + ")--";
					}
				}
			}

		}

	}

}
