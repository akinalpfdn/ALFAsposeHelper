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
	/// Sqldeki bir tabloya verileri yuklemek ve ya guncellemek icin kullanilir
	/// </summary>
	public class TableUpload
	{
		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, null, null, multiLanguage);
		}
		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
		/// <param name="AdditionalQuery"></param>
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
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, null, null, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			bool multiLanguage)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, null, null, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
		/// <param name="AdditionalQuery"></param>
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
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			object[] parameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, null, parameters, null, AdditionalUpdate, AdditionalQuery, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, null, null, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
		/// <param name="AdditionalQuery"></param>
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
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, null, SecondValidation, null, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
		/// <param name="AdditionalQuery"></param>
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
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
		}

		/// <summary>
		/// Excel Uzerinden sqldeki bir tabloya update ya da import icin kullanilir
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
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			bool multiLanguage
			)
		{
			Excute(info, globalAttributes, attUpload, logonUser, lang, globalId, CreateServerConnection, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, Validations, SecondValidation, parameters, secondParameters, null, null, multiLanguage);
		}

		private void Excute(Info info,
			GlobalAttributes globalAttributes, eBAAttachments attUpload,
			string logonUser,
			string language,
			string globalId,
			Func<eBAConnection> CreateServerConnection,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
				List<string>> Validations,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] parameters,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
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
				attUpload.ReadOnly = true; //Bir kez dosya yüklendikten sonra bir daha yüklenemesin diye readonly yapılıyor.
				List<string> columns = (List<string>)Fields.TakeFieldsForTable(info, null, CreateDatabaseProvider, globalAttributes)[0];
				List<Tuple<string, Type>> columnDetails = (List<Tuple<string, Type>>)Fields.TakeFieldsForTable(info, null, CreateDatabaseProvider, globalAttributes)[1];

				DataTable excelTable = new DataTable();
				DataTable excelTableTemp = new DataTable(); // Excelden gelen verileri typeları ile çekmek için geçiçi bir template table oluşturuluyor
				eBAConnection ebacon = CreateServerConnection();
				Stream data = null;
				ebacon.Open();
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				Connection.Open();
				FileSystem fs = ebacon.FileSystem;
				DMFile file = fs.GetFile(attUpload.Filename);
				DMCategoryContentCollection flist = file.GetAttachments(info.AttachmentCategory);
				if (flist.Count > 0)
				{
					List<string> formColumns = new List<string>();

					for (int a = 0; a < 1; a++)
					{
						DMFileContent att = flist[a];
						data = fs.CreateFileAttachmentContentDownloadStream(attUpload.Filename, info.AttachmentCategory, att.ContentName);
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
						var columnsTemp = new List<Tuple<string, Type>>();
						foreach (DataColumn cl in excelTableTemp.Columns)
						{
							foreach (Tuple<string, Type> column in columnDetails)
							{
								string columnName = column.Item1;
								Type columnType = column.Item2;
								if (cl.ColumnName == columnName)
								{
									columnsTemp.Add(column);
									//excelTable.Columns.Add(cl.ColumnName, !(column.Value.Equals(typeof(DateTime))) ? typeof(string) : typeof(DateTime));
									if (columnType.Equals(typeof(DateTime))) excelTable.Columns.Add(columnName, typeof(DateTime));
									else if (columnType.Equals(typeof(Decimal))) excelTable.Columns.Add(columnName, typeof(Decimal));
									else if (columnType.Equals(typeof(Double))) excelTable.Columns.Add(columnName, typeof(Double));
									else excelTable.Columns.Add(columnName, typeof(string));
								}

							}
						}
						columnDetails = columnsTemp;
						options.DataTable = excelTable;
						options.SkipErrorValue = true;
						Worksheet worksheet = mybook.Worksheets[0];
						try
						{
							excelTable = worksheet.Cells.ExportDataTable(0, 0, satirsay + 1, kolonsay + 1, options);
						}
						catch (IndexOutOfRangeException) // Yetkisiz alanlar var ise sistem index out of range hatası verecektir bu durumda ilerlenmemesi gerekir.
						{
							throw new Exception("Hatali bir excel yuklediniz, lütfen yeni bir template oluşturup tekrar deneyiniz!");
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
						idsToDelete = Validations(info, excelTable, CreateDatabaseProvider, columns, langCode, logRecord, textIdDict, parameters);
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
					}
					try
					{
						//ALFDebugHelper.Log(12121, excelTable,excelTable.Columns);
						UpdateInfo(info, globalAttributes, logRecord, excelTable, formColumns, columns, globalId, logonUser, langCode, CreateDatabaseProvider, SecondValidation, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
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
		/// sadece deneme amacli kullanilmalidir
		/// </summary>
		/// <param name="info"></param>
		/// <param name="globalAttributes"></param>
		/// <param name="logonUser"></param>
		/// <param name="language"></param>
		/// <param name="globalId"></param>
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
		public void Excute(Info info,
			GlobalAttributes globalAttributes,
			string logonUser,
			string language,
			string globalId,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string, eBAMessageBoxType> ShowMessageBox,
			Action<Byte[], string> WriteToResponse,
			Func<Info, DataTable, Func<eBADBProvider>, List<string>, string, InfoLog, Dictionary<string, Dictionary<string, string>>, object[],
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
				{
					List<string> columns = (List<string>)Fields.TakeFieldsForTable(info, null, CreateDatabaseProvider, globalAttributes)[0];
					List<Tuple<string, Type>> columnDetails = (List<Tuple<string, Type>>)Fields.TakeFieldsForTable(info, null, CreateDatabaseProvider, globalAttributes)[1];

					eBADB.eBADBProvider db = CreateDatabaseProvider();
					SqlConnection Connection = (SqlConnection)db.Connection;
					Connection.Open();
					DataTable excelTable = new DataTable();
					DataTable excelTableTemp = new DataTable(); // Excelden gelen verileri typeları ile çekmek için geçiçi bir template table oluşturuluyor
					List<string> formColumns = new List<string>();
					Workbook mybook = new Workbook(excelStream);
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
					var columnsTemp = new List<Tuple<string, Type>>();
					foreach (DataColumn cl in excelTableTemp.Columns)
					{
						foreach (Tuple<string, Type> column in columnDetails)
						{
							string columnName = column.Item1;
							Type columnType = column.Item2;
							if (cl.ColumnName == columnName)
							{
								columnsTemp.Add(column);
								//excelTable.Columns.Add(cl.ColumnName, !(column.Value.Equals(typeof(DateTime))) ? typeof(string) : typeof(DateTime));
								if (columnType.Equals(typeof(DateTime))) excelTable.Columns.Add(columnName, typeof(DateTime));
								else if (columnType.Equals(typeof(Decimal))) excelTable.Columns.Add(columnName, typeof(Decimal));
								else if (columnType.Equals(typeof(Double))) excelTable.Columns.Add(columnName, typeof(Double));
								else excelTable.Columns.Add(columnName, typeof(string));
							}

						}
					}
					columnDetails = columnsTemp;
					options.DataTable = excelTable;
					options.SkipErrorValue = true;
					Worksheet worksheet = mybook.Worksheets[0];
					try
					{
						excelTable = worksheet.Cells.ExportDataTable(0, 0, satirsay + 1, kolonsay + 1, options);
					}
					catch (IndexOutOfRangeException) // Yetkisiz alanlar var ise sistem index out of range hatası verecektir bu durumda ilerlenmemesi gerekir.
					{
						throw new Exception("Hatali bir excel yuklediniz, lütfen yeni bir template oluşturup tekrar deneyiniz!");
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
					excelTable.Rows.RemoveAt(0);
					excelTable.Rows.RemoveAt(0);
					excelTableTemp.Rows.RemoveAt(0);
					excelTableTemp.Rows.RemoveAt(0);
					//VALİDASYON KONTROLLERİ
					List<string> idsToDelete = new List<string>(); ; ; ;
					Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)Fields.TakeMtlValues(info, columns, langCode, CreateDatabaseProvider, Fields.AddDictionary, new Dictionary<string, DataTable>(), "", multiLanguage);
					idsToDelete = Fields.CellValidations(globalAttributes, excelTable, excelTableTemp, logRecord, dictFields);
					Dictionary<string, Dictionary<string, string>> textIdDict = Fields.TakeMtlDictionary(info, columns, langCode, Connection, multiLanguage);
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
						idsToDelete = Validations(info, excelTable, CreateDatabaseProvider, columns, langCode, logRecord, textIdDict, parameters);
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
					}

					try
					{
						//ALFDebugHelper.Log(12121, excelTable,excelTable.Columns);
						UpdateInfo(info, globalAttributes, logRecord, excelTable, formColumns, columns, globalId, logonUser, langCode, CreateDatabaseProvider, SecondValidation, secondParameters, AdditionalUpdate, AdditionalQuery, multiLanguage);
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
			 List<string> formColumns,
			List<string> realColumns,
			string globalId,
			string logonUser,
			string langCode,
			Func<eBADBProvider> CreateDatabaseProvider,
			Func<Info, int, DataColumnCollection, InfoLog, DataTable, object[], bool> SecondValidation,
			object[] secondParameters,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{


			string uniqueIds = "";
			string uniqueColumnName = info.UniqueColumnInfo.Item1;
			//EXCELDEN GENEL PERSONNELLERİN TÜM VERİLERİNİ TEK SEFERDE ALMAK İÇİN IDLERİNİ BİRLEŞTİREN DÖNGÜ
			for (int i = 0; i < dtExcel.Rows.Count; i++)
			{
				uniqueIds += "'" + dtExcel.Rows[i][0].ToString() + "',";
			}
			uniqueIds += "'111'";
			DataTable dtForms = new DataTable();
			var comm = string.Format(@"SELECT FRM.*
                                         FROM {0} FRM WITH(NOLOCK)
                                        WHERE {1} in ({2})
                                        ", info.MainFormName, uniqueColumnName, uniqueIds);
			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
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

				//PERSONNELE AİT MODAL FORMU GETİREN DATATABLE EXPRESSİON
				DataRow drMainForm = dtForms
					.AsEnumerable()
					.FirstOrDefault(r => (dtExcel.Columns[uniqueColumnName].DataType == typeof(string) && r.Field<string>(uniqueColumnName) == dtExcel.Rows[i][uniqueColumnName].ToString())
					|| ((dtExcel.Columns[uniqueColumnName].DataType == typeof(Int32) || dtExcel.Columns[uniqueColumnName].DataType == typeof(double)) && r.Field<Int32>(uniqueColumnName).ToString() == dtExcel.Rows[i][uniqueColumnName].ToString()));
				//EĞER O FORMDA UPDATE EDİLECEK KOLON BULUNMUYORSA GERİ KALAN İŞLEMLERİN ES GEÇİŞLMESİ GEREKLİ

				DataTable dtMainModal2 = UpdateMainModal(info, logRecord, globalAttributes, formColumns, drMainForm, dtExcel.Rows[i], Connection, globalId, logonUser, langCode, AdditionalUpdate, AdditionalQuery, multiLanguage);

				InsertLogRow(logRecord, drMainForm, dtMainModal2.Rows[0], realColumns, globalAttributes);

			}
			Connection.Close();
		}

		/// <summary>
		/// Personel ve Form bazlı main formları güncelleyen method.
		/// </summary>
		/// <returns>Formların güncellenmiş halini geri döner</returns>
		private DataTable UpdateMainModal(Info info, InfoLog logRecord,
			GlobalAttributes globalAttributes,
			List<string> columns,
			DataRow drMain,
			DataRow drUpdate,
			SqlConnection Connection,
			string globalId,
			string logonUser,
			string langCode,
			Func<List<string>, DataRow, string, string, string> AdditionalUpdate,
				Func<List<string>, DataRow, string, string> AdditionalQuery,
			bool multiLanguage)
		{
			string mainFormId = drMain["ID"].ToString();
			string setQuery = "";
			Dictionary<string, Dictionary<string, string>> textIdDict = Fields.TakeMtlDictionary(info, columns, langCode, Connection, multiLanguage);
			foreach (string column in columns)
			{//DEĞİŞMEMESİ GEREKEN ALANLAR VARSA BU ALANLAR QUERYE EKLENMEZ
				if (column == "ID" || column == "docGlobalId" || globalAttributes.RestrictedFields.Contains(column)) continue;
				else if (column.StartsWith("mtl"))
				{
					//Bazı alanlar dropdownList olarak tutuluyor ve bu alanlarda seçim yapılmadıysa idlerinin -1 olması gerekli bu sebeple null değerlerde -1 atanıyor,
					//comboboxlarda idlerin -1 olması bir hataya yol açmadığı için onlar için ayrı bir validasyon kullanılmıyor.
					if (string.Equals(drMain[column].ToString(), "-1", StringComparison.CurrentCultureIgnoreCase))
						setQuery += (drUpdate[column] == DBNull.Value) ?
							(column + "_TEXT = '' ," + column + "='-1',") :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + textIdDict[column][drUpdate[column].ToString()] + "',";
					else
					{
						setQuery += (drUpdate[column] == DBNull.Value) ?
							column + "_TEXT = NULL ," + column + " = NULL ," :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + textIdDict[column][drUpdate[column].ToString()] + "',";
					}
					//if (drUpdate[column] == DBNull.Value)
					{
						//ALFeBAHelper.ALFDebugHelper.Log(23, column+": " +drMain[column].ToString());
					}
				}
				else if (column.StartsWith("cmb"))
				{
					setQuery += (drUpdate[column] == DBNull.Value) ?
							column + "_TEXT = NULL ," + column + " = NULL ," :
							column + "_TEXT = N'" + drUpdate[column].ToString() + "'," + column + " = N'" + textIdDict[column][drUpdate[column].ToString()] + "',";
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
			setQuery = setQuery.Remove(setQuery.Length - 1, 1);
			DataTable dt;
			var comm = string.Format(@"UPDATE {2}
                                        SET {0}
                                        WHERE ID = {1}
										{3}
                                        SELECT *
                                        FROM {2}
                                        WHERE ID = {1}
                                        ", setQuery, mainFormId, info.MainFormName, AdditionalQuery != null ? AdditionalQuery(columns, drUpdate, mainFormId) : " "
										);
			dt = SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
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
