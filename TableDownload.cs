using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using eBAControls.eBABaseForm;
using eBADB;
using eBADocumentBuilder;
using eBAPI.Connection;
using eBAPI.DocumentManagement;

namespace ALFAsposeHelper
{/// <summary>
 /// 
 /// </summary>
	public class TableDownload
	{
		/// <summary>
		/// Sqldeki bir tablonun  bilgilerin update import icin hazirlanan templatei olusturur
		/// </summary>
		/// 
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="excel">exceldeki satirlarin stylelari icin aspose.cell stylerini iceren class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="language">Turkish English gibi kullanilan dil belirtilmeli</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBar">dllde hata verdigi icin ebadaki ShowMessageBar methodu parametre olarak verilir</param>
		/// <param name="dictPersonnels">Update Yapilacaksa update edilecek sql satirlarinin unique edilerini iceren bir datatable sistemin global idsi ile dictionarye eklenmeli</param>
		/// <param name="multiLanguage">Lookupta multilanguage varsa true degilse false</param>
		/// <param name="dictLabel">excelde basliklarin isimlendirilmesi icin gerekli dictionary</param>
		/// <param name="hasLookUp">tabloda mtl alan varsa true degilse false</param>
		public void Download(Info info,
		  GlobalAttributes globalAttributes,
				  Excel excel,
		  string language,
		  string globalId,
		  Func<eBADBProvider> CreateDatabaseProvider,
		  Action<Stream, string> WriteToResponse,
		  Action<string, int, ShowInfoBarType> ShowMessageBar,
		  Dictionary<string, DataTable> dictPersonnels,
			Dictionary<string, string> dictLabel,
		  bool multiLanguage,
			   bool hasLookUp)
		{
			Execute(info, globalAttributes, excel, language, globalId, CreateDatabaseProvider, WriteToResponse, ShowMessageBar, dictPersonnels, dictLabel, multiLanguage, hasLookUp, null);

		}
		/// <summary>
		/// Sqldeki bir tablonun  bilgilerin update import icin hazirlanan templatei olusturur
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="excel">exceldeki satirlarin stylelari icin aspose.cell stylerini iceren class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="language">Turkish English gibi kullanilan dil belirtilmeli</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBar">dllde hata verdigi icin ebadaki ShowMessageBar methodu parametre olarak verilir</param>
		/// <param name="dictPersonnels">Update Yapilacaksa update edilecek sql satirlarinin unique edilerini iceren bir datatable sistemin global idsi ile dictionarye eklenmeli</param>
		/// <param name="multiLanguage">Lookupta multilanguage varsa true degilse false</param>
		/// <param name="dictLabel">excelde basliklarin isimlendirilmesi icin gerekli dictionary</param>
		/// <param name="hasLookUp">tabloda mtl alan varsa true degilse false</param>
		/// <param name="dtFields">eger sadece belirli alanlar indirilmek isteniliyorsa o alanlari iceren list</param>
		public void Download(Info info,
			   GlobalAttributes globalAttributes,
			   Excel excel,
			string language,
			string globalId,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<Stream, string> WriteToResponse,
			Action<string, int, ShowInfoBarType> ShowMessageBar,
			Dictionary<string, DataTable> dictPersonnels,
			Dictionary<string, string> dictLabel,
			bool multiLanguage,
			bool hasLookUp,
			List<string> dtFields)
		{
			Execute(info, globalAttributes, excel, language, globalId, CreateDatabaseProvider, WriteToResponse, ShowMessageBar, dictPersonnels, dictLabel, multiLanguage, hasLookUp, dtFields);

		}

		void Execute(Info info,
			   GlobalAttributes globalAttributes,
			   Excel excel,
			string language,
			string globalId,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<Stream, string> WriteToResponse,
			Action<string, int, ShowInfoBarType> ShowMessageBar,
			Dictionary<string, DataTable> dictPersonnels,
			Dictionary<string, string> dictLabel,
			bool multiLanguage,
			bool hasLookUp,
			List<string> dtFields = null)
		{
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
			List<string> columnsOld = (List<string>)Fields.TakeFieldsForTable(info, dtFields, CreateDatabaseProvider, globalAttributes)[0];//FORMLARDAKİ ALANLARIN BAŞLIKLARINI TOPLAYAN METHOD 
			List<string> columns = new List<string>();
			Dictionary<string, DataRow> dictAllMtl = null;
			if (hasLookUp)
			{
				dictAllMtl = Fields.GetAllMtl(info, CreateDatabaseProvider);
			}
			if (dtFields == null)
			{
				columns = columnsOld;
			}
			else
			{
				foreach (string column in dtFields)
				{
					foreach (string columnDetail in columnsOld)
					{
						if (columnDetail == column)
						{
							columns.Add(columnDetail);
							break;
						}
					}
				}
			}

			List<KeyValuePair<string[], DataTable>> mtlFields = (List<KeyValuePair<string[], DataTable>>)Fields.TakeMtlValues(info, columns, langCode, CreateDatabaseProvider, Fields.AddDataTable, new List<KeyValuePair<string[], DataTable>>(), "", multiLanguage); //Excel validasyonu için mtl alanların verilerini barındıran List
			string uniqueIds = "";
			//if (tblPersonnels.Data.Rows.Count > 0)
			//if (dictPersonnels[globalId].Rows.Count == 0)
			{
				//ShowMessageBox(info.notEnoughObjectMessage);
			}
			//else
			{
				if (info.DownloadType != Info.EnumDownloadType.Import)
				{
					foreach (DataRow dr in dictPersonnels[globalId].Rows)
					{
						uniqueIds += "'" + dr[0].ToString() + "',";
					}
					if (dictPersonnels[globalId].Rows.Count != 0) uniqueIds = uniqueIds.Remove(uniqueIds.Length - 1);
				}

				Workbook wb = new Workbook();

				string sheetName = "Data";
				Worksheet ws = wb.Worksheets[0];
				ws.Name = sheetName;
				DataTable dt = new DataTable();
				int i = 0;
				int j = 0;

				string comm = "";
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				Connection.Open();
				var indexes = Fields.FillPredefinedColumns(i, j, info, uniqueIds, Connection, ws, excel);
				i = indexes[0];
				j = indexes[1];

				CultureInfo provider = CultureInfo.InvariantCulture;
				indexes = FillDefinedColumns(i, j, uniqueIds, comm, langCode, info, Connection, ws, excel, columns, dictAllMtl, dictLabel, multiLanguage);
				i = indexes[0];
				j = indexes[1];
				ws.Cells.HideRow(0);
				ws.Cells.HideRow(1);
				sheetName = "Data Validation";
				wb.Worksheets.Add();
				Worksheet ws2 = wb.Worksheets[1];
				ws2.Name = sheetName;

				foreach (KeyValuePair<string[], DataTable> mtlField in mtlFields) //Validasyon için Exceldeki Gizli Sheete mtl alanların verilerini ekliyoruz
				{

					for (int col = 0; col < i; col++)
					{
						if (ws.Cells[0, col].Value != null && ws.Cells[0, col].Value.ToString() == mtlField.Key[0])
						{

							ws2.Cells[0, j].PutValue(ws.Cells[2, col].Value);
							ws2.Cells[0, j].SetStyle(excel.Labelstyle);
						}
					}
					for (int row = 0; row < mtlField.Value.Rows.Count; row++)
					{
						ws2.Cells[row + 1, j].PutValue(mtlField.Value.Rows[row][0]);
						ws2.Cells[row + 1, j].SetStyle(excel.CellStyle);
					}
					j++;

				}



				ws.AutoFitColumns();
				ws2.AutoFitColumns();
				using (Stream respStream = new MemoryStream())
				{
					wb.Save(respStream, Aspose.Cells.SaveFormat.Xlsx);
					respStream.Seek(0, SeekOrigin.Begin);
					WriteToResponse(respStream, info.DownloadName);
				}

				ShowMessageBar("You Have Succcesfully Downloaded The Excel File", 3000, ShowInfoBarType.Success);
			}
		}



		static List<int> FillDefinedColumns(int i,
			int j,
			string uniqueIds,
			string comm,
			string langCode,
			Info info,
			SqlConnection Connection,
			 Worksheet ws, Excel excel,
			List<string> columns,
			Dictionary<string, DataRow> dictAllMtl,
			Dictionary<string, string> dictLabel,
			bool multiLanguage)
		{

			foreach (string column in columns)
			{
				DataTable dt = new DataTable();


				ws.Cells[0, i].PutValue(column);
				ws.Cells[1, i].PutValue(info.UniqueColumnInfo.Item2);

				ws.Cells[2, i].PutValue(dictLabel[column]);

				ws.Cells[2, i].SetStyle(excel.Labelstyle);


				if (info.DownloadType != Info.EnumDownloadType.Import)
				{
					string select = @"FRM." + column + (column.StartsWith("cmb") ? "_TEXT" : "");
					try
					{
						comm = string.Format(@"SELECT {0}
												FROM {1} FRM WITH(NOLOCK)
												WHERE FRM.{3} IN ({2 })
												ORDER BY FRM.{3}", select, info.MainFormName, uniqueIds, info.UniqueColumnInfo.Item1);
						dt = new DataTable();
						dt = SqlCalls.RunQueries(comm, Connection);
						if (column.StartsWith("mtl")) //Eğer Multi Language Alan ise formda seçilen dile göre getirilmeli o alandaki veriler
						{
							DataTable dtTemp = new DataTable();
							int index = 0;
							switch (langCode)
							{
								case "TR":
									index = 1;
									break;
								case "EN":
									index = 2;
									break;
								case "RU":
									index = 3;
									break;
								default:
									index = 1;
									break;
							}
							if (!multiLanguage) index = 1;
							dtTemp.Columns.Add("value", typeof(string));

							foreach (DataRow dr in dt.Rows)
							{
								DataRow drTemp = dtTemp.NewRow();
								if (dr[0] != DBNull.Value && dr[0].ToString() != "")
								{
									drTemp[0] = dictAllMtl[dr[0].ToString()][index];
								}
								dtTemp.Rows.Add(drTemp);
							}
							dt = dtTemp;
						}

					}
					catch
					{
						throw new Exception(comm);

					}
					for (int row = 0; row < dt.Rows.Count; row++)
					{
						ws.Cells[row + 3, i].PutValue(dt.Rows[row][0]);

						ws.Cells[row + 3, i].SetStyle(excel.CellStyle);
						if (dt.Columns[0].DataType == typeof(DateTime))
						{
							var DateStyle = ws.Cells[row + 3, i].GetStyle();
							DateStyle.Custom = "dd/mm/yyyy";
							ws.Cells[row + 3, i].SetStyle(DateStyle);

						}
					}
				}
				i++;

			}
			return new List<int>() { i, j };

		}
	}
}
