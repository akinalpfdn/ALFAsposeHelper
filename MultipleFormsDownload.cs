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
{
	/// <summary>
	/// log sistemi icin formun indirilmesini saglayan class
	/// eger update degilde import yapilacaksa dictPersonel empty dataTable gonderilebilir
	/// </summary>
	public class MultipleFormsDownload
	{

		/// <summary>
		/// Ebadaki Log sistemi mevcut modullerdeki  bilgilerin update import icin hazirlanan templatei olusturur
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="excel">exceldeki satirlarin stylelari icin aspose.cell stylerini iceren class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="language">Turkish English gibi kullanilan dil belirtilmeli</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBar">dllde hata verdigi icin ebadaki ShowMessageBar methodu parametre olarak verilir</param>
		/// <param name="dictPersonnels">Update Yapilacaksa update edilecek sql satirlarinin unique edilerini iceren bir datatable sistemin global idsi ile dictionarye eklenmeli</param>
		/// <param name="multiLanguage">Lookupsta multilanguage varsa true degilse false</param>
		/// <param name="labelMultiLanguage">Labellarda multilanguage varsa true degilse false</param>
		public void Download(Info info,
			Excel excel,
			   GlobalAttributes globalAttributes,
			   string language,
			   string globalId,
			   Func<eBADBProvider> CreateDatabaseProvider,
			   Action<string> ShowMessageBox,
			   Action<Stream, string> WriteToResponse,
			   Action<string, int, ShowInfoBarType> ShowMessageBar,
			   Dictionary<string, DataTable> dictPersonnels,
			   bool multiLanguage,
			bool labelMultiLanguage)
		{
			Execute(info, excel, globalAttributes, language, globalId, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, ShowMessageBar, dictPersonnels, multiLanguage, labelMultiLanguage, null, null);

		}
		/// <summary>
		/// Ebadaki Log sistemi mevcut modullerdeki  bilgilerin update import icin hazirlanan templatei olusturur
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="excel">exceldeki satirlarin stylelari icin aspose.cell stylerini iceren class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="language">Turkish English gibi kullanilan dil belirtilmeli</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBar">dllde hata verdigi icin ebadaki ShowMessageBar methodu parametre olarak verilir</param>
		/// <param name="dictPersonnels">Update Yapilacaksa update edilecek sql satirlarinin unique edilerini iceren bir datatable sistemin global idsi ile dictionarye eklenmeli</param>
		/// <param name="multiLanguage">Lookupsta multilanguage varsa true degilse false</param>
		/// <param name="labelMultiLanguage">Labellarda multilanguage varsa true degilse false</param>
		/// <param name="dtFields">Alanlar Dinamik secilmeyecekse string list halinde verilmeleri gerekir</param>
		public void Download(Info info,
			Excel excel,
			   GlobalAttributes globalAttributes,
			  string language,
			  string globalId,
			  Func<eBADBProvider> CreateDatabaseProvider,
			  Action<string> ShowMessageBox,
			  Action<Stream, string> WriteToResponse,
			  Action<string, int, ShowInfoBarType> ShowMessageBar,
			  Dictionary<string, DataTable> dictPersonnels,
			  bool multiLanguage,
			bool labelMultiLanguage,
			List<string> dtFields)
		{
			Execute(info, excel, globalAttributes, language, globalId, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, ShowMessageBar, dictPersonnels, multiLanguage, labelMultiLanguage, dtFields, null);

		}
		/// <summary>
		/// Ebadaki Log sistemi mevcut modullerdeki  bilgilerin update import icin hazirlanan templatei olusturur
		/// </summary>
		/// <param name="info">proje hakkindaki bilgilerin tutuldugu class</param>
		/// <param name="excel">exceldeki satirlarin stylelari icin aspose.cell stylerini iceren class</param>
		/// <param name="globalAttributes">form isimleri, yasakli alanlar gibi verilerin tutuldugu class</param>
		/// <param name="language">Turkish English gibi kullanilan dil belirtilmeli</param>
		/// <param name="globalId">formun Unique IDsi</param>
		/// <param name="CreateDatabaseProvider">dllde hata verdigi icin ebadaki CreateDatabaseProvider methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBox">dllde hata verdigi icin ebadaki ShowMessageBox methodu parametre olarak verilir</param>
		/// <param name="WriteToResponse">dllde hata verdigi icin ebadaki WriteToResponse methodu parametre olarak verilir</param>
		/// <param name="ShowMessageBar">dllde hata verdigi icin ebadaki ShowMessageBar methodu parametre olarak verilir</param>
		/// <param name="dictPersonnels">Update Yapilacaksa update edilecek sql satirlarinin unique edilerini iceren bir datatable sistemin global idsi ile dictionarye eklenmeli</param>
		/// <param name="multiLanguage">Lookupsta multilanguage varsa true degilse false</param>
		/// <param name="labelMultiLanguage">Labellarda multilanguage varsa true degilse false</param>
		/// <param name="dtFields">Alanlar Dinamik secilmeyecekse string list halinde verilmeleri gerekir</param>
		/// <param name="passColumns">Bazi projelere ozel olarak bazi sartlarda kolonlarin gecilmesi gerekebilir, bu gibi durumlarda bu method kullanilmalidir</param>
		public void Download(Info info,
			Excel excel,
			   GlobalAttributes globalAttributes,
			string language,
			string globalId,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string> ShowMessageBox,
			Action<Stream, string> WriteToResponse,
			Action<string, int, ShowInfoBarType> ShowMessageBar,
			Dictionary<string, DataTable> dictPersonnels,
			bool multiLanguage,
			bool labelMultiLanguage,
			List<string> dtFields,
			Func<string, string, Boolean> passColumns)
		{
			Execute(info, excel, globalAttributes, language, globalId, CreateDatabaseProvider, ShowMessageBox, WriteToResponse, ShowMessageBar, dictPersonnels, multiLanguage, labelMultiLanguage, dtFields, passColumns);

		}

		/// <summary>
		/// Seçilen personellerin, seçilen alanlardaki bilgileriyle dolu bir excel üreten method <br/><br/>
		/// dtFields >> ilgili kolonların isimlerini barındıran List <br/>
		/// langugage >> dillerin dropdown listteki kodu (1,2,3)<br/>
		/// logonUser >> user ID <br/> 
		/// globalId >> Formun Global IDsi <br/>
		/// passcolumns ilk parametre olarak formun adini ve ikinci parametre olarak kolonun adini alir boylece custom kurallar ile o kolonun gecilip gecilmeyecegi belirlenir
		/// </summary>
		public void Execute(Info info,
			Excel excel,
			   GlobalAttributes globalAttributes,
			string language,
			string globalId,
			Func<eBADBProvider> CreateDatabaseProvider,
			Action<string> ShowMessageBox,
			Action<Stream, string> WriteToResponse,
			Action<string, int, ShowInfoBarType> ShowMessageBar,
			Dictionary<string, DataTable> dictPersonnels,
			bool multiLanguage,
			bool labelMultiLanguage,
			List<string> dtFields = null,
			Func<string, string, Boolean> passColumns = null)
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
			List<Tuple<string, string, Type>> columnsOld = Fields.TakeFields(info, "download", dtFields, CreateDatabaseProvider, globalAttributes, passColumns);//FORMLARDAKİ ALANLARIN BAŞLIKLARINI TOPLAYAN METHOD 
			List<Tuple<string, string, Type>> columns = new List<Tuple<string, string, Type>>();
			Dictionary<string, DataRow> dictAllMtl = Fields.GetAllMtl(info, CreateDatabaseProvider);
			if (dtFields == null)
			{
				columns = columnsOld;
			}
			else
			{
				foreach (string column in dtFields)
				{
					foreach (Tuple<string, string, Type> columnDetail in columnsOld)
					{
						if (columnDetail.Item1 == column)
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
			if (dictPersonnels[globalId].Rows.Count == 0)
			{
				ShowMessageBox(info.NotEnoughObjectMessage);
			}
			else
			{
				foreach (DataRow dr in dictPersonnels[globalId].Rows)
				{
					uniqueIds += "'" + dr[0].ToString() + "',";
				}
				if (dictPersonnels[globalId].Rows.Count != 0) uniqueIds = uniqueIds.Remove(uniqueIds.Length - 1);
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
				indexes = FillDefinedColumns(i, j, uniqueIds, comm, language, langCode, info, globalAttributes, Connection, ws, excel, columns, dictAllMtl, labelMultiLanguage);
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
			string language,
			string langCode,
			Info info,
			GlobalAttributes globalAttributes,
			SqlConnection Connection,
			 Worksheet ws, Excel excel,
			List<Tuple<string, string, Type>> columns,
			Dictionary<string, DataRow> dictAllMtl,
			bool multiLanguage)
		{

			DataTable dtLabel = new DataTable();
			foreach (Tuple<string, string, Type> column in columns)
			{
				DataTable dt = new DataTable();


				ws.Cells[0, i].PutValue(column.Item1);
				ws.Cells[1, i].PutValue(column.Item2);


				if (multiLanguage)
				{
					comm = string.Format(@"SELECT Language, Text
												  FROM [MLDICTIONARIES]
												 WHERE WORD = (SELECT  TOP(1) WORD
																 FROM [dbo].[MTDTDBFIELDS] FRM WITH(NOLOCK)
																INNER JOIN[MLDICTIONARIES] DICT WITH(NOLOCK) ON DICT.WORD like '%' + FRM.FORM + '%' AND cast(DICT.TEXT as nvarchar(max)) = DESCRIPTION
																WHERE OBJECTNAME = '{0}')", column.Item1);// excelde başlıkların MultiLanguage olması için

					dtLabel = SqlCalls.RunQueries(comm, Connection);
					DataRow drLabel = dtLabel
						.AsEnumerable()
						.FirstOrDefault(r => r.Field<string>("Language") == language);

					if (drLabel == null) drLabel = dtLabel
						  .AsEnumerable()
						  .FirstOrDefault(r => r.Field<string>("Language") == "English");

					try
					{
						ws.Cells[2, i].PutValue(drLabel["Text"].ToString().Replace(":", ""));
						ws.Cells[2, i].SetStyle(excel.Labelstyle);
					}
					catch (NullReferenceException ex)//eger bir sekilde multilanguage calismiyorsa
					{
						try
						{
							comm = string.Format(@"SELECT TOP(1) DESCRIPTION FROM [MTDTDBFIELDS]
																WHERE PROJECT = '{0}' AND FORM = '{1}' AND OBJECTNAME = '{2}'", info.ProjectName, column.Item2, column.Item1);// excelde başlıkların MultiLanguage olması için

							dtLabel = SqlCalls.RunQueries(comm, Connection);

							ws.Cells[2, i].PutValue(dtLabel.Rows[0][0].ToString().Replace(":", ""));

							ws.Cells[2, i].SetStyle(excel.Labelstyle);
						}
						catch
						{

							throw new Exception(comm);
						}
					}
				}
				else
				{
					try
					{
						comm = string.Format(@"SELECT TOP(1) DESCRIPTION FROM [MTDTDBFIELDS]
																WHERE PROJECT = '{0}' AND FORM = '{1}' AND OBJECTNAME = '{2}'", info.ProjectName, column.Item2, column.Item1);// excelde başlıkların MultiLanguage olması için

						dtLabel = SqlCalls.RunQueries(comm, Connection);

						ws.Cells[2, i].PutValue(dtLabel.Rows[0][0].ToString().Replace(":", ""));

						ws.Cells[2, i].SetStyle(excel.Labelstyle);
					}
					catch
					{

						throw new Exception(comm);
					}
				}
				//i += 1;
				if (info.UniqueColumnInfo != null)
				{
					string select = @"MDL." + column.Item1 + (column.Item1.StartsWith("cmb") ? "_TEXT" : "");
					try
					{
						comm = string.Format(@"SELECT {0}
												FROM E_{4}_{5} FRM WITH(NOLOCK)
												LEFT JOIN E_{4}_{1} MDL WITH(NOLOCK) ON MDl.ID = FRM.{2}
												WHERE FRM.{6} IN ({3 })
												ORDER BY FRM.{6}", select, column.Item2, globalAttributes.DictFormTextBoxes[column.Item2][0], uniqueIds, info.ProjectName, info.MainFormName, info.UniqueColumnInfo.Item1);
						dt = new DataTable();
						dt = SqlCalls.RunQueries(comm, Connection);
						if (column.Item1.StartsWith("mtl")) //Eğer Multi Language Alan ise formda seçilen dile göre getirilmeli o alandaki veriler
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
