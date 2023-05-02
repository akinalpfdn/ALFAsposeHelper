using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ALFeBAHelper;
using eBADB;
using eBADocumentBuilder;
using Aspose.Cells;

namespace ALFAsposeHelper
{
	class Fields
	{
		//excelde kullanılacak alanları dinamik bir şekilde toplayan method
		/// <summary>
		/// Excelde kullanılacak Kolonları dinamik bir şekilde toplayan method
		/// </summary>
		/// <returns>Sırası ile kolon adı, kolonun bulunduğu tablo ve kolonun tipini Barındıran Liste döner. </returns>
		public static List<Tuple<string, string, Type>> TakeFields(Info info,string type, List<string> dtFields,  Func<eBADBProvider> CreateDatabaseProvider, GlobalAttributes GlobalAttributes, Func<string, string, Boolean> passColumns =null )
		{
			List<string> tableColumns = new List<string>();
			if (type == "download" && dtFields != null)
			{
				foreach (string column in dtFields)
				{
					//if (dr["CHECKED"].ToString() == "1")
					{
						tableColumns.Add(column);
					}
				}
			}
			DataTable dtForm = new DataTable();
			if(info.AuthQuery==null)
			{
                dtForm = ALFIntegrationHelper.ExecuteIntegrationQuery(info.AuthConnectionName, info.AuthQueryName, info.AuthParameters); //Personelin sahip olduğu düzenleme yetkilerini getiren sorgu
            }
			else
            {
                dtForm = ExecuteIntegrationQuery(CreateDatabaseProvider, info.AuthQuery, info.AuthParameters); //Personelin sahip olduğu düzenleme yetkilerini getiren sorgu
            }
			if (dtForm.Rows.Count == 0) return null;
			List<string> allowedForms = new List<string>();
			foreach (DataRow dr in dtForm.Rows)
			{
				//if (dr["ID"].ToString() == "5954") continue;
				try { allowedForms.Add(GlobalAttributes.AuthForms[dr["ID"].ToString()]); }
				catch (KeyNotFoundException) { }
			}
			allowedForms.Add(info.MainFormName);
			List<Tuple<string, string, Type>> columns = new List<Tuple<string, string, Type>>();
			foreach (Tuple<string, string, string> form in GlobalAttributes.Forms)
			{
				if (!allowedForms.Contains(form.Item1)) continue;
				DataTable dt = new DataTable();
				var comm = string.Format(@"SELECT  *     
                                          FROM E_{0}_{1}
                                        ",info.ProjectName, form.Item1);
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				dt = SqlCalls.RunQueries(comm, Connection);
				Connection.Close();

				foreach (DataColumn column in dt.Columns)
				{
					if (form.Item1 == "FrmPersonnel" && (column.ColumnName == "txtPassportNo" || column.ColumnName == "txtTcNo")) continue;
					if (passColumns != null && passColumns(form.Item1, column.ColumnName)) continue;
					if (type == "download" && dtFields !=null)
					{
						if (tableColumns.Contains(column.ColumnName))
							columns.Add(new Tuple<string, string, Type>(column.ColumnName, form.Item1, column.DataType));
					}
					else
					{
						if (!GlobalAttributes.Fields.Contains(column.ColumnName) && !column.ColumnName.EndsWith("_TEXT"))
							columns.Add(new Tuple<string, string, Type>(column.ColumnName, form.Item1, column.DataType));
					}

				}

			}
			//ALFDebugHelper.Log(88, columns);
			return columns;
		}

		private static DataTable ExecuteIntegrationQuery(Func<eBADBProvider> CreateDatabaseProvider,string query, List<KeyValuePair<string, string>> authParameters)
		{

			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
			foreach (var parameter in authParameters)
			{
				query = query.Replace("<?=" + parameter.Key + ">", parameter.Value);

			}


			return SqlCalls.RunQueries(query, Connection); ;
		}

		public static List<Tuple<string, string, Type>> TakeFields(Info info,  List<string> dtFields, Func<eBADBProvider> CreateDatabaseProvider, GlobalAttributes GlobalAttributes)
		{
			List<string> tableColumns = new List<string>();
			if ( dtFields != null)
			{
				foreach (string column in dtFields)
				{
					//if (dr["CHECKED"].ToString() == "1")
					{
						tableColumns.Add(column);
					}
				}
			}
			List<Tuple<string, string, Type>> columns = new List<Tuple<string, string, Type>>();
				DataTable dt ;
				var comm = string.Format(@"SELECT  *     
                                          FROM E_{0}_{1}
                                        ", info.ProjectName, info.MainFormName);
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				dt = SqlCalls.RunQueries(comm, Connection);
				Connection.Close();

				foreach (DataColumn column in dt.Columns)
				{
					if ( dtFields != null)
					{
						if (tableColumns.Contains(column.ColumnName))
							columns.Add(new Tuple<string, string, Type>(column.ColumnName, info.MainFormName, column.DataType));
					}
					else
					{
						if (!GlobalAttributes.Fields.Contains(column.ColumnName) && !column.ColumnName.EndsWith("_TEXT"))
							columns.Add(new Tuple<string, string, Type>(column.ColumnName, info.MainFormName, column.DataType));
					}

				}

			
			//ALFDebugHelper.Log(88, columns);
			return columns;
		}

		public static List<object> TakeFieldsForTable(Info info, List<string> dtFields, Func<eBADBProvider> CreateDatabaseProvider, GlobalAttributes GlobalAttributes)
		{
			List<string> tableColumns = new List<string>();
			if (dtFields != null)
			{
				foreach (string column in dtFields)
				{
					//if (dr["CHECKED"].ToString() == "1")
					{
						tableColumns.Add(column);
					}
				}
			}
			List<string> columns = new List<string>();
			List<Tuple<string,Type>> columnDetails = new List<Tuple<string, Type>>();
			DataTable dt ;
				var comm = string.Format(@"SELECT  *     
                                          FROM {0}
                                        ", info.MainFormName);
				eBADB.eBADBProvider db = CreateDatabaseProvider();
				SqlConnection Connection = (SqlConnection)db.Connection;
				dt = SqlCalls.RunQueries(comm, Connection);
				Connection.Close();

				foreach (DataColumn column in dt.Columns)
				{
					if (dtFields != null)
					{
					if (tableColumns.Contains(column.ColumnName))
					{
						columns.Add(column.ColumnName);
						columnDetails.Add(new Tuple<string, Type>(column.ColumnName, column.DataType));
					}
				}
					else
					{
					if (!GlobalAttributes.Fields.Contains(column.ColumnName) && !column.ColumnName.EndsWith("_TEXT") && !column.ColumnName.EndsWith("_Text"))
					{
						columns.Add(column.ColumnName);
						columnDetails.Add(new Tuple<string, Type>(column.ColumnName, column.DataType));
					}
				}

				}
			return new List<object>() { columns, columnDetails };
		}

		public static void AddDictionary(object list, string columnName, DataTable dt)
		{
			Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)list;
			dictFields.Add(columnName, dt);
		}
		public static Dictionary<string, DataRow> GetAllMtl(Info info, Func<eBADBProvider> CreateDatabaseProvider)
		{
			DataTable dtTemp = new DataTable();
			dtTemp.Columns.Add("0", typeof(string));
			dtTemp.Columns.Add("1", typeof(string));
			dtTemp.Columns.Add("2", typeof(string));
			dtTemp.Columns.Add("3", typeof(string));
			DataRow drTemp = dtTemp.NewRow();
			dtTemp.Rows.Add(drTemp);
			DataRow drTemp2 = dtTemp.NewRow();
			dtTemp.Rows.Add(drTemp2);
			Dictionary<string, DataRow> dictAllMtl = new Dictionary<string, DataRow>
			{
				{ "-1", dtTemp.Rows[0] },
				{ "0", dtTemp.Rows[1] }
			};
			DataTable dt = new DataTable();
			var comm = string.Format(@"SELECT *
											    FROM {0}	 WITH (NOLOCK)
												ORDER BY ID
                                        ",info.MtlTableName);
			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
			try
			{
				SqlCommand Command = new SqlCommand(comm, Connection);
				SqlDataAdapter Adapter = new SqlDataAdapter(Command);
				Adapter.Fill(dt);
			}
			catch
			{
				//ALFDebugHelper.Log(16, comm);
				throw new Exception(comm);
			}
			foreach (DataRow dr in dt.Rows)
			{
				dictAllMtl.Add(dr[0].ToString(), dr);
			}

			return dictAllMtl;
		}

		public static void AddDataTable(object list, string columnName, DataTable dt)
		{
			List<KeyValuePair<string[], DataTable>> fields = (List<KeyValuePair<string[], DataTable>>)list;
			fields.Add(new KeyValuePair<string[], DataTable>(new string[] { columnName, dt.Rows.Count.ToString() }, dt));
		}


		public static Object TakeMtlValues(Info info, List<Tuple<string, string, Type>> columns, string lang, Func<eBADBProvider> CreateDatabaseProvider, Action<Object, string, DataTable> addFunc, object fields, string type, bool multiLanguage)
		{
			//Object fields = new Object();
			
			foreach (Tuple<string, string, Type> column in columns)
			{
				string columnName = column.Item1;
				string FormName = column.Item2;
				if (columnName.StartsWith("mtl"))
				{
					DataTable dt = new DataTable();
					var comm = string.Format(@"SELECT {1}
												  FROM {3}
												 WITH (NOLOCK)
												WHERE TYPE=(SELECT TOP(1) TYPE
															  FROM {3}
															  WITH (NOLOCK)
															 WHERE ID= (Select TOP(1) {0} 
																		  FROM E_{4}_{2} FRM WITH(NOLOCK)
																		  WHERE {0} IS NOT NULL AND {0} > 0))
															   AND  ISACTIVE=1
																ORDER BY {1}
                                        ", columnName, multiLanguage? info.MtlColumnName+lang + type: info.MtlColumnName + type, FormName, info.MtlTableName,info.ProjectName);
					eBADB.eBADBProvider db = CreateDatabaseProvider();
					SqlConnection Connection = (SqlConnection)db.Connection;
					try
					{
						SqlCommand Command = new SqlCommand(comm, Connection);
						SqlDataAdapter Adapter = new SqlDataAdapter(Command);
						Adapter.Fill(dt);
					}
					catch 
					{
						//ALFDebugHelper.Log(16, comm);
						throw new Exception(comm);
					}
					addFunc(fields, columnName, dt);
				}
				if (columnName.StartsWith("cmb") && info.DictCmbFieldQueries.ContainsKey(columnName))
				{
					DataTable dt ;
					var comm = string.Format(info.DictCmbFieldQueries[columnName]);
					eBADB.eBADBProvider db = CreateDatabaseProvider();
					SqlConnection Connection = (SqlConnection)db.Connection;
					dt = SqlCalls.RunQueries(comm, Connection);
					addFunc(fields, columnName, dt);
				}
			}
			return fields;
		}
		public static Object TakeMtlValues(Info info, List<string> columns, string langCode, Func<eBADBProvider> CreateDatabaseProvider, Action<Object, string, DataTable> addFunc, object fields, string type, bool multiLanguage)
		{
			//Object fields = new Object();

			foreach (string column in columns)
			{
				string columnName = column;
				if (columnName.StartsWith("mtl"))
				{
					DataTable dt = new DataTable();
					var comm = string.Format(@"SELECT {1}
												  FROM {3}
												 WITH (NOLOCK)
												WHERE TYPE=(SELECT TOP(1) TYPE
															  FROM {3}
															  WITH (NOLOCK)
															 WHERE ID= (Select TOP(1) {0} 
																		  FROM {2} FRM WITH(NOLOCK)
																		  WHERE {0} IS NOT NULL AND {0} > 0))
															   AND  ISACTIVE=1
																ORDER BY {1}
                                        ", columnName, multiLanguage ? info.MtlColumnName + langCode + type : info.MtlColumnName + type, info.MainFormName, info.MtlTableName);
					eBADB.eBADBProvider db = CreateDatabaseProvider();
					SqlConnection Connection = (SqlConnection)db.Connection;
					try
					{
						SqlCommand Command = new SqlCommand(comm, Connection);
						SqlDataAdapter Adapter = new SqlDataAdapter(Command);
						Adapter.Fill(dt);
					}
					catch 
					{
						//ALFDebugHelper.Log(16, comm);
						throw new Exception(comm);
					}
					addFunc(fields, columnName, dt);
				}
				if (columnName.StartsWith("cmb") && info.DictCmbFieldQueries.ContainsKey(columnName))
				{
					DataTable dt ;
					var comm = string.Format(info.DictCmbFieldQueries[columnName]);
					eBADB.eBADBProvider db = CreateDatabaseProvider();
					SqlConnection Connection = (SqlConnection)db.Connection;
					dt = SqlCalls.RunQueries(comm, Connection);
					addFunc(fields, columnName, dt);
				}
			}
			return fields;
		}

		public static Dictionary<string,Dictionary<string,string>> TakeMtlDictionary(Info info, List<string> columns, string lang, SqlConnection Connection,  bool multiLanguage)
		{
			Dictionary<string, Dictionary<string, string>> fields = new Dictionary<string, Dictionary<string, string>>();

			foreach (string column in columns)
			{
				string columnName = column;
				if (columnName.StartsWith("mtl"))
				{
					DataTable dt = new DataTable();
					var comm = string.Format(@"SELECT {1} ,ID
												  FROM {3}
												 WITH (NOLOCK)
												WHERE TYPE=(SELECT TOP(1) TYPE
															  FROM {3}
															  WITH (NOLOCK)
															 WHERE ID= (Select TOP(1) {0} 
																		  FROM {2} FRM WITH(NOLOCK)
																		  WHERE {0} IS NOT NULL AND {0} > 0))
															   AND  ISACTIVE=1
																ORDER BY {1}
                                        ", columnName, multiLanguage ? info.MtlColumnName + lang  : info.MtlColumnName , info.MainFormName, info.MtlTableName);
					try
					{
						SqlCommand Command = new SqlCommand(comm, Connection);
						SqlDataAdapter Adapter = new SqlDataAdapter(Command);
						Adapter.Fill(dt);
					}
					catch (Exception ex)
					{
						//ALFDebugHelper.Log(16, comm);
						throw new Exception(comm+" "+ex.Message);
					}
					Dictionary<string, string> columnDict = new Dictionary<string, string>();
					foreach (DataRow dr in dt.Rows)
					{
						columnDict[dr[0].ToString()] = dr[1].ToString();
					}
					fields.Add(columnName, columnDict);
				}
				if (columnName.StartsWith("cmb") && info.DictCmbFieldQueries.ContainsKey(columnName))
				{
					DataTable dt ;
					var comm = string.Format(info.DictCmbFieldQueries[columnName]);
					dt = SqlCalls.RunQueries(comm, Connection);
					Dictionary<string, string> columnDict = new Dictionary<string, string>();
					foreach (DataRow dr in dt.Rows)
					{
						columnDict[dr[0].ToString()] = dr[1].ToString();
					}
					fields.Add(columnName, columnDict);
				}
			}
			return fields;
		}

        public static Dictionary<string, Dictionary<string, string>> TakeMtlDictionary(Info info, List<Tuple<string, string, Type>> columns, string lang, SqlConnection Connection, bool multiLanguage)
        {
            Dictionary<string, Dictionary<string, string>> fields = new Dictionary<string, Dictionary<string, string>>();

            foreach (Tuple<string, string, Type> column in columns)
            {
                string columnName = column.Item1;
                string FormName = column.Item2;
                if (columnName.StartsWith("mtl"))
                {
                    DataTable dt = new DataTable();
                    var comm = string.Format(@"SELECT {1} ,ID
												  FROM {3}
												 WITH (NOLOCK)
												WHERE TYPE=(SELECT TOP(1) TYPE
															  FROM {3}
															  WITH (NOLOCK)
															 WHERE ID= (Select TOP(1) {0} 
																		  FROM {2} FRM WITH(NOLOCK)
																		  WHERE {0} IS NOT NULL AND {0} > 0))
															   AND  ISACTIVE=1
																ORDER BY {1}
                                        ", columnName, multiLanguage ? info.MtlColumnName + lang : info.MtlColumnName, "E_"+info.ProjectName+"_"+FormName, info.MtlTableName);
                    try
                    {
                        SqlCommand Command = new SqlCommand(comm, Connection);
                        SqlDataAdapter Adapter = new SqlDataAdapter(Command);
                        Adapter.Fill(dt);
                    }
                    catch (Exception ex)
                    {
                        //ALFDebugHelper.Log(16, comm);
                        throw new Exception(comm + " " + ex.Message);
                    }
                    Dictionary<string, string> columnDict = new Dictionary<string, string>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        columnDict[dr[0].ToString()] = dr[1].ToString();
                    }
                    fields.Add(columnName, columnDict);
                }
                if (columnName.StartsWith("cmb") && info.DictCmbFieldQueries.ContainsKey(columnName))
                {
                    DataTable dt ;
                    var comm = string.Format(info.DictCmbFieldQueries[columnName]);
                    dt = SqlCalls.RunQueries(comm, Connection);
                    Dictionary<string, string> columnDict = new Dictionary<string, string>();
                    foreach (DataRow dr in dt.Rows)
                    {
                        columnDict[dr[0].ToString()] = dr[1].ToString();
                    }
                    fields.Add(columnName, columnDict);
                }
            }
            return fields;
        }

		public static List<int> FillPredefinedColumns(int i, int j, Info info, string uniqueIds, SqlConnection Connection, Worksheet ws, Excel excel)
		{
			string comm ;
			DataTable dt ;
			if (info.UniqueColumnInfo != null)
			{
				ws.Cells[0, i].PutValue(info.UniqueColumnInfo.Item1);
				ws.Cells[1, i].PutValue(info.UniqueColumnInfo.Item2);
				ws.Cells[2, i].PutValue(info.UniqueColumnInfo.Item3);

				ws.Cells[2, i].SetStyle(excel.Labelstyle);

				if (info.DownloadType != Info.EnumDownloadType.Import)
				{
					comm = string.Format(info.UniqueColumnInfo.Item4, uniqueIds);
					dt = SqlCalls.RunQueries(comm, Connection);
					for (int row = 0; row < dt.Rows.Count; row++)
					{
						ws.Cells[row + 3, i].PutValue(dt.Rows[row][0]);
						ws.Cells[row + 3, i].SetStyle(excel.CellStyle);
					}
				}
				i++;
			}
			if (info.PreDefinedColums != null)
			{
				foreach (var parameters in info.PreDefinedColums)
				{
					ws.Cells[0, i].PutValue(parameters.Item1);
					ws.Cells[1, i].PutValue(parameters.Item2);
					ws.Cells[2, i].PutValue(parameters.Item3);

					ws.Cells[2, i].SetStyle(excel.Labelstyle);


					if (info.DownloadType != Info.EnumDownloadType.Import)
					{
						comm = string.Format(parameters.Item4, uniqueIds);
						dt = SqlCalls.RunQueries(comm, Connection);
						for (int row = 0; row < dt.Rows.Count; row++)
						{
							ws.Cells[row + 3, i].PutValue(dt.Rows[row][0]);
							ws.Cells[row + 3, i].SetStyle(excel.CellStyle);
						}
					}
					i++;
				}
			}
			return new List<int>() { i, j };
		}
		public static List<string> CellValidations( GlobalAttributes globalAttributes, DataTable excelTable, DataTable excelTableTemp, InfoLog logRecord, Dictionary<string, DataTable> dictFields,bool hasUniqueId = true)
		{
			DataColumnCollection excelColumns = excelTable.Columns;
			List<string> lstDateColumns = new List<string>();
			List<string> lstMtlColumns = new List<string>();
			List<string> idsToDelete = new List<string>();
			List<string> lstDecimalColumns = new List<string>();
			foreach (DataColumn excelColumn in excelColumns)
			{
				if (excelColumn.ColumnName.Contains("Date") || globalAttributes.DateColumns.Contains(excelColumn.ColumnName))
				{
					lstDateColumns.Add(excelColumn.ColumnName);
				}
				else if (excelColumn.ColumnName.StartsWith("mtl") || excelColumn.ColumnName.StartsWith("cmb"))
				{
					lstMtlColumns.Add(excelColumn.ColumnName);
				}
				else if (excelColumn.DataType == typeof(Decimal) || excelColumn.DataType == typeof(double))
				{
					lstDecimalColumns.Add(excelColumn.ColumnName);
				}
			}
			for (int i = 0; i < excelTable.Rows.Count; i++)
			{
				string dateErrorLog = (hasUniqueId ? excelTable.Rows[i][0].ToString() + "ID'li" : (i + 4).ToString() + ".") + " satirin hatalı tarih alanları : ";
				string mtlErrorLog = (hasUniqueId ? excelTable.Rows[i][0].ToString() + "ID'li" : (i + 4).ToString() + ".") + " satirin hatalı combobox alanları : ";
				string decimalErrorLog = (hasUniqueId ? excelTable.Rows[i][0].ToString() + "ID'li" : (i + 4).ToString() + ".") + " satirin hatalı sayısal alanlar : ";
				string dateErrorFields = "";
				string mtlErrorFields = "";
				string decimalErrorFields = "";
				//EĞER VALİDASYONLARA UYMAYAN KOLONLAR VAR İSE ONLARI HATA LOGLARINA EKLEYİP O PERSONELİ ES GEÇİYORUZ.
				if (excelTableTemp !=null)
				{
					for (int j = 0; j < lstDateColumns.Count; j++)
					{
						if (excelTableTemp.Rows[i][lstDateColumns[j]] != DBNull.Value && excelTable.Rows[i][lstDateColumns[j]] == DBNull.Value)
						{
							dateErrorFields += lstDateColumns[j] + "(" + excelTable.Rows[i][lstMtlColumns[j]].ToString() + ")" + ", ";
						}
					}

					for (int j = 0; j < lstDecimalColumns.Count; j++)
					{
						if (excelTableTemp.Rows[i][lstDecimalColumns[j]] != DBNull.Value && excelTable.Rows[i][lstDecimalColumns[j]] == DBNull.Value)
						{
							decimalErrorFields += lstDecimalColumns[j] + "(" + excelTable.Rows[i][lstMtlColumns[j]].ToString() + ")" + ", ";
						}
					}
				}
				for (int j = 0; j < lstMtlColumns.Count; j++)
				{
					if (!dictFields[lstMtlColumns[j]].AsEnumerable().Any(row => excelTable.Rows[i][lstMtlColumns[j]].ToString() == row[0].ToString()))
					{
						if (!string.IsNullOrEmpty(excelTable.Rows[i][lstMtlColumns[j]].ToString()))
						{
							mtlErrorFields += lstMtlColumns[j]+"("+ excelTable.Rows[i][lstMtlColumns[j]].ToString() +")"+ ", ";
						}
					}
				}
				if (dateErrorFields != "" || mtlErrorFields != "" || decimalErrorFields != "")
				{
					logRecord.ErrorLog += Environment.NewLine + Environment.NewLine;
				}
				if (dateErrorFields != "")
				{
					logRecord.ErrorLog += dateErrorLog + dateErrorFields + Environment.NewLine;
					idsToDelete.Add(hasUniqueId ? excelTable.Rows[i][0].ToString() : (i).ToString());
				}
				if (mtlErrorFields != "")
				{
					logRecord.ErrorLog += mtlErrorLog + mtlErrorFields + Environment.NewLine;
					idsToDelete.Add(hasUniqueId ? excelTable.Rows[i][0].ToString() : (i).ToString());
				}
				if (decimalErrorFields != "")
				{
					logRecord.ErrorLog += decimalErrorLog + decimalErrorFields + Environment.NewLine;
					idsToDelete.Add(hasUniqueId ? excelTable.Rows[i][0].ToString() : (i).ToString());
				}
			}
			return idsToDelete;
		}



	}

}
