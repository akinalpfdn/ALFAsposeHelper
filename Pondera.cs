using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using eBADB;

namespace ALFAsposeHelper
{
	 class Pondera
	{
		 public List<string> PersonnelValidations(Info info,
			   GlobalAttributes globalAttributes, DataTable excelTable, Func<eBADBProvider> CreateDatabaseProvider, List<Tuple<string, string, Type>> columns, string lang, InfoLog logRecord, string logonUser,bool multiLanguage)
		{
			DataColumnCollection excelColumns = excelTable.Columns;
			List<string> lstDateColumns = new List<string>();
			List<string> lstMtlColumns = new List<string>();
			List<string> idsToDelete = new List<string>();
			List<string> lstDecimalColumns = new List<string>();

			Dictionary<string, DataTable> dictFields = (Dictionary<string, DataTable>)Fields.TakeMtlValues(info, columns, lang, CreateDatabaseProvider, Fields.AddDictionary, new Dictionary<string, DataTable>(), " ,PARENTID, ID", multiLanguage);
			
			foreach (DataColumn excelColumn in excelColumns)
			{
				 if (excelColumn.ColumnName.StartsWith("mtl"))
				{
					lstMtlColumns.Add(excelColumn.ColumnName);
				}
			
			}
			
			DataTable dtRedList = GetRedList(logRecord, logonUser, CreateDatabaseProvider);
			Dictionary<string, PersonnelStatus> dicPersonnelStatus = GetPersonnelStatus(CreateDatabaseProvider);
			Tuple<Dictionary<string, string>, Dictionary<string, string>> dictStatuses = GetStatuses(CreateDatabaseProvider, lang);
			for (int i = 0; i < excelTable.Rows.Count; i++)
			{
				string dateErrorLog = excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personelin hatalı tarih alanları : ";
				string mtlErrorLog = excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personelin hatalı combobox alanları : ";
				string decimalErrorLog = excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personelin hatalı sayısal alanlar : ";
				//string dateErrorFields = "";
				string mtlErrorFields = "";
				//string decimalErrorFields = "";
				string[] exitMtl = new string[] { "Cikis", "Exit", "Уволен" };
				if (excelTable.Columns.Contains("mtlGeneralStatus") && !exitMtl.Contains(excelTable.Rows[i]["mtlGeneralStatus"].ToString()))
				{

					if (dtRedList.AsEnumerable().Any(row2 => excelTable.Rows[i]["txtEmployeeId"].ToString() == row2.Field<String>("txtEmployeeId")))
					{
						logRecord.ErrorLog += excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personel Kirmizi Listededir! " + Environment.NewLine;
						idsToDelete.Add(excelTable.Rows[i]["txtEmployeeId"].ToString());
					}
				}
				for (int j = 0; j < lstMtlColumns.Count; j++)
				{

					if (lstMtlColumns[j] == "mtlEmployeeGender" || lstMtlColumns[j] == "mtlEmployeeNationality") continue;
					if (lstMtlColumns[j] == "mtlProjectRegion")
					{
						if (!excelTable.Columns.Contains("mtlLocationStatus"))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Lokasyon Bilgisi Eksik), ";
						}
						if (!excelTable.Columns.Contains("mtlProject"))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Proje Bilgisi Eksik), ";
						}
					}
					if (lstMtlColumns[j] == "mtlProject")
					{
						if (!excelTable.Columns.Contains("mtlLocationStatus"))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Lokasyon Bilgisi Eksik), ";
						}
						if (!excelTable.Columns.Contains("mtlProjectRegion") || (!string.IsNullOrEmpty(excelTable.Rows[i]["mtlProject"].ToString()) && string.IsNullOrEmpty(excelTable.Rows[i]["mtlProjectRegion"].ToString())))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Proje Bölgesi Eksik), ";
						}
						else if ((string.IsNullOrEmpty(excelTable.Rows[i]["mtlProject"].ToString()) && !string.IsNullOrEmpty(excelTable.Rows[i]["mtlProjectRegion"].ToString())))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Proje Bilgisi Eksik), ";
						}
						else
						{
							try
							{
								DataRow drProject = dictFields["mtlProject"].Select(string.Format("VALUE1{1} ='{0}' ", excelTable.Rows[i]["mtlProject"].ToString(), lang))[0];
								DataRow drProjectRegion = dictFields["mtlProjectRegion"].Select(string.Format("VALUE1{1} ='{0}' ", excelTable.Rows[i]["mtlProjectRegion"].ToString(), lang))[0];
								if (drProject["PARENTID"].ToString() != drProjectRegion["ID"].ToString())
								{
									mtlErrorFields += lstMtlColumns[j] + "(Proje-Bölge Eşleşmiyor), ";
								}
							}
							catch (IndexOutOfRangeException )
							{
								mtlErrorFields += lstMtlColumns[j] + "(Hatalı/Eksik Bilgi), ";
							}
						}

					}
					else if (lstMtlColumns[j] == "mtlLocationStatus")
					{
						if (!excelTable.Columns.Contains("mtlProjectRegion"))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Proje Bölgesi Eksik), ";
						}
						if (!excelTable.Columns.Contains("mtlProject") || (!string.IsNullOrEmpty(excelTable.Rows[i]["mtlLocationStatus"].ToString()) && string.IsNullOrEmpty(excelTable.Rows[i]["mtlProject"].ToString())))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Proje Bilgisi Eksik), ";
						}
						if ((string.IsNullOrEmpty(excelTable.Rows[i]["mtlLocationStatus"].ToString()) && !string.IsNullOrEmpty(excelTable.Rows[i]["mtlProject"].ToString())))
						{
							mtlErrorFields += lstMtlColumns[j] + "(Lokasyon Bilgisi Eksik), ";
						}
						else
						{
							try
							{
								DataRow drLocationStatus = dictFields["mtlLocationStatus"].Select(string.Format("VALUE1{1} ='{0}' ", excelTable.Rows[i]["mtlLocationStatus"].ToString(), lang))[0];
								DataRow drProject = dictFields["mtlProject"].Select(string.Format("VALUE1{1} ='{0}' ", excelTable.Rows[i]["mtlProject"].ToString(), lang))[0];
								if (drLocationStatus["PARENTID"].ToString() != "0" && drLocationStatus["PARENTID"].ToString() != drProject["ID"].ToString())
								{
									mtlErrorFields += lstMtlColumns[j] + "(Lokasyon-Proje Eşleşmiyor), ";
									//ALFDebugHelper.Log(33, excelTable.Rows[i]["mtlLocationStatus"].ToString(), drLocationStatus, excelTable.Rows[i]["mtlProject"].ToString(), drProject);
								}
							}
							catch (IndexOutOfRangeException )
							{
								mtlErrorFields += lstMtlColumns[j] + "(Hatalı/Eksik Bilgi), ";
							}
						}
					}
				}
				if (mtlErrorFields == "" && (excelTable.Columns.Contains("mtlGeneralStatus") || excelTable.Columns.Contains("mtlCurrentStatus")))
				{
					string currentGeneralStatusId = dicPersonnelStatus[excelTable.Rows[i]["txtEmployeeId"].ToString()].GeneralStatusId;
					string currentCurrentStatusId = dicPersonnelStatus[excelTable.Rows[i]["txtEmployeeId"].ToString()].CurrentStatusId;
					string generalStatusId = "";
					string currentStatusId = "";
					if (excelTable.Columns.Contains("mtlGeneralStatus") && !excelTable.Columns.Contains("mtlCurrentStatus"))
					{
						string generalStatus = excelTable.Rows[i]["mtlGeneralStatus"].ToString();
						generalStatusId = dictStatuses.Item1[generalStatus];
						currentStatusId = currentCurrentStatusId;
					}
					else if (!excelTable.Columns.Contains("mtlGeneralStatus") && excelTable.Columns.Contains("mtlCurrentStatus"))
					{
						string currentStatus = excelTable.Rows[i]["mtlCurrentStatus"].ToString();
						currentStatusId = dictStatuses.Item2[currentStatus];
						generalStatusId = currentGeneralStatusId;

					}
					else if (excelTable.Columns.Contains("mtlGeneralStatus") && excelTable.Columns.Contains("mtlCurrentStatus"))
					{
						string generalStatus = excelTable.Rows[i]["mtlGeneralStatus"].ToString();
						generalStatusId = dictStatuses.Item1[generalStatus];

						string currentStatus = excelTable.Rows[i]["mtlCurrentStatus"].ToString();
						currentStatusId = dictStatuses.Item2[currentStatus];
					}
					//ALFDebugHelper.Log(235, currentStatusId, generalStatusId, excelTable.Rows[i]["txtEmployeeId"].ToString());
					if (generalStatusId == "47" && (currentStatusId == "50" || currentStatusId == "49" || currentStatusId == "6813"))
					{
						logRecord.ErrorLog += excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personel Genel Durum mevcut iken Güncel Durum çıkış, bekleniyor ya da sevki iptal olamaz! " + Environment.NewLine;
						idsToDelete.Add(excelTable.Rows[i]["txtEmployeeId"].ToString());
					}
					else if (generalStatusId == "45" && (currentStatusId != "49" && !string.IsNullOrEmpty(currentStatusId)))
					{
						logRecord.ErrorLog += excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personel Genel Durum bekleniyor iken Güncel Durum  bekleniyor olmalıdır! " + Environment.NewLine;
						idsToDelete.Add(excelTable.Rows[i]["txtEmployeeId"].ToString());
					}
					else if (generalStatusId == "6812" && (currentStatusId != "6813" && !string.IsNullOrEmpty(currentStatusId)))
					{
						logRecord.ErrorLog += excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personel Genel Durum sevki iptal iken Güncel Durum  sevki iptal olmalıdır! " + Environment.NewLine;
						idsToDelete.Add(excelTable.Rows[i]["txtEmployeeId"].ToString());
					}
					else if (generalStatusId == "46" && (currentStatusId != "50" && !string.IsNullOrEmpty(currentStatusId)))
					{
						logRecord.ErrorLog += excelTable.Rows[i]["txtEmployeeId"].ToString() + "ID'li personel Genel Durum çıkış iken Güncel Durum  çıkış olmalıdır! " + Environment.NewLine;
						idsToDelete.Add(excelTable.Rows[i]["txtEmployeeId"].ToString());
					}

				}
				
			}
			return idsToDelete;
		}
		public DataTable GetRedList(InfoLog logRecord,
			string logonUser,
			Func<eBADBProvider> CreateDatabaseProvider)
		{
			var comm = string.Format(@"/****** Script for SelectTopNRows command from SSMS  ******/
										SELECT    FRM.[txtEmployeeId]
										,FRM.txtName
										,FRM.txtSurname
										,FRM.txtNameSurnameRussian

										--,mtlExitReason
										FROM E_Pon002Personnel_FrmPersonnel FRM
										CROSS  APPLY(select top 1 * from [E_Pon002Personnel_MdlPersonnelEmployeeInfo] EMP where EMP.txtEmployeeId = FRM.txtEmployeeId AND mtlExitReason IN (5755,
										5761,
										5772,
										5775,
										5776,
										5777,
										5778,
										5779,
										5785,
										5787,
										5788,
										6257,
										6258)  order by ID DESC) EMPOLD
										INNER JOIN E_Pon002Personnel_MdlPersonnelEmployeeInfo EMP WITH(NOLOCK) ON  EMP.ID = FRM.txtEmployeeInfoModalFormId
										WHERE EMP.mtlGeneralStatus = 46


                                        ");
			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
			return SqlCalls.RunQueries(comm, Connection, logRecord, logonUser);
		}


		public string PersonnelAdditionalUpdate(KeyValuePair<string[], List<string>> columns, DataRow drUpdate, string setQuery, string globalId,Dictionary<string,object> dictProjectRegion)
		{
			if (string.Compare(columns.Key[0], "MdlPersonnelContract") == 0)
			{
				setQuery += dictProjectRegion[globalId] == DBNull.Value ?
					" txtProjectRegion = NULL, " :
					" txtProjectRegion =  '" + dictProjectRegion[globalId].ToString() + "', ";
			}
			if (string.Compare(columns.Key[0], "MdlPersonnelEmployeeInfo") == 0)
			{
				DataColumnCollection cols = drUpdate.Table.Columns;
				if (cols.Contains("cmbResponsibleStaff"))
				{
					setQuery += @"cmbResponsibleStaff_TEXT=(SELECT NAME FROM TbPon015ForemanList WHERE ID = '" + drUpdate["cmbResponsibleStaff"].ToString() + @"') , ";
				}
				if (cols.Contains("cmbGroupChief"))
				{
					setQuery += @" cmbGroupChief_TEXT=(SELECT NAME FROM TbPon015ForemanList WHERE ID = '" + drUpdate["cmbGroupChief"].ToString() + @"') ,";
				}
			}
			return setQuery;
		}
		public bool PersonnelValidations(Info info, int i, DataColumnCollection excelColumns, InfoLog logRecord, DataTable dtExcel,GlobalAttributes GlobalAttributes)
		{
			bool validation = false;
			string uniqueColumnName = info.UniqueColumnInfo.Item1;
			if (excelColumns.Contains("txtPassportIssuingAuthority"))
			{
				foreach (char character in dtExcel.Rows[i]["txtPassportIssuingAuthority"].ToString())
				{
					if (!GlobalAttributes.Alphabet.Contains(character))
					{
						logRecord.ErrorLog += dtExcel.Rows[i][uniqueColumnName].ToString() + " ID'li Satırın Pasaportu Veren Makam'ı Kiril Validasyonundan Geçememiştir." + Environment.NewLine;
						validation = true;
						break;
					}
				}
			}
			if (excelColumns.Contains("txtRegistrationAddress1"))
			{
				foreach (char character in dtExcel.Rows[i]["txtRegistrationAddress1"].ToString())
				{
					if (!GlobalAttributes.Alphabet.Contains(character))
					{
						logRecord.ErrorLog += dtExcel.Rows[i][uniqueColumnName].ToString() + " ID'li Satırın Önceki Registrasyon Adresi Kiril Validasyonundan Geçememiştir." + Environment.NewLine;
						validation = true;
						break;
					}
				}
			}
			if (excelColumns.Contains("txtRegistrationAddress2"))
			{
				foreach (char character in dtExcel.Rows[i]["txtRegistrationAddress2"].ToString())
				{
					if (!GlobalAttributes.Alphabet.Contains(character))
					{
						logRecord.ErrorLog += dtExcel.Rows[i][uniqueColumnName].ToString() + " ID'li Satırın Registrasyon Adresi Kiril Validasyonundan Geçememiştir." + Environment.NewLine;
						validation = true;
						break;
					}
				}
			}
			return validation;
		}
		public class PersonnelStatus
		{
			public List<string> GeneralStatus { get; set; }
			public List<string> CurrentStatus { get; set; }
			public string GeneralStatusId { get; set; }
			public string CurrentStatusId { get; set; }

			public PersonnelStatus(List<string> generalStatus, List<string> currentStatus)
			{
				this.GeneralStatus = generalStatus;
				this.CurrentStatus = currentStatus;
			}
			public PersonnelStatus(string generalStatusId, string currentStatusId)
			{
				this.GeneralStatusId = generalStatusId;
				this.CurrentStatusId = currentStatusId;
			}
		}

		public static Tuple<Dictionary<string, string>, Dictionary<string, string>> GetStatuses(Func<eBADBProvider> CreateDatabaseProvider, string lang)
		{
			Dictionary<string, string> dictStatusGeneral = new Dictionary<string, string>();
			Dictionary<string, string> dictStatusCurrent = new Dictionary<string, string>();

			var comm = string.Format(@"/****** Script for SelectTopNRows command from SSMS  ******/
										SELECT    ID
												,{0}

										FROM  tbpon000lookupValues GNS WHERE ISACTIVE = 1 AND TYPE IN (13)
                                        ", "VALUE1" + lang);
			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
			DataTable dtStatusGeneral = SqlCalls.RunQueries(comm, Connection);

			comm = string.Format(@"/****** Script for SelectTopNRows command from SSMS  ******/
										SELECT    ID
												,{0}

										FROM  tbpon000lookupValues GNS WHERE ISACTIVE = 1 AND TYPE IN (14)
                                        ", "VALUE1" + lang);
			DataTable dtStatusCurrent = SqlCalls.RunQueries(comm, Connection);
			foreach (DataRow dr in dtStatusGeneral.Rows)
			{
				dictStatusGeneral[dr[1].ToString()] = dr[0].ToString();
			}
			foreach (DataRow dr in dtStatusCurrent.Rows)
			{
				dictStatusCurrent[dr[1].ToString()] = dr[0].ToString();
			}

			return new Tuple<Dictionary<string, string>, Dictionary<string, string>>(dictStatusGeneral, dictStatusCurrent);
		}

		public Dictionary<string, PersonnelStatus> GetPersonnelStatus(Func<eBADBProvider> CreateDatabaseProvider)
		{
			Dictionary<string, PersonnelStatus> dictStatus = new Dictionary<string, PersonnelStatus>();

			var comm = string.Format(@"/****** Script for SelectTopNRows command from SSMS  ******/
										SELECT    FRM.[txtEmployeeId]
												,EMP.mtlGeneralStatus
												,EMP.mtlCurrentStatus
												,GNS.VALUE1TR GNSTR
												,GNS.VALUE1EN GNSEN
												,GNS.VALUE1RU GNSRU
												,CRS.VALUE1TR CRSTR
												,CRS.VALUE1EN CRSEN
												,CRS.VALUE1RU CRSRU

										FROM E_Pon002Personnel_FrmPersonnel FRM
										INNER JOIN E_Pon002Personnel_MdlPersonnelEmployeeInfo EMP WITH(NOLOCK) ON  EMP.ID = FRM.txtEmployeeInfoModalFormId
										LEFT JOIN tbpon000lookupValues GNS ON GNS.ID = EMP.mtlGeneralStatus
										LEFT JOIN tbpon000lookupValues CRS ON CRS.ID = EMP.mtlCurrentStatus
                                        ");
			eBADB.eBADBProvider db = CreateDatabaseProvider();
			SqlConnection Connection = (SqlConnection)db.Connection;
			DataTable dtStatus = SqlCalls.RunQueries(comm, Connection);
			foreach (DataRow dr in dtStatus.Rows)
			{
				dictStatus[dr["txtEmployeeId"].ToString()] = new PersonnelStatus(dr["mtlGeneralStatus"].ToString(), dr["mtlCurrentStatus"].ToString());
			}

			return dictStatus;
		}

	}
}
