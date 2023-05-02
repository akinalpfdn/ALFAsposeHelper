using System;
using System.Data;
using System.Data.SqlClient;
using ALFeBAHelper;

namespace ALFAsposeHelper
{
	internal class SqlCalls
	{
		/// <summary>
		/// Sql sorgularını çalıştırıp gelen veriyi datatablea aktaran method
		/// </summary>
		/// <param name="comm"></param>
		/// <param name="Connection"></param>
		/// <returns></returns>
		public static DataTable RunQueries(string comm, SqlConnection Connection)
		{
			DataTable dt = new DataTable();
			try
			{
				SqlCommand Command = new SqlCommand(comm, Connection);
				SqlDataAdapter Adapter = new SqlDataAdapter(Command);
				Adapter.Fill(dt);
			}
			catch (Exception ex)
			{
				//ALFDebugHelper.Log(16, comm);
				throw new Exception(comm +" "+ex.Message);
			}
			return dt;
		}
		public static void RunQueriesNoReturn(string comm, SqlConnection Connection)
		{
			try
			{
				SqlCommand Command = new SqlCommand(comm, Connection);
				Command.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				//ALFDebugHelper.Log(16, comm);
				throw new Exception(comm + " " + ex.Message);
			}
		}
		/// <summary>
		/// Hata alınma ihtimali olan querylerde hata mesajını da oluşturmak için kullanılan method
		/// </summary>
		/// <param name="comm"></param>
		/// <param name="Connection"></param>
		/// <param name="logRecord"></param>
		/// <param name="logonUser"></param>
		/// <returns></returns>
		public static DataTable RunQueries(string comm, SqlConnection Connection, InfoLog logRecord, string logonUser)
		{
			DataTable dt = new DataTable();
			try
			{
				//Connection.Open();
				SqlCommand Command = new SqlCommand(comm, Connection);
				SqlDataAdapter Adapter = new SqlDataAdapter(Command);
				Adapter.Fill(dt);
			}
			catch
			{
				//ALFDebugHelper.Log(16, comm);
				logRecord.ErrorLog = "İşlem sırasında bir hata meydana gelmiştir." + Environment.NewLine;
				if (logonUser == "afidan" || logonUser == "admin")
				{
					logRecord.ErrorLog += "Hata ile ilgili sorguya aşağıda ulaşabilirsiniz." + Environment.NewLine;
					logRecord.ErrorLog += comm + Environment.NewLine;
				}
				throw new Exception(comm);
			}
			return dt;
		}
	}
}
