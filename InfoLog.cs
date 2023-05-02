using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace ALFAsposeHelper
{
	/// <summary>
	/// Excel yuklendikten sonra olusan bilgilendirme mesajlarini tutan class
	/// </summary>
	public class InfoLog
	{
		/// <summary>
		/// 
		/// </summary>
		public string ChangedFieldsNote { get; set; } = ""; //Değişen alanları en sonda çıktı vermek için kullanılan global nesne
		/// <summary>
		/// 
		/// </summary>
		public string MissingForms { get; set; } //Personelin oluşturulmamış alanlarını tutan nesne
		/// <summary>
		/// 
		/// </summary>
		public string ErrorLog { get; set; } //Gerçekleşen hataların kayıtlarını tutan 
		/// <summary>
		/// 
		/// </summary>
		public string UsedForms { get; set; } //Gerçekleşen hataların kayıtlarını tutan nesne
		/// <summary>
		/// Upload işlemi sırasında değişen alanları, oluşan hataları, açık ya da <br/> eksik formları txt dosyaları haline 
		/// getirip zipleyerek indirten method.
		/// </summary>
		internal void DownloadLogZip(string globalId, Action<Byte[], string> WriteToResponse)
		{
			List<string> txtNames = new List<string>();
			List<byte[]> byteArrays = new List<byte[]>();
			if (!string.IsNullOrEmpty(ChangedFieldsNote))
			{
				byteArrays
					.Add(new UTF8Encoding(true)
					.GetBytes(ChangedFieldsNote));
				txtNames.Add("Değişen_Alanlar.txt");
			}
			if (!string.IsNullOrEmpty(MissingForms))
			{
				byteArrays
					.Add(new UTF8Encoding(true)
					.GetBytes(MissingForms));
				txtNames.Add("Eksik_Formlar.txt");
			}
			if (!string.IsNullOrEmpty(UsedForms))
			{
				UsedForms = "EĞER FORM İLGİLİ KULLANICIDA AÇIK DEĞİLSE, UYGUN BİR ŞEKİLDE KAPATILMADIĞI ANLAMINA GELİR, BU" +
					" DURUMDA LÜTFEN 3 SAAT SONRA TEKRAR DENEYİNİZ YA DA FORMDA MANUEL DÜZENLEME YAPINIZ!" + Environment.NewLine + Environment.NewLine + UsedForms;
				byteArrays
					.Add(new UTF8Encoding(true)
					.GetBytes(UsedForms));
				txtNames.Add("Açık_Formlar.txt");
			}
			if (!string.IsNullOrEmpty(ErrorLog))
			{
				ErrorLog = "AŞAĞIDA YER ALAN IDLERINs BAZI ALANLARINDA HATALAR OLDUĞU İÇİN UPDATE EDİLEMEMİŞLERDİR!" + Environment.NewLine + Environment.NewLine + ErrorLog;
				byteArrays
					.Add(new UTF8Encoding(true)
					.GetBytes(ErrorLog));
				txtNames.Add("Hata_Logu.txt");
			}
			using (var compressedFileStream = new MemoryStream())
			{
				//Create an archive and store the stream in memory.
				using (var zipArchive = new ZipArchive(compressedFileStream, ZipArchiveMode.Create, false))
				{
					int i = 0;
					foreach (var byteFile in byteArrays)
					{
						//Create a zip entry for each attachment
						var zipEntry = zipArchive.CreateEntry(txtNames[i]);
						//Get the stream of the attachment
						using (var originalFileStream = new MemoryStream(byteFile))
						using (var zipEntryStream = zipEntry.Open())
						{
							//Copy the attachment stream to the zip entry stream
							originalFileStream.CopyTo(zipEntryStream);
						}
						i++;
					}
				}

				//ALFeBAHelper.ALFDebugHelper.Log(17, errorLog);
				WriteToResponse(compressedFileStream.ToArray(), "Log" + globalId + ".zip");
			}
		}
	}
}
