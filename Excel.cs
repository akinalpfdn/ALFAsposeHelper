using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ALFAsposeHelper
	{/// <summary>
	/// Bir DataTablei excele aktarmak icin kullanilan class
	/// </summary>
	public class Excel
	{
		/// <summary>
		/// rgb renklerden kirmizi icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Red {get;set;}
		/// <summary>
		/// rgb renklerden Yesil icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Green { get; set; }
		/// <summary>
		/// rgb renklerden Mavi icin in deger 0-255 orasinda olmali
		/// </summary>
		public int Blue { get; set; }
		/// <summary>
		/// Exceldeki Basliklar isin style 
		/// </summary>
		public Style Labelstyle { get; set; }
		/// <summary>
		/// Exceldeki celler icin style
		/// </summary>
		public Style CellStyle { get; set; }
		/// <summary>
		/// Default Constructor
		/// </summary>
		public Excel()
		{
			Red = 100;
			Green = 180;
			Blue = 250;
			Labelstyle = new CellsFactory().CreateStyle();
			Labelstyle.ForegroundColor = System.Drawing.Color.FromArgb(Red, Green, Blue);
			Labelstyle.Pattern = BackgroundType.Solid;
			Labelstyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.TopBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.RightBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			Labelstyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, System.Drawing.Color.Black);
			CellStyle = new CellsFactory().CreateStyle();
		}
		/// <summary>
		/// Basliklarin renklerini degistirmek icin kullanilacak constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		public Excel(int red,int green,int blue)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
		}
		/// <summary>
		/// Basliklarin renkleri ve stilini degistirmek icin kullanilmasi gereken constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		/// <param name="labelstyle">Aspose cell Stil elementi</param>
		public Excel(int red, int green, int blue,Style labelstyle)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
			this.Labelstyle = labelstyle;
		}
		/// <summary>
		/// Basliklarin renkleri ve stilini degistirmek icin kullanilmasi gereken constructor
		/// </summary>
		/// <param name="red">rgb renklerden kirmizi icin in deger 0-255 orasinda olmali</param>
		/// <param name="green">rgb renklerden Yesil icin in deger 0-255 orasinda olmali</param>
		/// <param name="blue">rgb renklerden Mavi icin in deger 0-255 orasinda olmali</param>
		/// <param name="labelstyle">Aspose cell Stil elementi</param>
		/// <param name="cellStyle">Aspose cell Stil elementi</param>
		public Excel(int red, int green, int blue, Style labelstyle, Style cellStyle)
		{
			this.Red = red;
			this.Green = green;
			this.Blue = blue;
			this.Labelstyle = labelstyle;
			this.CellStyle = cellStyle;
		}
		/// <summary>
		/// Excel Indirme methodu
		/// </summary>
		/// <param name="parameters">Her bir tuple 1 sheeti simgeler, 1. eleman sheet adi iken 2. eleman verileri tutan datatable idir(DataTablein kolon isimleri excelde gosterilir)</param>
		/// <param name="outputName">Indirilecek dosyanin uzanti dahil edilmemis haliyle adi</param>
		/// <param name="WriteToResponse">Ebada kullanilan dosya indirme methodu</param>
		public void Download(List<Tuple<string, DataTable>> parameters, string outputName, Action<Stream, string> WriteToResponse)
		{
			Workbook wb = new Workbook();
			int sheetNumber = 0;
			foreach (Tuple<string, DataTable> parameter in parameters)
			{
				string sheetName = parameter.Item1;
				DataTable data = parameter.Item2;
				if (sheetNumber > 0)
				{
					wb.Worksheets.Add();
				}
				Worksheet ws = wb.Worksheets[sheetNumber];
				ws.Name = sheetName;
				for (int col = 0; col < data.Columns.Count; col++)
				{
					ws.Cells[0, col].PutValue(data.Columns[col].ColumnName);
					ws.Cells[0, col].SetStyle(Labelstyle);
				}
				for (int row = 0; row < data.Rows.Count; row++)
				{
					for (int col = 0; col < data.Columns.Count; col++)
					{
						ws.Cells[row + 1, col].PutValue(data.Rows[row][col]);
						if (data.Columns[col].DataType == typeof(DateTime) || data.Columns[col].DataType == typeof(DateTime))
						{
							CellStyle.Number = 14;
						}
						else
						{
							CellStyle.Number = 0;
						}
						ws.Cells[row + 1, col].SetStyle(CellStyle);

					}
				}
				ws.AutoFitColumns();
				sheetNumber++;
			}
			using (Stream respStream = new MemoryStream())
			{
				wb.Save(respStream, Aspose.Cells.SaveFormat.Xlsx);
				respStream.Seek(0, SeekOrigin.Begin);
				WriteToResponse(respStream, outputName + ".xlsx");
			}
		}
	}
}
