using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALFAsposeHelper
	{/// <summary>
	/// Form ve kolon bilgilerini tutan class
	/// </summary>
	public class GlobalAttributes
	{
		/// <summary>
		/// Excele indirilmemsini istemedigimiz alanlarin isimlerini Hashset{string} olarak tutuyoruz
		/// </summary>
		public HashSet<string> Fields { get; set; } = new HashSet<string>(){"ID"
				,"txtLogonLang"
				,"txtPersonnelID"
				,"docGlobalId"
				,"txtErrorMessage"
				,"txtEmployeeId"

		}; 
		/// <summary>
		/// kiril alfabasi
		/// </summary>
		internal  string Alphabet { get;  } = @"АаБбВвГгДдЕеЁёЖжЗзИиЙйКкЛлМмНнОоПпРрСсТтУуФфХхЦцЧчШшЩщЪъЫыЬьЭэЮюЯя .,-_/\!'1234567890:=();№  "; //Kiril Sorguları için kullanılıyor.
		/// <summary>
		/// Log Modulu olan ve authorization sistemi olan projeler icin bu dictionary kullaniliyor, yetkinin lookupdaki IDsi ve ilgili yetkinin formu yazilmalidir
		/// Ornegin;
		/// Dictionary{string, string}()
		///{
		///	{"5548","MdlPersonnelContract" },
		///	{"5549","MdlPersonnelEmployeeInfo" },
		///	{"5550","MdlPersonnelPassport" },
		///};
		/// </summary>
		public  Dictionary<string, string> AuthForms { get; set; } 
		/// <summary>
		/// Log Modulunde Modal formlarin isimlerini ve ilgili Modal formlarin Ana formda tutulduklari textbox ile auth yetkisindeki IDlerinin tutuldugu Dictionary
		/// </summary>
		public  Dictionary<string, string[]> DictFormTextBoxes { get; set; } = new Dictionary<string, string[]>(){
						 {"MdlPersonnelPassport",new string[]{"txtPassportModalFormId", "5550" } },
						{"MdlPersonnelPatentApplication",new string[]{"txtPatentApplicationModalFormId", "5554" } },
						{"MdlPersonnelWorkCard",new string[]{"txtWorkCardModalFormId","5551" }},
						{"MdlPersonnelRegistration",new string[]{"txtRegistrationModalFormId","5552" }},
						{"MdlPersonnelEmployeeInfo",new string[]{"txtEmployeeInfoModalFormId","5549" }},
						{"MdlPersonnelContract",new string[]{"txtContractModalFormId" ,"5548"}},
						{"MdlPersonnelInvitation",new string[]{"txtInvitationModalFormId","5553" }},
						{"MdlPersonnelEmployeeInfo2",new string[]{"txtEmployeeInfo2ModalFormId","5555" }},
						{"MdlPersonnelAdditionalInfo",new string[]{"txtAdditionalInfoModalFormId","6223" }},
						{"FrmPersonnel",new string[]{"ID","5607" } },
						{"MdlPersonnelInformation",new string[]{"txtPersonnelInformationModalFormId","5955" }}}; // form isimleri ile modal formlara ait Idleri tutan textboxların isimlerini tutan sözlük
		/// <summary>
		/// Log modulleri icin modal formlarin isimlerini , main formda IDlerini tutan textboxlari ve keywordlerini tutan field
		/// <para>ornek olarak;</para>
		///<para>List&lt;Tuple&lt;string, string, string&gt;&gt;(){</para>	
		///<para>⠀⠀⠀⠀new Tuple&lt;string, string, string&gt;("MdlPersonnelPassport","txtPassportModalFormId","passport"),</para>	
		///<para>⠀⠀⠀⠀new Tuple&lt;string, string, string&gt;("MdlPersonnelPatentApplication","txtPatentApplicationModalFormId","patent"),</para>	
		///<para>⠀⠀⠀⠀new Tuple&lt;string, string, string&gt;("FrmPersonnel","ID","personnel" )};</para>	
		/// </summary>
		public List<Tuple<string,string,string>> Forms { get; set; } = new List<Tuple<string, string, string>>(){
						  new Tuple<string,string,string>("MdlPersonnelPassport","txtPassportModalFormId","passport"),
						 new Tuple<string,string,string>("MdlPersonnelPatentApplication","txtPatentApplicationModalFormId","patent"),
						 new Tuple<string,string,string>("FrmPersonnel","ID","personnel" )};

		
		/// <summary>
		/// Eger kolonlar arasinda tarih olup icinde Date gecmeyenler varsa bu listeye eklenmeleri gerekli
		/// </summary>
		public HashSet<string> DateColumns { get; set; } = new HashSet<string>(){ "txtUrineTest", "txtXray", "txtNarcology", "txtJobEntryNotice", "txtOfficialExitNotice" }; // İçinde "Date" içermeyen tarih alanlarını tutan array

		/// <summary>
		/// Veriyi excele cektigimizde gosterilmesini istedigimiz ama veriyi yuklerken sisteme aktarilmasini istemedigimiz(guvenlik nedenleriyle) kolonlar
		/// </summary>
		public  HashSet<string> RestrictedFields { get; set; } = new HashSet<string>(){ };// { "mtlProjectRegion", "mtlProject", "mtlLocationStatus" };
	}
}
