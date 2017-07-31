/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: pawel.pietralik
 * Data: 2017-07-28
 * Godzina: 09:28
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;

namespace ConsoleBeta {	
	/// <summary>
	/// Description of DomainAge.
	/// </summary>
	public class DomainAge {
		
		String filePath = @"C:\users\pawel.pietralik\desktop\domainsAge.xlsx", slice;
		int dn, wby, control;
		int data = DateTime.Now.Year;
		String [] splitedCells;
 		
		List<String> dns = new List<String>();
		List<String> wbys = new List<String>();
		
		List<int> wbyi = new List<int>();
		
		public static Dictionary<string, int> ageBox = new Dictionary<string, int>();
		
		public DomainAge() {
			
			FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
			IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
			            
			DataSet result = excelReader2007.AsDataSet();
					
			foreach (DataTable table in result.Tables) {
				for (int i = 0; i < table.Rows.Count; i++) {
					for (int j = 0; j < table.Columns.Count; j++){
						if(table.Rows[i].ItemArray[j].ToString() == "Domain") {
							dn = j;
						} else if(table.Rows[i].ItemArray[j].ToString() == "WBY") {
							wby = j;
						} 
					}
				}
			}
			excelReader2007.Close();
		}
		
		public void displayBasicData() {
			FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
			IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
			            
			DataSet result = excelReader2007.AsDataSet();
			
			foreach (DataTable table in result.Tables) {
		    	for (int i = 0; i < table.Rows.Count; i++) {
					Console.Write("\"" + i + " -> " + table.Rows[i].ItemArray[dn] + "\";");
					Console.WriteLine("\"" + table.Rows[i].ItemArray[wby] + "\";");
		       	}
		    }
			excelReader2007.Close();
		}
		
		public void displayDomainsAge() {
			FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
			IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
			            
			DataSet result = excelReader2007.AsDataSet();
			
			foreach (DataTable table in result.Tables) {
		    	for (int i = 0; i < table.Rows.Count; i++) {
					dns.Add(table.Rows[i].ItemArray[dn].ToString());
					wbys.Add(table.Rows[i].ItemArray[wby].ToString());
					//Console.Write("\"" + i + " -> " + table.Rows[i].ItemArray[dn] + "\";");
					//Console.WriteLine("\"" + table.Rows[i].ItemArray[wby] + "\";");
		       	}
		    }
			
			foreach(String param in wbys) {
				if(Int32.TryParse(param.Split('-')[0], out control)){
					wbyi.Add(Int32.Parse(param.Split('-')[0]));
		    	}
		    }
			
			foreach (DataTable table in result.Tables) {
		    	for (int i = 1; i < table.Rows.Count; i++) {						
						//Console.Write("\"" + i + " -> " + table.Rows[i-1].ItemArray[dn] + "\";");
						//Console.WriteLine("\"" + (data - wbyi[i - 1]) + "\";");
						ageBox.Add(table.Rows[i - 1].ItemArray[dn].ToString(), (data - wbyi[i - 1]));
				}
		    }
			excelReader2007.Close();
		}
	}
}
