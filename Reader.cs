/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: pawel.pietralik
 * Data: 2017-07-24
 * Godzina: 10:43
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

namespace ConsoleBeta
{
	/// <summary>
	///Basic class for reading excel tables and processing them.
	/// </summary>
	public class Reader {
		
		String filePath, filePath2, savePath, line, directoryGzPath;
		Boolean goodPath;
		int eb, rd, tf, cf, item, control, licznik, dn, wby;
		Double score;
		
		List<int> ebli = new List<int>();
		List<int> rdli = new List<int>();
		List<int> tfli = new List<int>();
		List<int> cfli = new List<int>();
		
		List<String> ebls = new List<String>();
		List<String> rdls = new List<String>();
		List<String> tfls = new List<String>();
		List<String> cfls = new List<String>();
		List<String> outputFile = new List<String>();
		
		Dictionary<string, Double> sortBox = new Dictionary<string, Double>();
		Dictionary<string, int> compareBox = new Dictionary<string, int>();
		
		public Reader() {
			
			while(!goodPath) {
				try {
					unzipGzFile();
					
					Console.Write("Podaj ścieżkę dostepu do pliku z rozszerzeniem .xlsx: ");
					filePath = Console.ReadLine();
					
					if(filePath == "default")
						filePath = @"c:\users\pawel.pietralik\desktop\test\domeny.xlsx";
					
					FileStream stream = new FileStream(@filePath, FileMode.Open, FileAccess.Read);
			        IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
			            
			        DataSet result = excelReader2007.AsDataSet();
					
					foreach (DataTable table in result.Tables) {
						for (int i = 0; i < table.Rows.Count; i++) {
							for (int j = 0; j < table.Columns.Count; j++){
								if(table.Rows[i].ItemArray[j].ToString() == "Item") {
									item = j;
								} else if(table.Rows[i].ItemArray[j].ToString() == "External Backlinks") {
									eb = j;
								} else if (table.Rows[i].ItemArray[j].ToString() == "Referring Domains") {
									rd = j;
								} else if (table.Rows[i].ItemArray[j].ToString() == "Trust Flow") {
									tf = j;
								} else if (table.Rows[i].ItemArray[j].ToString() == "Citation Flow") {
									cf = j;
								} 
							}
						}
					}
			        goodPath = true;
			    	excelReader2007.Close();
			    	
				} catch(FileNotFoundException e) {
					Console.WriteLine("Nie znaleziono podanego pliku, podaj nową ścieżkę dostępu");
				} catch(DirectoryNotFoundException e) {
					Console.WriteLine("Nieprawidłowa ścieżka dostępu, podaj właściwą ścieżkę dostepu");
				} catch(IOException e) {
					Console.WriteLine("Nie można otworzyć pliku, ponieważ jest on aktualnie używany przez inny program. " +
					                  "Zamknij wszystkie programy z otwartym w nich plikiem.");
				}
			}
		}
		
		public void displayBasicFile() {
			goodPath = false;
			while(!goodPath) {
				try {
					FileStream stream = new FileStream(@filePath, FileMode.Open, FileAccess.Read);
			        IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
			            
			        DataSet result = excelReader2007.AsDataSet();
			            
			        Console.WriteLine();
			            
			        foreach (DataTable table in result.Tables) {
			            for (int i = 0; i < table.Rows.Count; i++) {
			            	for (int j = 0; j < table.Columns.Count; j++){
			            		Console.Write("\"" + table.Rows[i].ItemArray[j] + "\";");
			            	}
			            Console.WriteLine();	
			            }
			        }
			    goodPath = true;
			    excelReader2007.Close();
			    
				} catch(FileNotFoundException e) {
					Console.WriteLine("Nie znaleziono podanego pliku, podaj nową ścieżkę dostępu");
				} catch(DirectoryNotFoundException e) {
					Console.WriteLine("Nieprawidłowa ścieżka dostępu, podaj właściwą ścieżkę dostepu");
					filePath = Console.ReadLine();
				} catch(IOException e) {
					Console.WriteLine("Nie można otworzyć pliku, ponieważ jest on aktualnie używany przez inny program. " +
					                  "Zamknij wszystkie programy z otwartym w nich plikiem.");
				}
			}
		}
		
		public void displayCalculationParameters() {
			goodPath = false;
			while(!goodPath) {
				try {
					FileStream stream = new FileStream(@filePath, FileMode.Open, FileAccess.Read);
		            IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
		            
		            DataSet result = excelReader2007.AsDataSet();
		            
		            Console.WriteLine();
					
		            foreach (DataTable table in result.Tables) {
		            	for (int i = 0; i < table.Rows.Count; i++) {
		            		Console.Write("\"" + table.Rows[i].ItemArray[item] + "\";");
		            		Console.Write("\"" + table.Rows[i].ItemArray[eb] + "\";");
		            		Console.Write("\"" + table.Rows[i].ItemArray[rd] + "\";");
		            		Console.Write("\"" + table.Rows[i].ItemArray[tf] + "\";");
		            		Console.WriteLine("\"" + table.Rows[i].ItemArray[cf] + "\";");
		            		Console.WriteLine("==================================================================");
		            	}
		            }
		            goodPath = true;
		            excelReader2007.Close();
				} catch(FileNotFoundException e) {
					Console.WriteLine("Niepoprawna ścieżka dostępu, podaj właściwą ścieżkę");
				} catch(DirectoryNotFoundException e) {
					Console.WriteLine("Nieprawidłowa ścieżka dostępu, podaj właściwą ścieżkę dostepu");
				} catch(IOException e) {
					Console.WriteLine("Nie można otworzyć pliku, ponieważ jest on aktualnie używany przez inny program. " +
					                  "Zamknij wszystkie programy z otwartym w nich plikiem.");
				}
			}	
		}
		
		public void calculateDomainScores() {
			goodPath = false;
			while(!goodPath) {
				try {
					FileStream stream = new FileStream(@filePath, FileMode.Open, FileAccess.Read);
		            IExcelDataReader excelReader2007 = ExcelReaderFactory.CreateOpenXmlReader(stream);
		            
		            DataSet result = excelReader2007.AsDataSet();
		            
		            Console.WriteLine();
		            
		            foreach (DataTable table in result.Tables) {
		            	for (int i = 0; i < table.Rows.Count; i++) {
		            		ebls.Add(table.Rows[i].ItemArray[eb].ToString());
		            		rdls.Add(table.Rows[i].ItemArray[rd].ToString());
		            		tfls.Add(table.Rows[i].ItemArray[tf].ToString());
		            		cfls.Add(table.Rows[i].ItemArray[cf].ToString());
		            	}
		            }
		            
		            foreach(String param in ebls) {
		            	if(Int32.TryParse(param, out control)){
		            		ebli.Add(Int32.Parse(param));
		            	}
		            }
		            
		            foreach(String param in rdls) {
		            	if(Int32.TryParse(param, out control)){
		            		rdli.Add(Int32.Parse(param));
		            	}
		            }
		            
		            foreach(String param in tfls) {
		            	if(Int32.TryParse(param, out control)){
		            		tfli.Add(Int32.Parse(param));
		            	}
		            }
		            
		            foreach(String param in cfls) {
		            	if(Int32.TryParse(param, out control)){
		            		cfli.Add(Int32.Parse(param));
		            	}
		            }
		            
		            foreach (DataTable table in result.Tables) {
		            	for (int i = 0; i < table.Rows.Count; i++) {
		            		
		            		if(i == 0) {
		            			Console.Write("\"" + "DOMENA" + "\";");
		            			Console.WriteLine("\"" + "WYNIK" + "\";");
		            		} else {
		            			score = ((ebli[i - 1]*0.0005) + (rdli[i - 1]*7) + (tfli[i - 1]*8) + (cfli[i - 1]*2));
		            			//Console.Write("\"" + "||" + i + "||" + table.Rows[i].ItemArray[item] + "\";");
		            			//Console.WriteLine("\"" +  + "\";");
		            			
		            			sortBox.Add(table.Rows[i].ItemArray[item].ToString(), score);
		            		}
		            	}
		            }
		            var dict = DomainAge.ageBox;
		            //var dict2 = dict.Where(entry => sortBox[entry.Key] != entry.Value).ToDictionary(entry => entry.Key, entry => entry.Value);
		            var keysDictionary1HasThat2DoesNot = sortBox.Keys.Except(dict.Keys);
		            int licz = 0;
		            
		            foreach(String param in keysDictionary1HasThat2DoesNot) {
						int value = dict[param];
		            	Console.WriteLine(++licz + param + " -> " + value);
		            }
		            
		            var items = (from pair in sortBox orderby pair.Value descending select pair).Take(100);
		            
		            
		            foreach (KeyValuePair<string, Double> pair in items) {
		            	line = pair.Key + ";" + pair.Value;
		            	line = line.Replace(",", ".");
    					//Console.WriteLine(pair.Key + "  ->  "  + pair.Value);
    					outputFile.Add(line);
					}
		            
		            Console.WriteLine("\n========================================================");
		            Console.Write("Podaj ścieżkę gdzie chcesz zapisać obliczone dane: ");
		            savePath = Console.ReadLine();
		            Console.WriteLine("========================================================");
		            
		            if(savePath == "default") 
		            	savePath = @"C:\Users\pawel.pietralik\desktop\test\DomainsScores.txt";
		            else if(savePath == "desktop")
						savePath = @"c:\users\pawel.pietralik\desktop\DomainsScores.txt";
		                              
		            
		            File.WriteAllLines(savePath, outputFile);
		            
		            Console.WriteLine("Plik został poprawnie zapisany w:\n" + savePath + "\n");
		            Console.WriteLine("Otworzyć zapisany plik ? -> y/n");
		            String sign =  Console.ReadLine();
		            if(sign == "y")
		            	//Process.Start("notepad.exe", savePath); 
		            	Process.Start(savePath);
		            
		            goodPath = true;
		            excelReader2007.Close();
				} catch(FileNotFoundException e) {
					Console.WriteLine("Niepoprawna ścieżka dostępu, podaj właściwą ścieżkę");
				} catch(DirectoryNotFoundException e) {
					Console.WriteLine("Nieprawidłowa ścieżka dostępu, podaj właściwą ścieżkę dostepu");
				} catch(IOException e) {
					Console.WriteLine("Nie można otworzyć pliku, ponieważ jest on aktualnie używany przez inny program. " +
					                  "Zamknij wszystkie programy z otwartym w nich plikiem.");
				}
			}
		}
		
		public void unzipGzFile() {
			directoryGzPath = @"C:\Users\pawel.pietralik\Desktop\test";

            DirectoryInfo directorySelected = new DirectoryInfo(directoryGzPath);

            /*foreach (FileInfo fileToCompress in directorySelected.GetFiles())
            {
                Compress(fileToCompress);
            }*/

            foreach (FileInfo fileToDecompress in directorySelected.GetFiles("*.gz")) {
                Decompress(fileToDecompress);
            }
		}
		
		public static void Decompress(FileInfo fileToDecompress) {
            using (FileStream originalFileStream = fileToDecompress.OpenRead()) {
                string currentFileName = fileToDecompress.FullName;
                string newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);

                using (FileStream decompressedFileStream = File.Create(newFileName)) {
                    using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress)) {
                        decompressionStream.CopyTo(decompressedFileStream);
                        Console.WriteLine("========================================================");
                        Console.WriteLine("Decompressed: {0}", fileToDecompress.Name);
                        Console.WriteLine("Press any key to continue . . . ");
                        Console.WriteLine("========================================================");
						Console.ReadKey(true);
                    }
                }
            }
        }
	}
}
