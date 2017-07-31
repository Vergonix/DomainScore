/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: pawel.pietralik
 * Data: 2017-07-24
 * Godzina: 10:41
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;

namespace ConsoleBeta
{
	class Program
	{
		public static void Main(string[] args)
		{
			// TODO: Implement Functionality Here
			var domainAge = new DomainAge();
			domainAge.displayDomainsAge();
			var reader = new Reader();
			reader.calculateDomainScores();
			
			
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}