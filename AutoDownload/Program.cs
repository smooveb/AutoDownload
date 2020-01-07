using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading.Tasks;

using System.Collections.ObjectModel;

namespace AutoDownload
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args == null || args.Length < 1)
			{
				ShowHelp();
				return;
			}
			string inputURL = args[0];
			if (string.IsNullOrEmpty(inputURL) || !inputURL.StartsWith("http"))
			{
				ShowHelp();
				return;
			}

			int beginPos = inputURL.IndexOf('<');
			int endPos = inputURL.IndexOf('>');

			if (beginPos > -1 && endPos > -1 && beginPos < endPos)
			{
				DateTime targetDate = DateTime.Today;

				if (args.Length > 1)
				{
					string targetDayAbbr = args[1];
					DayOfWeek? targetDWK = null;
					if (!string.IsNullOrEmpty(targetDayAbbr))
					{
						switch (targetDayAbbr.ToUpper())
						{
							case "SU":
								targetDWK = DayOfWeek.Sunday;
								break;
							case "MO":
								targetDWK = DayOfWeek.Monday;
								break;
							case "TU":
								targetDWK = DayOfWeek.Tuesday;
								break;
							case "WE":
								targetDWK = DayOfWeek.Wednesday;
								break;
							case "TH":
								targetDWK = DayOfWeek.Thursday;
								break;
							case "FR":
								targetDWK = DayOfWeek.Friday;
								break;
							case "SA":
								targetDWK = DayOfWeek.Saturday;
								break;
						}
						if (targetDWK == null)
						{
							Console.WriteLine("Error: bad day of the week abbreviation.");
							return;
						}

						for (; targetDate.DayOfWeek != targetDWK; )
						{
							targetDate = targetDate.AddDays(-1);
						}
						Console.WriteLine("Dwk: " + targetDWK.Value.ToString() + " date:" + targetDate.ToShortDateString());
					}
				}
				string dateFormat = inputURL.Substring(beginPos + 1, endPos - beginPos-1);
				Console.WriteLine("Date format: " + dateFormat );
				Console.WriteLine("Formatted Date: " + targetDate.ToString(dateFormat));

				inputURL = inputURL.Substring(0, beginPos) + targetDate.ToString(dateFormat) + inputURL.Substring(endPos+1);
			}

			int filePos = inputURL.LastIndexOf("/");
			if (filePos < -1)
			{
				Console.WriteLine("Error: bad URL, missing filename.");
				return;
			}

			string filename = inputURL.Substring(filePos + 1);
			//string filename = "wort_" + targetDate.ToString("yyMMdd") + "_200001wonder.mp3";


			if (string.IsNullOrEmpty(filename))
			{
				Console.WriteLine("Error: bad URL, missing filename.");
				return;
			}


			string downloadFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, filename);
			FileInfo fileInfo = new FileInfo(downloadFilePath);

			Console.WriteLine("URL: " + inputURL);
			Console.WriteLine("File: " + downloadFilePath);

			if (fileInfo.Exists)
			{
				Console.WriteLine("File exists.");
				return;   //forget it, file is already downloaded.
			}

			string scriptText = "import-module BitsTransfer" + Environment.NewLine +
				"Start-BitsTransfer -Source " + inputURL + " -Destination " + downloadFilePath;

			//Console.ReadLine();
			//return;

			Console.WriteLine(RunScript(scriptText));

			//Console.ReadLine();
		}

		private static void ShowHelp()
		{
			Console.WriteLine("To download a file enter the url address of the file as the first command line parameter.\n");
			Console.WriteLine("To include a date value enter the date format as part of the URL enclosed in <> brackets.");
			Console.WriteLine("To download on a specific day of the week specify the day as the 2nd parameter: Su,Mo,Tu,We,Th,Fr,Sa\n");
			Console.WriteLine("Example Usage:");
			Console.WriteLine("AutoDownload.exe http://archive.wortfm.org/mp3/wort_<yyMMdd>_200001wonder.mp3 Mo");
			Console.ReadLine();
		}

		//The code below is from:
		//http://www.codeproject.com/Articles/18229/How-to-run-PowerShell-scripts-from-C
		private static string RunScript(string scriptText)
		{
			// create Powershell runspace

			Runspace runspace = RunspaceFactory.CreateRunspace();

			// open it

			runspace.Open();

			// create a pipeline and feed it the script text

			Pipeline pipeline = runspace.CreatePipeline();
			pipeline.Commands.AddScript(scriptText);

			// add an extra command to transform the script
			// output objects into nicely formatted strings

			// remove this line to get the actual objects
			// that the script returns. For example, the script

			// "Get-Process" returns a collection
			// of System.Diagnostics.Process instances.

			pipeline.Commands.Add("Out-String");

			// execute the script

			Collection<PSObject> results = pipeline.Invoke();

			// convert the script result into a single string

			StringBuilder stringBuilder = new StringBuilder();
			foreach (PSObject obj in results)
			{
				stringBuilder.AppendLine(obj.ToString());
			}

			// close the runspace
			runspace.Close();

			return stringBuilder.ToString();
		}
	}
}
