using System;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using IronXL;
using static System.Net.Mime.MediaTypeNames;

public class Program
{
    public static void Main()
    {

        string url = "https://bakerhughesrigcount.gcs-web.com/intl-rig-count?c=79687&p=irol-rigcountsintl";


        using (WebClient client = new WebClient())
        {

            string html = client.DownloadString(url);



            string href = "";

            string element = "";


            Regex regex = new Regex("<a\\s+href=.+.+>\\s*Worldwide\\s+Rig\\s+Counts\\s+-\\s+Current\\s+.*\\s*Historical\\s*Data</a>", RegexOptions.IgnoreCase);
            Match match;
            for (match = regex.Match(html); match.Success; match = match.NextMatch())
            {



                foreach (Group group in match.Groups)
                {
                    element = group.Value;

                }
            }



            Regex regex2 = new Regex("href\\s*=\\s*(?:\"(?<1>[^\"]*)\"|(?<1>\\S+))", RegexOptions.IgnoreCase);
            Match match2;
            for (match2 = regex2.Match(element); match2.Success; match2 = match2.NextMatch())
            {
                foreach (Group group2 in match2.Groups)
                {
                    href = group2.ToString();
                }
            }

            Console.WriteLine(href);

            Regex regex3 = new Regex(".*.com", RegexOptions.IgnoreCase);
            Match match3;

            string pathToFile = "";

            for (match3 = regex3.Match(url); match3.Success; match3 = match3.NextMatch())
            {
                foreach (Group group in match3.Groups)
                {
                    pathToFile = group.Value + href;
                }
            }

            Console.WriteLine(pathToFile);


            client.DownloadFile(pathToFile, "WorldwideRigCounts.xlsx");





            int numOfNewRows = 29;
        Workbook wb = new Workbook("C:\\Users\\teodo\\Documents\\RigCountExtraction\\RigCountExtractionTest\\Worldwide Rig Count Dec 2022.xlsx");
        Worksheet ws = wb.Worksheets[0];

        int beginning = 0;
        int end = 0;

        int year = DateTime.Now.Year;



        Cells cells = ws.Cells;

        string beginningComp = (year).ToString();
        string endComp = (year - 3).ToString();

        FindOptions findOptions = new FindOptions();
        findOptions.CaseSensitive = false;
        findOptions.LookInType = LookInType.Values;

        Aspose.Cells.Cell foundCell = cells.Find(beginningComp, null, findOptions);
        if (foundCell != null)
        {
            beginning = foundCell.Row;
            Console.WriteLine(beginning);

        }
        else
        {

        }



        foundCell = cells.Find(endComp, null, findOptions);
        if (foundCell != null)
        {
            end = foundCell.Row;
            Console.WriteLine(end);

        }
        

        if (end < beginning)
        {
            foundCell = cells.Find((year - 1).ToString(), null, findOptions);

            if (foundCell != null)
            {
                beginning = foundCell.Row;
                Console.WriteLine(beginning);
            }

        }

        ws.Cells.DeleteRows(0, beginning);

        ws.Cells.DeleteRows(end - beginning, ws.Cells.MaxRow - end + beginning);

        wb.Save("RigCountExtraction.csv");
    }



    }
}