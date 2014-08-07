using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Web.UI;
using Sheet;
using UI;
namespace UnitTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var js = new JavaScriptSerializer();
            var json =
                "[{" +
                    "'title':'simple table 1'," +
                    "'rows':[" +
                        "{" +
                            "'height':18," +
                            "'columns':[" +
                                "{'formula':'100','value':'100','style':'height: 50px; background-color: red; color: blue;','class':'styleBold styleCenter'}," +
                                "{'formula':'A1*5','value':'500','style':'display:none;','class':'none'}" +
                            "]" +
                        "}" +
                    "]," +
                    "'metadata':{" +
                        "'widths':['120','120']," +
                        "'frozenAt':{" +
                            "'row':1," +
                            "'col':1" +
                        "}" +
                    "}" +
                "}]";

            var i = 0;
            while (i++ < 1000)
            {
                var jsonDeserialized = js.Deserialize<List<Sheet.Sheet>>(json);
                Console.WriteLine(jsonDeserialized.ToString());

                var htmlControl = new HtmlTables(jsonDeserialized).ToString();
                Console.Write(htmlControl);
            }

            Console.Write("Done!");


            var spreadsheets = new Spreadsheets();
            var spreadsheet = spreadsheets.AddSpreadsheet();

            var row = spreadsheet.AddRow();

            var cellA1 = row.AddCell("250");
            var cellB1 = row.AddCell("250");
            var cellC1 = row.AddCell("800 - (SUM(A1:B1) + 100)", true);

            var cell = spreadsheet["C", 1];
            var value = cell.UpdateValue();
            Console.Write(value.ToDouble());
            Console.Read();
        }
    }
}
