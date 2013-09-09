using System;
using System.Web.Script.Serialization;

namespace Sheet
{
	class MainClass
	{
		public static void Main (string[] args)
		{
			var js = new JavaScriptSerializer ();
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
							"'row':0," +
							"'col':0" +
						"}" +
					"}" +
				"}]";
			var jsonDeserialized = js.Deserialize<Sheet> (json);
			Console.WriteLine (jsonDeserialized.ToString());
		}
	}
}
