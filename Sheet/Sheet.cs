using System;
using System.Collections.Generic;

namespace Sheet
{
    [Serializable]
	public class Column
	{
		public string formula { get; set; }
		public string value { get; set; }
		public string style { get; set; }
		public string @class { get; set; }
	}

    [Serializable]
	public class Row
	{
		public int height { get; set; }
		public List<Column> columns { get; set; }
	}

    [Serializable]
	public class FrozenAt
	{
		public int row { get; set; }
		public int col { get; set; }
	}

    [Serializable]
	public class Metadata
	{
		public List<string> widths { get; set; }
		public FrozenAt frozenAt { get; set; }
	}

    [Serializable]
	public class Sheet
	{
		public string title { get; set; }
		public List<Row> rows { get; set; }
		public Metadata metadata { get; set; }
	}
}

