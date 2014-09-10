using System;
using Sheet;
namespace UnitTest
{
	public class SpreadsheetsTester : Spreadsheets
	{
		new public Cell CellAt(int sheet, int row, int col)
		{
			Spreadsheet spreadsheet = null;
			Row _row = null;
			Cell cell = null;

			if (this.ContainsKey (sheet)) {
				spreadsheet = this [sheet];
			} else {
				spreadsheet = this [sheet] = new Spreadsheet (sheet);
			}

			if (spreadsheet.ContainsKey (row)) {
				_row = spreadsheet [row];
			} else {
				_row = spreadsheet [row] = new Row ();
			}

			if (_row.ContainsKey (col)) {
				cell = _row [col];
			} else {
				cell = _row [col] = new Cell ();
			}

			return cell;
		}
	}
}

