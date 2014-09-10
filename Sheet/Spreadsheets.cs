using System;
using System.Linq;
using System.Collections.Generic;

namespace Sheet
{
	public class Spreadsheets : Dictionary<int, Spreadsheet>
	{
        public int ActiveSpreadsheet = -1;
	    public int SpreadsheetIndex = -1;

		public void Calc()
		{
			foreach (var spreadsheet in Values) {
				foreach (var row in spreadsheet.Values) {
					foreach (var cell in row) {
						cell.Value.UpdateValue();
					}
				}
			}
		}

		public DateTime CalcLast;

		public Expression UpdateCellValue(Cell cell)
        {
			if (cell == null) {
				return null;
			}

            if (cell.HasFormula && cell.State.Count < 1)
            {
                cell.State.Push("Parsing");
				cell.CalcCount++;
				cell.CalcLast = CalcLast;
                var formula = new Formula();
                var value = formula.Parse(cell.Formula);
                cell.State.Pop();
                return value;
            }
            return cell.Exp;
        }

		public Expression CellValue(int spreadsheet, int row, int col)
	    {
			Cell cell = CellAt (spreadsheet, row, col);

            var value = UpdateCellValue(cell);
            return value;
	    }

		public Expression CellValue(Location loc)
        {
			Cell cell = CellAt (loc);

            var value = UpdateCellValue(cell);
            return value;
        }

		public Expression CellValue(Location locStart, Location locEnd)
        {
			var range = new Expression();

            for (var row = locStart.Row; row <= locEnd.Row; row++)
            {
                for (var col = locStart.Col; col <= locEnd.Col; col++)
                {
                    range.Push(
						CellAt(locStart.Sheet, row, col).UpdateValue()
                    );
                }
            }

            return range;
        }

        public Spreadsheet AddSpreadsheet()
	    {
            SpreadsheetIndex++;
            var spreadsheet = new Spreadsheet(SpreadsheetIndex);
            spreadsheet.Parent = this;
            Add(SpreadsheetIndex, spreadsheet);
            return spreadsheet;
	    }

		public Cell CellAt(Location loc)
		{
			return CellAt (loc.Sheet, loc.Row, loc.Col);
		}
		
		public Cell CellAt(int sheet, int row, int col) {
			Spreadsheet spreadsheet = null;
			Row _row = null;
			Cell cell = null;


			if (this.ContainsKey (sheet)) {
				spreadsheet = this[sheet];
			} else {
				return null;
			}

			if (spreadsheet.ContainsKey(row)) {
				_row = spreadsheet[row];
			} else {
				return null;
			}

			if (_row.ContainsKey (col)) {
				cell = _row[col];
			} else {
				return null;
			}

			return cell;
		}
    }
}
