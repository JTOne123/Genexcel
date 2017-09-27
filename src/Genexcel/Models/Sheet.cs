using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Genexcel.Models {
	public class Sheet {
		public Sheet(Document document, string name =null) {
			this.Document = document;
			this.Name = name ?? "Plan";
		}
		//public Sheet(string name) { this.Name = name; }
		public string Name { get; set; }

		public List<Chart> Charts { get; } = new List<Chart>();
		
		Dictionary<int, Dictionary<int, Cell>> _cells = new Dictionary<int, Dictionary<int, Cell>>();
		//internal List<Column> Columns { get; } = new List<Column>();
		internal Column[] Columns { get; } = new Column[16384];
		internal bool HasCustomColumn { get; private set; }

		public Document Document { get; private set; }

		//Cell InitCell(int row, int col) {
		//	if (!_cells.ContainsKey(row)) { _cells[row] = new Dictionary<int, Cell>(); }
		//	if (!_cells[row].ContainsKey(col)) { _cells[row][col] = new Cell(row, col); }
		//	return _cells[row][col];
		//}

		public Sheet Add(Cell cell) {
			if (!_cells.ContainsKey(cell.Row)) { _cells[cell.Row] = new Dictionary<int, Cell>(); }
			_cells[cell.Row][cell.Col] = cell;
			cell.Sheet = this;
			return this;
		}

		//public Sheet Add(Column column) {
		//}

		internal IEnumerable<Cell> GetCells() {
			return _cells.Values.SelectMany(v => v.Values);
		}


		public void SetColumnWidth(double width, int from, int? to = null) {
			to = to ?? from;
			var colDef = new Column(/*from, to.Value*/) { Width = width };
			for(int i = from; i <= to; i++) {
				Columns[i] = colDef;
			}
			HasCustomColumn = true;
			//Columns.Add(new Column(from, to.Value) { Width = width });
		}

		//public Cell WriteToCell(int row, int col, string value) {
		//	var cell = InitCell(row, col);
		//	cell.Value = value;
		//	return cell;
		//}

		//public Cell WriteToCell(int row, int col, int value) {
		//	var cell = InitCell(row, col);
		//	cell.Value = value;
		//	return cell;
		//}

		//public Cell WriteToCell(int row, int col, decimal value) {
		//	var cell = InitCell(row, col);
		//	cell.Value = value;
		//	return cell;
		//}
	}
}
