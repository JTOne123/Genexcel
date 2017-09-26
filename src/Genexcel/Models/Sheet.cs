using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Genexcel.Models {
	public class Sheet {
		public Sheet() {}
		public Sheet(string name) { this.Name = name; }
		public string Name { get; set; } = "Plan";

		public List<Chart> Charts { get; } = new List<Chart>();
		
		Dictionary<int, Dictionary<int, Cell>> _cells = new Dictionary<int, Dictionary<int, Cell>>();
		//internal List<Column> Columns { get; } = new List<Column>();
		internal Column[] Columns { get; } = new Column[16384];
		internal bool HasCustomColumn { get; private set; }


		Cell InitCell(int row, int col) {
			if (!_cells.ContainsKey(row)) { _cells[row] = new Dictionary<int, Cell>(); }
			if (!_cells[row].ContainsKey(col)) { _cells[row][col] = new Cell(row, col); }
			return _cells[row][col];
		}

		

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

		public void WriteToCell(int row, int col, string value) {
			InitCell(row, col).Value = value;
		}

		public void WriteToCell(int row, int col, int value) {
			InitCell(row, col).Value = value;
		}

		public void WriteToCell(int row, int col, decimal value) {
			InitCell(row, col).Value = value;
		}
	}
}
