using Genexcel.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Genexcel {
	public class Sheet {
		public Sheet() {}
		public Sheet(string name) { this.Name = name; }
		public string Name { get; set; } = "Plan";

		public List<Chart> Charts { get; } = new List<Chart>();

		//List<List<Cell>> _cells = new List<List<Cell>>();
		Dictionary<int, Dictionary<int, Cell>> _cells = new Dictionary<int, Dictionary<int, Cell>>();

		Cell InitCell(int row, int col) {
			if (!_cells.ContainsKey(row)) { _cells[row] = new Dictionary<int, Cell>(); }
			if (!_cells[row].ContainsKey(col)) { _cells[row][col] = new Cell(row, col); }
			return _cells[row][col];
		}

		public IEnumerable<Cell> GetCells() {
			return _cells.Values.SelectMany(v => v.Values);
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
