using System;
using System.Collections.Generic;
using System.Text;

namespace Genexcel.Models {
	public class Cell {
		public int Row { get; private set; }
		public int Col { get; private set; }
		public object Value { get; internal set; }
		internal Cell(int row, int col) {
			this.Row = row;
			this.Col = col;
		}
	}
}
