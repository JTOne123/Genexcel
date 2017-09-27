using System;
using System.Collections.Generic;
using System.Text;

namespace Genexcel.Models {
	public class Cell {
		public int Row { get; private set; }
		public int Col { get; private set; }
		public Sheet Sheet { get; internal set; }
		public object Value { get; internal set; }
		public string Hyperlink { get; set; }
		private Cell(int row, int col, object value) {
			this.Row = row;
			this.Col = col;
			this.Value = value;
		}
		//For now, only supporting string, decimal and int
		public Cell(int row, int col, string value) : this(row, col, (object)value) { }
		public Cell(int row, int col, int value) : this(row, col, (object)value) { }
		public Cell(int row, int col, decimal value) : this(row, col, (object)value) { }
		public void SetValue(string value) { this.Value = value; }
		public void SetValue(int value) { this.Value = value; }
		public void SetValue(decimal value) { this.Value = value; }
	}
}
