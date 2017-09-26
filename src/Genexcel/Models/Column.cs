using System;
using System.Collections.Generic;
using System.Text;

namespace Genexcel.Models {
	public class Column {
		public const double DEFAULT_WIDTH = 9.140625;
		//public Column(int min, int max) {
		//	Min = min;
		//	Max = max;
		//}
		//public int Min { get; private set; }
		//public int Max { get; private set; }
		public double Width { get; set; } = DEFAULT_WIDTH;
	}
}
