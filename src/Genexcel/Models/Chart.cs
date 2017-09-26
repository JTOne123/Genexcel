using System;
using System.Collections.Generic;
using System.Text;

namespace Genexcel.Models {
	public abstract class Chart {
		public ChartData Data { get; set; } = new ChartData();
	}

	public class ChartData {
		public List<string> Labels { get; set; } = new List<string>();
		public List<ChartDataset> Datasets { get; set; } = new List<ChartDataset>();
	}

	public class ChartDataset {
		public string Title { get; set; }
		public List<decimal> Data { get; set; } = new List<decimal>();
	}

	public class AreaChart : Chart {
	}

	public class BarChart : Chart {

	}
}
