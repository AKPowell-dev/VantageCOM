namespace ExcelAddIn1.FormatPainter;

public sealed class Options
{
	public sealed class myChart
	{
		private bool A;

		private bool B;

		private bool C;

		private bool D;

		public bool Size
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool Top
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public bool Left
		{
			get
			{
				return C;
			}
			set
			{
				C = value;
			}
		}

		public bool Format
		{
			get
			{
				return D;
			}
			set
			{
				D = value;
			}
		}

		public myChart()
		{
			A = false;
			B = false;
			C = false;
			D = false;
		}
	}

	public sealed class myPlotArea
	{
		private bool A;

		private bool B;

		private bool C;

		public bool Size
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool Location
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public bool Format
		{
			get
			{
				return C;
			}
			set
			{
				C = value;
			}
		}

		public myPlotArea()
		{
			A = false;
			B = false;
			C = false;
		}
	}

	public sealed class mySeries
	{
		private bool A;

		private bool B;

		private bool C;

		private bool D;

		private bool E;

		private bool F;

		private bool G;

		public bool Format
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool GapWidthOverlap
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public bool FirstSliceAngle
		{
			get
			{
				return C;
			}
			set
			{
				C = value;
			}
		}

		public bool Explosion
		{
			get
			{
				return D;
			}
			set
			{
				D = value;
			}
		}

		public bool DataLabels
		{
			get
			{
				return E;
			}
			set
			{
				E = value;
			}
		}

		public bool ErrorBars
		{
			get
			{
				return F;
			}
			set
			{
				F = value;
			}
		}

		public bool UpDownBars
		{
			get
			{
				return G;
			}
			set
			{
				G = value;
			}
		}

		public mySeries()
		{
			A = false;
			B = false;
			C = false;
			D = false;
			E = false;
			F = false;
			G = false;
		}
	}

	public sealed class myLegend
	{
		private bool A;

		private bool B;

		public bool Position
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool Format
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public myLegend()
		{
			A = false;
			B = false;
		}
	}

	public sealed class myTitle
	{
		private bool A;

		private bool B;

		public bool Format
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool Position
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public myTitle()
		{
			A = false;
			B = false;
		}
	}

	public sealed class myDataTable
	{
		private bool A;

		public bool Format
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public myDataTable()
		{
			A = false;
		}
	}

	public class myPrimaryValueAxis
	{
		private bool A;

		private bool B;

		private bool C;

		private bool D;

		public bool Scale
		{
			get
			{
				return A;
			}
			set
			{
				A = value;
			}
		}

		public bool Gridlines
		{
			get
			{
				return B;
			}
			set
			{
				B = value;
			}
		}

		public bool Ticks
		{
			get
			{
				return C;
			}
			set
			{
				C = value;
			}
		}

		public bool Title
		{
			get
			{
				return D;
			}
			set
			{
				D = value;
			}
		}

		public myPrimaryValueAxis()
		{
			A = false;
			B = false;
			C = false;
			D = false;
		}
	}

	public sealed class myPrimaryCategoryAxis : myPrimaryValueAxis
	{
	}

	public sealed class mySecondaryValueAxis : myPrimaryValueAxis
	{
	}

	public sealed class mySecondaryCategoryAxis : myPrimaryValueAxis
	{
	}

	private myChart A;

	private myPlotArea A;

	private mySeries A;

	private myLegend A;

	private myTitle A;

	private myDataTable A;

	private myPrimaryValueAxis A;

	private myPrimaryCategoryAxis A;

	private mySecondaryValueAxis A;

	private mySecondaryCategoryAxis A;

	public myChart Chart
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myPlotArea PlotArea
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public mySeries Series
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myLegend Legend
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myTitle Title
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myDataTable DataTable
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myPrimaryValueAxis PrimaryValueAxis
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public myPrimaryCategoryAxis PrimaryCategoryAxis
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public mySecondaryValueAxis SecondaryValueAxis
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public mySecondaryCategoryAxis SecondaryCategoryAxis
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public Options()
	{
		Chart = new myChart();
		PlotArea = new myPlotArea();
		Series = new mySeries();
		PrimaryValueAxis = new myPrimaryValueAxis();
		PrimaryCategoryAxis = new myPrimaryCategoryAxis();
		SecondaryValueAxis = new mySecondaryValueAxis();
		SecondaryCategoryAxis = new mySecondaryCategoryAxis();
		Legend = new myLegend();
		Title = new myTitle();
		DataTable = new myDataTable();
	}
}
