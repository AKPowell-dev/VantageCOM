using System.Runtime.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Queries;

public abstract class BaseQuery
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private bool E;

	[CompilerGenerated]
	private bool F;

	[CompilerGenerated]
	private string B;

	[CompilerGenerated]
	private string C;

	[CompilerGenerated]
	private bool G;

	internal string UniqueId
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal bool LookInComments
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	internal bool LookInCharts
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal bool LookInEmptyCells
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	internal bool LookInFormulas
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal bool LookInValues
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal bool LookInHyperlinks
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal string Input1
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	internal string Input2
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal bool MatchCase
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal BaseQuery(string A)
	{
		UniqueId = A;
	}
}
