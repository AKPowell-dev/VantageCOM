using System.Runtime.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class ParenthesesPair
{
	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private string A;

	public int StartIndex
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

	public int Length
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

	public string FunctionName
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

	public ParenthesesPair(int intStart, string strFunction)
	{
		StartIndex = intStart;
		Length = 0;
		FunctionName = strFunction;
	}
}
