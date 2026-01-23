using System.Runtime.CompilerServices;

namespace A;

internal abstract class JF
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private string C;

	[CompilerGenerated]
	private string D;

	[CompilerGenerated]
	private string E;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

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

	public string Title
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

	internal int Arguments
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

	internal string PlaceholderText
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

	internal string TextBoxToolTip
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

	public string ToolTip
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

	internal bool ShowIgnoreEmptyCells
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

	internal bool IsMatchCaseEnabled
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

	internal JF(string A, string B, int C)
	{
		UniqueId = A;
		Title = B;
		Arguments = C;
		IsMatchCaseEnabled = false;
		ShowIgnoreEmptyCells = false;
		PlaceholderText = "";
		TextBoxToolTip = null;
		ToolTip = null;
	}
}
