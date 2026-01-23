using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.TextOps;

public sealed class ReplaceFontsOptions
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	internal bool A;

	internal bool B;

	internal bool C;

	internal string Find
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

	internal string Replace
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

	public ReplaceFontsOptions(string strFind, string strReplace, bool blnBold, bool blnItalic, bool blnUnderline)
	{
		Find = strFind;
		Replace = strReplace;
		A = blnBold;
		B = blnItalic;
		C = blnUnderline;
	}
}
