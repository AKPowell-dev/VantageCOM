using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class IndentOption : ReformatOption
{
	[CompilerGenerated]
	private Indent A;

	public Indent Indent
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

	public IndentOption(Indent ind, string strLabel)
		: base(strLabel)
	{
		Indent = ind;
	}
}
