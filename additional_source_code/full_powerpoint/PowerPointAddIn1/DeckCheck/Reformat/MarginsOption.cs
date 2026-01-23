using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class MarginsOption : ReformatOption
{
	[CompilerGenerated]
	private Margins A;

	public Margins Margins
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

	public MarginsOption(Margins mgns, string strLabel)
		: base(strLabel)
	{
		Margins = mgns;
	}
}
