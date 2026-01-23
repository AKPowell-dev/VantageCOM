using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class FontStyleOption : ReformatOption
{
	[CompilerGenerated]
	private FontStyle A;

	public FontStyle Style
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

	public FontStyleOption(FontStyle sty, string strLabel)
		: base(strLabel)
	{
		Style = sty;
	}
}
