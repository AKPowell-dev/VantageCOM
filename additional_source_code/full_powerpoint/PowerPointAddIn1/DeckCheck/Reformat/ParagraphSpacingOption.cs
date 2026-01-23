using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class ParagraphSpacingOption : ReformatOption
{
	[CompilerGenerated]
	private LineSpacing A;

	public LineSpacing Spacing
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

	public ParagraphSpacingOption(LineSpacing sp, string strLabel)
		: base(strLabel)
	{
		Spacing = sp;
	}
}
