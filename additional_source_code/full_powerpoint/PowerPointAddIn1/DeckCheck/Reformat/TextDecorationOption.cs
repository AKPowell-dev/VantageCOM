using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class TextDecorationOption : ReformatOption
{
	[CompilerGenerated]
	private TextDecoration A;

	public TextDecoration Decoration
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

	public TextDecorationOption(TextDecoration decor, string strLabel)
		: base(strLabel)
	{
		Decoration = decor;
	}
}
