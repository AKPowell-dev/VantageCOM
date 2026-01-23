using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public class ReformatOption
{
	[CompilerGenerated]
	private string A;

	public string Label
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

	public ReformatOption(string strLabel)
	{
		Label = strLabel;
	}
}
