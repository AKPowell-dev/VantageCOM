using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class BulletStyleOption : ReformatOption
{
	[CompilerGenerated]
	private BulletStyle A;

	public BulletStyle Style
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

	public BulletStyleOption(BulletStyle bs, string strLabel)
		: base(strLabel)
	{
		Style = bs;
	}
}
