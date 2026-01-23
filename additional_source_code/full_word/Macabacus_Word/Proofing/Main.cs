namespace Macabacus_Word.Proofing;

public sealed class Main
{
	private static Analysis A;

	public static Analysis Analysis
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}
}
