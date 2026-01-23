using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class SpellingCanceled : BaseTextCheck
{
	public SpellingCanceled()
	{
		base.RegexObj = new Regex(XC.A(25524), RegexOptions.IgnoreCase);
	}

	public override void Check(Range rng, string strText)
	{
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
	}
}
