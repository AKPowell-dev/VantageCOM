using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing;

public abstract class BaseTextCheck : BaseCheck
{
	private Regex A;

	private string A;

	public Regex RegexObj
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public string Fix
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

	public abstract void Check(Range rng, string strText);

	public abstract void Check(object shp, TextRange2 rng, string strText);

	public override void Check()
	{
	}
}
