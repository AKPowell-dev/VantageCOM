using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarAn : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexA
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

	private Regex RegexAn
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public GrammarAn()
	{
		string text = AH.A(15694);
		RegexA = new Regex(text + AH.A(15719), RegexOptions.IgnoreCase);
		RegexAn = new Regex(text + AH.A(16267), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		List<TextRange2> list = A(rng, strText, RegexA, 2);
		if (list.Count > 0)
		{
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAn(sld, shp, list, rng));
		}
		list = A(rng, strText, RegexAn, 2);
		if (list.Count > 0)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAn(sld, shp, list, rng));
		}
		list = null;
	}
}
