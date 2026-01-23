using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarIts : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexIts
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

	private Regex RegexItsApostrophe
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

	public GrammarIts()
	{
		RegexIts = new Regex(AH.A(17279), RegexOptions.IgnoreCase);
		RegexItsApostrophe = new Regex(AH.A(17298), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = RegexIts.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string input = Regex.Replace(obj.Value, AH.A(17216), AH.A(17225));
					input = Regex.Replace(input, AH.A(17234), AH.A(17225));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarIts(sld, shp, para.get_Characters(obj.Index + 1, obj.Length), input));
					obj = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = RegexItsApostrophe.Matches(strText).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Group obj2 = ((Match)enumerator2.Current).Groups[1];
					string input = Regex.Replace(obj2.Value, AH.A(17243), AH.A(17254));
					input = Regex.Replace(input, AH.A(17261), AH.A(17272));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarIts(sld, shp, para.get_Characters(obj2.Index + 1, obj2.Length), input));
					obj2 = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}
}
