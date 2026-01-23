using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarEnsureInsure : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexInsure
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

	private Regex RegexEnsure
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

	public GrammarEnsureInsure()
	{
		RegexInsure = new Regex(AH.A(17061), RegexOptions.IgnoreCase);
		RegexEnsure = new Regex(AH.A(17136), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = RegexInsure.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string input = Regex.Replace(obj.Value, AH.A(17051), AH.A(8112));
					input = Regex.Replace(input, AH.A(17056), AH.A(7914));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarEnsureInsure(sld, shp, para.get_Characters(obj.Index + 1, obj.Length), input));
					obj = null;
				}
				while (true)
				{
					switch (1)
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
				enumerator2 = RegexEnsure.Matches(strText).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Group obj2 = ((Match)enumerator2.Current).Groups[1];
					string input = Regex.Replace(obj2.Value, AH.A(15526), AH.A(8124));
					input = Regex.Replace(input, AH.A(15531), AH.A(7926));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarEnsureInsure(sld, shp, para.get_Characters(obj2.Index + 1, obj2.Length), input));
					obj2 = null;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (4)
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
