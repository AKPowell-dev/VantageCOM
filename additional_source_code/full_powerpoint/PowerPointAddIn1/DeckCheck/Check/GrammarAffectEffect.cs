using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarAffectEffect : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexAffect
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

	private Regex RegexEffect
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

	public GrammarAffectEffect()
	{
		RegexAffect = new Regex(AH.A(15536), RegexOptions.IgnoreCase);
		RegexEffect = new Regex(AH.A(15615), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = RegexAffect.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string input = Regex.Replace(obj.Value, AH.A(15416), AH.A(8112));
					input = Regex.Replace(input, AH.A(15426), AH.A(7914));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAffectEffect(sld, shp, para.get_Characters(obj.Index + 1, obj.Length), input));
					obj = null;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = RegexEffect.Matches(strText).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Group obj2 = ((Match)enumerator2.Current).Groups[1];
					string input = Regex.Replace(obj2.Value, AH.A(15526), AH.A(8100));
					input = Regex.Replace(input, AH.A(15531), AH.A(7902));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAffectEffect(sld, shp, para.get_Characters(obj2.Index + 1, obj2.Length), input));
					obj2 = null;
				}
				while (true)
				{
					switch (4)
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
						switch (1)
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
