using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class GrammarAdverseAverse : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexAdverse
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

	private Regex RegexAverse
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

	public GrammarAdverseAverse()
	{
		RegexAdverse = new Regex(AH.A(15436), RegexOptions.IgnoreCase);
		RegexAverse = new Regex(AH.A(15483), RegexOptions.IgnoreCase);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = RegexAdverse.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[1];
					string input = Regex.Replace(obj.Value, AH.A(15402), AH.A(8100));
					input = Regex.Replace(input, AH.A(15409), AH.A(7902));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAdverseAverse(sld, shp, para.get_Characters(obj.Index + 1, obj.Length), input));
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
			IEnumerator enumerator2 = RegexAverse.Matches(strText).GetEnumerator();
			try
			{
				while (enumerator2.MoveNext())
				{
					Group obj2 = ((Match)enumerator2.Current).Groups[1];
					string input = Regex.Replace(obj2.Value, AH.A(15416), AH.A(15421));
					input = Regex.Replace(input, AH.A(15426), AH.A(15431));
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.GrammarAdverseAverse(sld, shp, para.get_Characters(obj2.Index + 1, obj2.Length), input));
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
				IDisposable disposable = enumerator2 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
	}
}
