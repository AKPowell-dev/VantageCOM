using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MissingParentheses : BaseTextCheck
{
	[CompilerGenerated]
	private new Regex A;

	[CompilerGenerated]
	private Regex B;

	private Regex RegexOpening
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

	private Regex RegexClosing
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

	public MissingParentheses()
	{
		base.RegexObj = new Regex(AH.A(17489));
		RegexOpening = new Regex(AH.A(17514));
		RegexClosing = new Regex(AH.A(17519));
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		string text = strText;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				text = Text.MaskText(text, match);
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
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		checked
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = RegexOpening.Matches(text).GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Match match2 = (Match)enumerator2.Current;
					Main.Analysis.Errors.Add(new MissingClosingParenthesis(sld, shp, para.get_Characters(match2.Index + 1, match2.Length)));
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00d3;
					}
					continue;
					end_IL_00d3:
					break;
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
			IEnumerator enumerator3 = RegexClosing.Matches(text).GetEnumerator();
			try
			{
				while (enumerator3.MoveNext())
				{
					Match match3 = (Match)enumerator3.Current;
					Main.Analysis.Errors.Add(new MissingOpeningParenthesis(sld, shp, para.get_Characters(match3.Index + 1, match3.Length)));
				}
				while (true)
				{
					switch (5)
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
				IDisposable disposable = enumerator3 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
	}
}
