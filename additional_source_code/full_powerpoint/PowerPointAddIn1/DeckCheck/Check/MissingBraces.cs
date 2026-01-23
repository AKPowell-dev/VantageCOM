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

public sealed class MissingBraces : BaseTextCheck
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

	public MissingBraces()
	{
		base.RegexObj = new Regex(AH.A(17402));
		RegexOpening = new Regex(AH.A(17427));
		RegexClosing = new Regex(AH.A(17432));
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		string text = strText;
		IEnumerator enumerator = base.RegexObj.Matches(strText).GetEnumerator();
		try
		{
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
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
					Main.Analysis.Errors.Add(new MissingClosingCurlyBrace(sld, shp, para.get_Characters(match2.Index + 1, match2.Length)));
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00cd;
					}
					continue;
					end_IL_00cd:
					break;
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
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = RegexClosing.Matches(text).GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Match match3 = (Match)enumerator3.Current;
					Main.Analysis.Errors.Add(new MissingOpeningCurlyBrace(sld, shp, para.get_Characters(match3.Index + 1, match3.Length)));
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
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator3 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
	}
}
