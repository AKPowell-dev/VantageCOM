using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SpaceBeforeClosing : BaseTextCheck
{
	public SpaceBeforeClosing()
	{
		base.RegexObj = Text.RegexSpaceBeforeClose();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = base.RegexObj.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match match = (Match)enumerator.Current;
					if (para.get_Characters(match.Index + 1, 1).Font.Superscript == MsoTriState.msoFalse)
					{
						string strFix;
						if (match.Value.Contains(AH.A(14255)))
						{
							strFix = AH.A(14255);
						}
						else if (match.Value.Contains(AH.A(15138)))
						{
							while (true)
							{
								switch (6)
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
							strFix = AH.A(15138);
						}
						else
						{
							strFix = AH.A(17801);
						}
						Main.Analysis.Errors.Add(new ExtraSpaceBeforeClosing(sld, shp, para.get_Characters(match.Index + 1, match.Length), strFix));
					}
					match = null;
				}
				while (true)
				{
					switch (3)
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
		}
	}
}
