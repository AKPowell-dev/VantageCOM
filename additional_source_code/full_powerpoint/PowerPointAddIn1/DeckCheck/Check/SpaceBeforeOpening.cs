using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SpaceBeforeOpening : BaseTextCheck
{
	public SpaceBeforeOpening()
	{
		base.RegexObj = Text.RegexSpaceBeforeOpen();
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
					Group obj = ((Match)enumerator.Current).Groups[1];
					if (para.get_Characters(obj.Index + 1, 1).Font.Superscript == MsoTriState.msoFalse)
					{
						while (true)
						{
							switch (3)
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
						string strFix;
						if (obj.Value.Contains(AH.A(17795)))
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
							strFix = AH.A(17804);
						}
						else if (obj.Value.Contains(AH.A(15135)))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							strFix = AH.A(17809);
						}
						else
						{
							strFix = AH.A(17814);
						}
						Main.Analysis.Errors.Add(new MissingSpaceBeforeOpening(sld, shp, para.get_Characters(obj.Index + 1, obj.Length), strFix));
					}
					obj = null;
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
						switch (2)
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
