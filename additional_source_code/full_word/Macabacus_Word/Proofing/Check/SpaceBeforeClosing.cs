using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Check;

public sealed class SpaceBeforeClosing : BaseTextCheck
{
	public SpaceBeforeClosing()
	{
		base.RegexObj = Text.RegexSpaceBeforeClose();
	}

	public override void Check(Range rng, string strText)
	{
		Range duplicate = rng.Duplicate;
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			foreach (Match item in matchCollection)
			{
				Microsoft.Office.Interop.Word.Document document = rng.Document;
				object Start = item.Index + 1;
				object End = 1;
				if (document.Range(ref Start, ref End).Font.Superscript == 0)
				{
					while (true)
					{
						switch (5)
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
					if (item.Value.Contains(XC.A(20696)))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						strFix = XC.A(20696);
					}
					else if (item.Value.Contains(XC.A(6382)))
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
						strFix = XC.A(6382);
					}
					else
					{
						strFix = XC.A(25511);
					}
					duplicate.SetRange(rng.Characters[item.Index + 1].Start, rng.Characters[item.Index + item.Length].End);
					Main.Analysis.Errors.Add(new ExtraSpaceBeforeClosing(duplicate, strFix));
				}
				Match match = null;
			}
			matchCollection = null;
			duplicate = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = matchCollection.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match match = (Match)enumerator.Current;
					if (rng.get_Characters(match.Index + 1, 1).Font.Superscript == MsoTriState.msoFalse)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						string strFix;
						if (match.Value.Contains(XC.A(20696)))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							strFix = XC.A(20696);
						}
						else if (match.Value.Contains(XC.A(6382)))
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
							strFix = XC.A(6382);
						}
						else
						{
							strFix = XC.A(25511);
						}
						Main.Analysis.Errors.Add(new ExtraSpaceBeforeClosing(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(match.Index + 1, match.Length), strFix));
					}
					match = null;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			matchCollection = null;
		}
	}
}
