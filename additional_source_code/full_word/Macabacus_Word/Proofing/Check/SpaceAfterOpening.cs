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

public sealed class SpaceAfterOpening : BaseTextCheck
{
	public SpaceAfterOpening()
	{
		base.RegexObj = Text.RegexSpaceAfterOpen();
	}

	public override void Check(Range rng, string strText)
	{
		Range duplicate = rng.Duplicate;
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
					duplicate.SetRange(rng.Characters[match.Index + 1].Start, rng.Characters[match.Index + 2].End);
					if (duplicate.Font.Superscript == 0)
					{
						while (true)
						{
							switch (2)
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
						if (match.Value.Contains(XC.A(25505)))
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
							strFix = XC.A(25505);
						}
						else if (match.Value.Contains(XC.A(6379)))
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
							strFix = XC.A(6379);
						}
						else
						{
							strFix = XC.A(25508);
						}
						duplicate.SetRange(rng.Characters[match.Index + 1].Start, rng.Characters[match.Index + match.Length].End);
						Main.Analysis.Errors.Add(new ExtraSpaceAfterOpening(duplicate, strFix));
					}
					match = null;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0187;
					}
					continue;
					end_IL_0187:
					break;
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
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
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
						string strFix;
						if (match.Value.Contains(XC.A(25505)))
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
							strFix = XC.A(25505);
						}
						else if (match.Value.Contains(XC.A(6379)))
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
							strFix = XC.A(6379);
						}
						else
						{
							strFix = XC.A(25508);
						}
						Main.Analysis.Errors.Add(new ExtraSpaceAfterOpening(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(match.Index + 1, match.Length), strFix));
					}
					match = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0115;
					}
					continue;
					end_IL_0115:
					break;
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
