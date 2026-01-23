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

public sealed class SpaceBeforeOpening : BaseTextCheck
{
	public SpaceBeforeOpening()
	{
		base.RegexObj = Text.RegexSpaceBeforeOpen();
	}

	public override void Check(Range rng, string strText)
	{
		Range duplicate = rng.Duplicate;
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		checked
		{
			foreach (Match item in matchCollection)
			{
				Group obj = item.Groups[1];
				duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + 2].End);
				if (duplicate.Font.Superscript == 0)
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
					if (obj.Value.Contains(XC.A(25505)))
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
						strFix = XC.A(20691);
					}
					else if (obj.Value.Contains(XC.A(6379)))
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
						strFix = XC.A(25514);
					}
					else
					{
						strFix = XC.A(25519);
					}
					Main.Analysis.Errors.Add(new MissingSpaceBeforeOpening(duplicate, strFix));
				}
				obj = null;
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
					Group obj = ((Match)enumerator.Current).Groups[1];
					if (rng.get_Characters(obj.Index + 1, 1).Font.Superscript == MsoTriState.msoFalse)
					{
						string strFix = (obj.Value.Contains(XC.A(25505)) ? XC.A(20691) : ((!obj.Value.Contains(XC.A(6379))) ? XC.A(25519) : XC.A(25514)));
						Main.Analysis.Errors.Add(new MissingSpaceBeforeOpening(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(obj.Index + 1, obj.Length), strFix));
					}
					obj = null;
				}
				while (true)
				{
					switch (2)
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
