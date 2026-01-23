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

public sealed class SpaceAfterClosing : BaseTextCheck
{
	public SpaceAfterClosing()
	{
		base.RegexObj = Text.RegexSpaceAfterClose();
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
				string strFix;
				if (obj.Value.Contains(XC.A(20696)))
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
					strFix = XC.A(25490);
				}
				else
				{
					strFix = ((!obj.Value.Contains(XC.A(6382))) ? XC.A(25500) : XC.A(25495));
				}
				duplicate.SetRange(rng.Characters[obj.Index + 1].Start, rng.Characters[obj.Index + obj.Length].End);
				Main.Analysis.Errors.Add(new MissingSpaceAfterClosing(duplicate, strFix));
				obj = null;
			}
			matchCollection = null;
			duplicate = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = matchCollection.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[1];
				string strFix;
				if (obj.Value.Contains(XC.A(20696)))
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
					strFix = XC.A(25490);
				}
				else if (obj.Value.Contains(XC.A(6382)))
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
					strFix = XC.A(25495);
				}
				else
				{
					strFix = XC.A(25500);
				}
				Main.Analysis.Errors.Add(new MissingSpaceAfterClosing(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), strFix));
				obj = null;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_00fa;
				}
				continue;
				end_IL_00fa:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
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
