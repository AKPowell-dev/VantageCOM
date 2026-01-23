using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class AbbreviationSpacing : BaseTextCheck
{
	private string m_A;

	private int A
	{
		get
		{
			return Conversions.ToInteger(this.m_A);
		}
		set
		{
			this.m_A = Conversions.ToString(value);
		}
	}

	public AbbreviationSpacing(UnitsSpacing conv)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		if ((int)conv == 0)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.RegexObj = new Regex(XC.A(24548) + Constants.REGEX_ABBREV_SPACING + XC.A(24531));
					base.Fix = "";
					A = 0;
					return;
				}
			}
		}
		base.RegexObj = new Regex(XC.A(24563) + Constants.REGEX_ABBREV_SPACING + XC.A(24531));
		base.Fix = XC.A(18458);
		A = 1;
	}

	public override void Check(Range rng, string strText)
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
					Match obj = (Match)enumerator.Current;
					Range duplicate = rng.Duplicate;
					Group obj2 = obj.Groups[1];
					duplicate.SetRange(rng.Characters[obj2.Index + 1].Start, rng.Characters[obj2.Index + obj2.Length].End);
					obj2 = null;
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.AbbreviationSpacing(duplicate, base.RegexObj.Replace(rng.Text, XC.A(24538) + base.Fix + XC.A(24543)), A));
					duplicate = null;
				}
				while (true)
				{
					switch (3)
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
			rng = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		MatchCollection matchCollection = base.RegexObj.Matches(strText);
		TextRange2 textRange;
		foreach (Match item in matchCollection)
		{
			Group obj = item.Groups[1];
			textRange = rng.get_Characters(checked(obj.Index + 1), obj.Length);
			Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.AbbreviationSpacing(RuntimeHelpers.GetObjectValue(shp), textRange, base.RegexObj.Replace(textRange.Text, XC.A(24538) + base.Fix + XC.A(24543)), A));
			obj = null;
		}
		matchCollection = null;
		textRange = null;
	}
}
