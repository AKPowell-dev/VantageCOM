using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class AbbreviationSpacing : BaseTextCheck
{
	[CompilerGenerated]
	private new int A;

	private int RequiredSpaces
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
					base.RegexObj = new Regex(AH.A(15041) + Constants.REGEX_ABBREV_SPACING + AH.A(15056));
					base.Fix = "";
					RequiredSpaces = 0;
					return;
				}
			}
		}
		base.RegexObj = new Regex(AH.A(15063) + Constants.REGEX_ABBREV_SPACING + AH.A(15056));
		base.Fix = AH.A(14625);
		RequiredSpaces = 1;
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		List<TextRange2> list = new List<TextRange2>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[1];
				list.Add(para.get_Characters(checked(obj.Index + 1), obj.Length));
				obj = null;
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
		if (list.Count > 0)
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
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.AbbreviationSpacing(sld, shp, list, RequiredSpaces));
		}
		list = null;
	}
}
