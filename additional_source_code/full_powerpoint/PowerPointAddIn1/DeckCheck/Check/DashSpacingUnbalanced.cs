using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class DashSpacingUnbalanced : BaseTextCheck
{
	public DashSpacingUnbalanced(DashSpacing conv)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Invalid comparison between Unknown and I4
		base.RegexObj = new Regex(Constants.REGEX_DASH_SPACING);
		string fix;
		if ((int)conv != 1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			fix = AH.A(15125);
		}
		else
		{
			fix = AH.A(15092);
		}
		base.Fix = fix;
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(AH.A(15092)))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<TextRange2> list = new List<TextRange2>();
			try
			{
				enumerator = base.RegexObj.Matches(strText).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[0];
					if (obj.Length == 2)
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
						list.Add(rng.get_Characters(checked(obj.Index + 1), obj.Length));
					}
					obj = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_00a6;
					}
					continue;
					end_IL_00a6:
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
			if (list.Count > 0)
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.DashSpacingUnbalanced(sld, shp, list, base.Fix));
			}
			list = null;
			return;
		}
	}
}
