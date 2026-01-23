using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class IeEgComma : BaseTextCheck
{
	[CompilerGenerated]
	private new IeEgTrailingComma A;

	private IeEgTrailingComma Convention
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return A;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			A = value;
		}
	}

	public IeEgComma(IeEgTrailingComma conv)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_000f: Invalid comparison between Unknown and I4
		Convention = conv;
		if ((int)conv == 1)
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
					base.RegexObj = Text.RegexIeEgNoComma();
					return;
				}
			}
		}
		base.RegexObj = Text.RegexIeEgComma();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Invalid comparison between Unknown and I4
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		List<TextRange2> list = new List<TextRange2>();
		int groupnum = (((int)Convention == 1) ? 1 : 0);
		foreach (Match item in base.RegexObj.Matches(strText))
		{
			Group obj = item.Groups[groupnum];
			list.Add(para.get_Characters(checked(obj.Index + 1), obj.Length));
			obj = null;
		}
		if (list.Count > 0)
		{
			Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.IeEgComma(sld, shp, list, Convention));
		}
		list = null;
	}
}
