using System;
using System.Runtime.CompilerServices;
using A;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class MissingStraightQuotes : BaseTextCheck
{
	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(24629)))
		{
			return;
		}
		try
		{
			if (Strings.Split(strText, XC.A(24629)).Length % 2 != 0)
			{
				return;
			}
			checked
			{
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
					int num = strText.LastIndexOf('"') + 1;
					Range duplicate = rng.Duplicate;
					duplicate.SetRange(rng.Characters[num].Start, rng.Characters[num + 1].End);
					Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingStraightQuotes(duplicate));
					duplicate = null;
					return;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(XC.A(24629)))
		{
			return;
		}
		try
		{
			if (Strings.Split(strText, XC.A(24629)).Length % 2 != 0)
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingStraightQuotes(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(strText.LastIndexOf('"') + 1), 1)));
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
