using System;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class SlashSpacingUnbalanced : BaseTextCheck
{
	public SlashSpacingUnbalanced(SlashSpacing conv)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Invalid comparison between Unknown and I4
		base.RegexObj = new Regex(Constants.REGEX_SLASH_SPACING);
		if ((int)conv == 1)
		{
			base.Fix = XC.A(25450);
		}
		else
		{
			base.Fix = XC.A(25483);
		}
	}

	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(XC.A(25450)))
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				MatchCollection matchCollection;
				try
				{
					matchCollection = base.RegexObj.Matches(strText);
					foreach (Match item in matchCollection)
					{
						if (item.Groups[0].Length == 2)
						{
							Range duplicate = rng.Duplicate;
							duplicate.SetRange(rng.Characters[item.Index + 1].Start, rng.Characters[item.Index + item.Length].End);
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SlashSpacingUnbalanced(duplicate, base.Fix));
							duplicate = null;
						}
						_ = null;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				matchCollection = null;
				return;
			}
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(XC.A(25450)))
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
			MatchCollection matchCollection;
			try
			{
				matchCollection = base.RegexObj.Matches(strText);
				foreach (Match item in matchCollection)
				{
					Group obj = item.Groups[0];
					if (obj.Length == 2)
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
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.SlashSpacingUnbalanced(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(checked(obj.Index + 1), obj.Length), base.Fix));
					}
					obj = null;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			matchCollection = null;
			return;
		}
	}
}
