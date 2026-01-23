using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class TemplateRules
{
	public static void Check(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		if (!Rules.ProofingCheckRequiredSlides(pres))
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
			Main.Analysis.Errors.Add(new TemplateRulesRequiredSlides());
		}
		else if (!Rules.ProofingCheckCoverPosition(pres))
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
			Main.Analysis.Errors.Add(new TemplateRulesCoverPosition());
		}
		if (Rules.ProofingCheckLegalNotices(pres))
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			Main.Analysis.Errors.Add(new TemplateRulesLegalNotices());
			return;
		}
	}
}
