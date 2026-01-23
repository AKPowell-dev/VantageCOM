using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TemplateRulesLegalNotices : BaseError
{
	public TemplateRulesLegalNotices()
		: base(ErrorType.TemplateRulesLegalNotices, Main.Analysis.Options.TemplateRules, null, null, blnHasFix: false)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(41111);
		((BaseError)this).Subtitle = AH.A(41158);
	}
}
