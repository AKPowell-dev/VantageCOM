using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TemplateRulesCoverPosition : BaseError
{
	public TemplateRulesCoverPosition()
		: base(ErrorType.TemplateRulesCoverPosition, Main.Analysis.Options.TemplateRules, null, null, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(40922);
		((BaseError)this).Subtitle = AH.A(40971);
	}
}
