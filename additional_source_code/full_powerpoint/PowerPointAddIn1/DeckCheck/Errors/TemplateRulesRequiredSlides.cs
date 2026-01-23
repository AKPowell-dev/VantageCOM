using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TemplateRulesRequiredSlides : BaseError
{
	public TemplateRulesRequiredSlides()
		: base(ErrorType.TemplateRulesRequiredSlides, Main.Analysis.Options.TemplateRules, null, null, blnHasFix: false)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(41312);
		((BaseError)this).Subtitle = AH.A(41363);
		((BaseError)this).Tooltip = AH.A(41466);
	}
}
