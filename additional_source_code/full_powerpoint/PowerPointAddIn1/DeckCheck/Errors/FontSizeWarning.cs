using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FontSizeWarning : BaseError
{
	public FontSizeWarning(string strSubtitle)
		: base(ErrorType.FontSizeWarning, ((Settings)Main.Analysis.Options).FontFamilySizeCount, null, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(31018);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(31059) + ((Settings)Main.Analysis.Options).MaxFontSizes + AH.A(31104);
	}
}
