using A;
using MacabacusMacros.Proofing;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FontSizeInfo : BaseError
{
	public FontSizeInfo(string strSubtitle)
		: base(ErrorType.FontSizeInfo, (Severity)1, null, null, blnHasFix: false)
	{
		((BaseError)this).Title = AH.A(30832);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(30853) + ((Settings)Main.Analysis.Options).MaxFontSizes + AH.A(30987);
	}
}
