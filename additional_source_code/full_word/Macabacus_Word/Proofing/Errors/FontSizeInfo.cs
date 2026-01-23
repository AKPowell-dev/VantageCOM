using A;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing.Errors;

public sealed class FontSizeInfo : BaseError
{
	public FontSizeInfo(string strSubtitle)
		: base(ErrorType.FontSizeInfo, (Severity)1, null, blnHasFix: false)
	{
		((BaseError)this).Title = XC.A(26052);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(25724) + ((Settings)Main.Analysis.Options).MaxFontSizes + XC.A(26073);
	}
}
