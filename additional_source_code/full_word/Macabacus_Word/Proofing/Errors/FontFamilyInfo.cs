using A;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing.Errors;

public sealed class FontFamilyInfo : BaseError
{
	public FontFamilyInfo(string strSubtitle)
		: base(ErrorType.FontFamilyInfo, (Severity)1, null, blnHasFix: false)
	{
		((BaseError)this).Title = XC.A(25697);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(25724) + Main.Analysis.Options.MaxFontFamilies + XC.A(25858);
	}
}
