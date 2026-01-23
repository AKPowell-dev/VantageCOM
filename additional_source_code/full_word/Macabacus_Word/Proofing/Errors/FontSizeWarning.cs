using A;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing.Errors;

public sealed class FontSizeWarning : BaseError
{
	public FontSizeWarning(string strSubtitle)
		: base(ErrorType.FontSizeWarning, ((Settings)Main.Analysis.Options).FontFamilySizeCount, null, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(26104);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(25942) + ((Settings)Main.Analysis.Options).MaxFontSizes + XC.A(26145);
	}
}
