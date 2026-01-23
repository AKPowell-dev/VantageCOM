using A;
using MacabacusMacros.Proofing;

namespace Macabacus_Word.Proofing.Errors;

public sealed class FontFamilyWarning : BaseError
{
	public FontFamilyWarning(string strSubtitle)
		: base(ErrorType.FontFamilyWarning, ((Settings)Main.Analysis.Options).FontFamilySizeCount, null, blnHasFix: false)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(25895);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = XC.A(25942) + Main.Analysis.Options.MaxFontFamilies + XC.A(25987);
	}
}
