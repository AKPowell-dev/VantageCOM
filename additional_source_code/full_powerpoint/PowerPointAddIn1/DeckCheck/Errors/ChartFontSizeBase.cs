using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public abstract class ChartFontSizeBase : BaseError
{
	internal new int A;

	public ChartFontSizeBase(Slide sld, Shape shp, float size, int limit)
		: base(ErrorType.MaxMinFontSize, Main.Analysis.Options.MinMaxFontSize, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A = limit;
		if (size > (float)limit)
		{
			while (true)
			{
				switch (6)
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
			((BaseError)this).Title = AH.A(17837);
			((BaseError)this).Subtitle = AH.A(17862) + limit + AH.A(17909);
		}
		else
		{
			((BaseError)this).Title = AH.A(17914);
			((BaseError)this).Subtitle = AH.A(17943) + limit + AH.A(17909);
		}
		((BaseError)this).Tooltip = AH.A(17990);
	}
}
