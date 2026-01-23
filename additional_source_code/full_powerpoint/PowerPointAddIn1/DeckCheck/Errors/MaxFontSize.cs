using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MaxFontSize : MinMaxFontSize
{
	public MaxFontSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int limit)
		: base(sld, shp, listRanges, limit)
	{
		((BaseError)this).Title = AH.A(17837);
		((BaseError)this).Subtitle = AH.A(17862) + limit + AH.A(17909);
		((BaseError)this).Tooltip = AH.A(31672);
	}
}
