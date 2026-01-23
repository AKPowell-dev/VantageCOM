using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SeriesPointMarkerColor : BaseColorError
{
	private new readonly bool? A;

	public SeriesPointMarkerColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, ChartPoint pt, Severity sev, bool? isFore)
		: base(ErrorType.ColorPaletteBorder, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).ChartPoint = pt;
		A = isFore;
		if (!A.HasValue)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((BaseError)this).Title = AH.A(23896);
					((BaseError)this).Subtitle = AH.A(23947);
					return;
				}
			}
		}
		if (A == true)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					((BaseError)this).Title = AH.A(24087);
					((BaseError)this).Subtitle = AH.A(24152);
					return;
				}
			}
		}
		((BaseError)this).Title = AH.A(24306);
		((BaseError)this).Subtitle = AH.A(24367);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int num = ColorTranslator.ToOle(color);
		if (A.HasValue)
		{
			while (true)
			{
				switch (7)
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
			if (!A.Value)
			{
				goto IL_0067;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		((BaseError)this).ChartPoint.MarkerForegroundColor = num;
		goto IL_0067;
		IL_0067:
		if (A.HasValue)
		{
			if (A.Value)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		((BaseError)this).ChartPoint.MarkerBackgroundColor = num;
	}
}
