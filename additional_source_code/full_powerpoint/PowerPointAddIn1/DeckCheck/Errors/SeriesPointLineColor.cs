using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SeriesPointLineColor : BaseColorError
{
	private new readonly bool A;

	public SeriesPointLineColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, ChartPoint pt, Severity sev)
		: base(ErrorType.ColorPaletteBorder, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		A = false;
		((BaseError)this).ChartPoint = pt;
		A = clsCharts.UsesMarkers(shp.Chart);
		((BaseError)this).Title = AH.A(24517);
		((BaseError)this).Subtitle = AH.A(24564);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int num = ColorTranslator.ToOle(color);
		int num3 = default(int);
		int num5 = default(int);
		if (A)
		{
			int markerForegroundColor = ((BaseError)this).ChartPoint.MarkerForegroundColor;
			int markerBackgroundColor = ((BaseError)this).ChartPoint.MarkerBackgroundColor;
			int rGB = ((BaseError)this).ChartPoint.Format.Line.ForeColor.RGB;
			int num2;
			if (!object.Equals(markerForegroundColor, rGB))
			{
				while (true)
				{
					switch (4)
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
				num2 = markerForegroundColor;
			}
			else
			{
				num2 = num;
			}
			num3 = num2;
			int num4;
			if (!object.Equals(markerBackgroundColor, rGB))
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
				num4 = markerBackgroundColor;
			}
			else
			{
				num4 = num;
			}
			num5 = num4;
		}
		((BaseError)this).ChartPoint.Format.Line.ForeColor.RGB = num;
		if (!A)
		{
			return;
		}
		if (((BaseError)this).ChartPoint.MarkerForegroundColor != num3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			((BaseError)this).ChartPoint.MarkerForegroundColor = num3;
		}
		if (((BaseError)this).ChartPoint.MarkerBackgroundColor == num5)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			((BaseError)this).ChartPoint.MarkerBackgroundColor = num5;
			return;
		}
	}
}
