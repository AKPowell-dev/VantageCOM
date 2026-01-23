using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros.Proofing.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.UI;

public sealed class Callout
{
	public static readonly int POINTER_X_OFFSET = 25;

	private static List<Rect> A;

	private static wpfCallout A;

	private static wpfMarchingAnts A;

	private static bool A;

	public static List<Rect> DashBoxes
	{
		get
		{
			return Callout.A;
		}
		set
		{
			Callout.A = value;
		}
	}

	public static wpfCallout Dialog
	{
		get
		{
			return Callout.A;
		}
		set
		{
			Callout.A = value;
		}
	}

	public static wpfMarchingAnts MarchingAnts
	{
		get
		{
			return Callout.A;
		}
		set
		{
			Callout.A = value;
		}
	}

	public static bool DoNotClose
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public static void RemoveMarchingAnts()
	{
		if (MarchingAnts == null)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			MarchingAnts.CloseByCode = true;
			((System.Windows.Window)(object)MarchingAnts).Close();
			MarchingAnts = null;
			return;
		}
	}

	public static bool UseRelativePosition(BaseError err)
	{
		if (err.Shape != null)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return err.Shape.HasSmartArt == MsoTriState.msoTrue;
				}
			}
		}
		if (err.InlineShape != null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return err.InlineShape.HasSmartArt == MsoTriState.msoTrue;
				}
			}
		}
		if (err.Table != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		return false;
	}

	public static void Reposition(wpfCallout frm, double dblLeft, double dblTop)
	{
		wpfCallout wpfCallout2 = frm;
		wpfCallout2.Left = dblLeft - (double)POINTER_X_OFFSET;
		wpfCallout2.Top = dblTop - wpfCallout2.gridMain.ActualHeight - wpfCallout2.gridMain.Margin.Top;
		if (MarchingAnts != null)
		{
			((System.Windows.Window)(object)MarchingAnts).Top = wpfCallout2.Top + wpfCallout2.ActualHeight;
			((System.Windows.Window)(object)MarchingAnts).Left = dblLeft - wpfCallout2.XOffset;
		}
		wpfCallout2 = null;
	}

	public static Rect GetObjectRectangle(object obj)
	{
		PC.A.Application.ActiveWindow.GetPoint(out var ScreenPixelsLeft, out var ScreenPixelsTop, out var ScreenPixelsWidth, out var ScreenPixelsHeight, RuntimeHelpers.GetObjectValue(obj));
		return new Rect(ScreenPixelsLeft, ScreenPixelsTop, ScreenPixelsWidth, ScreenPixelsHeight);
	}

	public static Rect GetShapeRectangle(Microsoft.Office.Core.Shape shp)
	{
		return GetObjectRectangle(shp.Left, shp.Top, shp.Width, shp.Height);
	}

	public static Rect GetLabelRectangle(IMsoDataLabel lbl, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(lbl.Left + (double)sngLeftOffset), (float)(lbl.Top + (double)sngTopOffset), (float)lbl.Width, (float)lbl.Height);
	}

	public static Rect GetPlotAreaRectangle(PlotArea plot, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(plot.InsideLeft + (double)sngLeftOffset), (float)(plot.InsideTop + (double)sngTopOffset), (float)plot.InsideWidth, (float)plot.InsideHeight);
	}

	public static Rect GetLegendRectangle(Legend leg, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(leg.Left + (double)sngLeftOffset), (float)(leg.Top + (double)sngTopOffset), (float)leg.Width, (float)leg.Height);
	}

	public static Rect GetChartTitleRectangle(ChartTitle title, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(title.Left + (double)sngLeftOffset), (float)(title.Top + (double)sngTopOffset), (float)title.Width, (float)title.Height);
	}

	public static Rect GetAxisTitleRectangle(AxisTitle title, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(title.Left + (double)sngLeftOffset), (float)(title.Top + (double)sngTopOffset), (float)title.Width, (float)title.Height);
	}

	public static Rect GetAxisRectangle(Axis axis, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(axis.Left + (double)sngLeftOffset), (float)(axis.Top + (double)sngTopOffset), (float)axis.Width, (float)axis.Height);
	}

	public static Rect GetTextRangeRectangle(TextRange2 rng, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle(sngLeftOffset + rng.BoundLeft, sngTopOffset + rng.BoundTop, rng.BoundWidth, rng.BoundHeight);
	}

	public static Rect GetObjectRectangle(float sngLeft, float sngTop, float sngWidth, float sngHeight)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		Microsoft.Office.Interop.Word.Application application2 = application;
		object fVertical = false;
		checked
		{
			int num = (int)Math.Round(application2.PointsToPixels(sngLeft, ref fVertical));
			Microsoft.Office.Interop.Word.Application application3 = application;
			fVertical = true;
			int num2 = (int)Math.Round(application3.PointsToPixels(sngTop, ref fVertical));
			double x = num;
			double y = num2;
			Microsoft.Office.Interop.Word.Application application4 = application;
			float points = sngLeft + sngWidth;
			fVertical = false;
			double width = application4.PointsToPixels(points, ref fVertical);
			Microsoft.Office.Interop.Word.Application application5 = application;
			float points2 = sngTop + sngHeight;
			fVertical = true;
			Rect result = new Rect(x, y, width, application5.PointsToPixels(points2, ref fVertical));
			application = null;
			return result;
		}
	}
}
