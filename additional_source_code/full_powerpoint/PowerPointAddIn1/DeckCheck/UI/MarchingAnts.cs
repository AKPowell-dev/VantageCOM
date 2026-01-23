using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros.Proofing.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.UI;

public sealed class MarchingAnts
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<TextRange2, float> A;

		public static Func<TextRange2, float> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal float A(TextRange2 A)
		{
			return A.BoundTop;
		}

		[SpecialName]
		internal float B(TextRange2 A)
		{
			return A.BoundLeft;
		}
	}

	public static Rect GetShapeRectangle(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return GetObjectRectangle(shp.Left, shp.Top, shp.Width, shp.Height);
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

	public static Rect GetPlotAreaOuterRectangle(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		try
		{
			PlotArea plotArea = shp.Chart.PlotArea;
			return GetObjectRectangle((float)((double)shp.Left + plotArea.Left - 10.0), (float)((double)shp.Top + plotArea.Top - 10.0), (float)(plotArea.Width + 20.0), (float)(plotArea.Height + 20.0));
		}
		finally
		{
			PlotArea plotArea = null;
		}
	}

	public static Rect GetLegendRectangle(Legend leg, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(leg.Left + (double)sngLeftOffset), (float)(leg.Top + (double)sngTopOffset), (float)leg.Width, (float)leg.Height);
	}

	public static Rect GetLegendEntryRectangle(Microsoft.Office.Core.LegendEntry legEntry, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(legEntry.Left + (double)sngLeftOffset), (float)(legEntry.Top + (double)sngTopOffset), (float)legEntry.Width, (float)legEntry.Height);
	}

	public static Rect GetLegendKeyRectangle(IMsoLegendKey legKey, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle((float)(legKey.Left + (double)sngLeftOffset), (float)(legKey.Top + (double)sngTopOffset), (float)legKey.Width, (float)legKey.Height);
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
		double num = 5.0;
		return GetObjectRectangle((float)(axis.Left + (double)sngLeftOffset - num), (float)(axis.Top + (double)sngTopOffset - num), (float)(axis.Width + 2.0 * num), (float)(axis.Height + 2.0 * num));
	}

	public static Rect GetTextRangeRectangle(TextRange2 rng, float sngLeftOffset, float sngTopOffset)
	{
		return GetObjectRectangle(sngLeftOffset + rng.BoundLeft, sngTopOffset + rng.BoundTop, rng.BoundWidth, rng.BoundHeight);
	}

	public static Rect GetTextFrameRectangle(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shp.TextFrame2;
		return GetObjectRectangle(shp.Left + textFrame.MarginLeft, shp.Top + textFrame.MarginTop, shp.Width - textFrame.MarginLeft - textFrame.MarginRight, shp.Height - textFrame.MarginTop - textFrame.MarginBottom);
	}

	public static Rect GetLegendKeyRectangle(Microsoft.Office.Interop.PowerPoint.Shape shp, IMsoLegendKey legendKey)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(legendKey.Parent);
		Microsoft.Office.Core.LegendEntry obj = objectValue as Microsoft.Office.Core.LegendEntry;
		double width;
		if (obj == null)
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
			Microsoft.Office.Interop.PowerPoint.LegendEntry obj2 = objectValue as Microsoft.Office.Interop.PowerPoint.LegendEntry;
			if (obj2 == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				width = legendKey.Width;
			}
			else
			{
				width = obj2.Width;
			}
		}
		else
		{
			width = obj.Width;
		}
		double num = width;
		Microsoft.Office.Core.LegendEntry obj3 = objectValue as Microsoft.Office.Core.LegendEntry;
		double height;
		if (obj3 == null)
		{
			Microsoft.Office.Interop.PowerPoint.LegendEntry obj4 = objectValue as Microsoft.Office.Interop.PowerPoint.LegendEntry;
			if (obj4 == null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				height = legendKey.Height;
			}
			else
			{
				height = obj4.Height;
			}
		}
		else
		{
			height = obj3.Height;
		}
		double num2 = height;
		return GetObjectRectangle((float)((double)shp.Left + legendKey.Left), (float)((double)shp.Top + legendKey.Top), (float)num, (float)num2);
	}

	public static Rect GetChartPointRectangle(Microsoft.Office.Interop.PowerPoint.Shape shp, ChartPoint point)
	{
		Rect objectRectangle;
		try
		{
			double num = 0.0;
			try
			{
				num = Math.Max(point.MarkerSize, Math.Max(point.Width, point.Height));
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			if (num < 5.0)
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
				num = 5.0;
			}
			double num2 = num + 8.0 + wpfMarchingAnts.AntThickness;
			Rect rect = new Rect(point.Left - num2 / 2.0, point.Top - num2 / 2.0, num2, num2);
			objectRectangle = GetObjectRectangle((float)((double)shp.Left + rect.Left), (float)((double)shp.Top + rect.Top), (float)rect.Width, (float)rect.Height);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			objectRectangle = GetObjectRectangle(shp.Left, shp.Top, shp.Width, shp.Height);
			ProjectData.ClearProjectError();
		}
		return objectRectangle;
	}

	public static Rect GetObjectRectangle(float sngLeft, float sngTop, float sngWidth, float sngHeight)
	{
		Main.EnsureSlidePaneActive();
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		checked
		{
			int num = (int)Math.Round((double)activeWindow.PointsToScreenPixelsX(sngLeft) / Pane.Dpi.X);
			int num2 = (int)Math.Round((double)activeWindow.PointsToScreenPixelsY(sngTop) / Pane.Dpi.Y);
			Rect result = new Rect(num, num2, (double)activeWindow.PointsToScreenPixelsX(sngLeft + sngWidth) / Pane.Dpi.X - (double)num, (double)activeWindow.PointsToScreenPixelsY(sngTop + sngHeight) / Pane.Dpi.Y - (double)num2);
			activeWindow = null;
			return result;
		}
	}

	public static bool UseRelativePosition(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.HasTable != MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return shp.HasSmartArt == MsoTriState.msoTrue;
				}
			}
		}
		return true;
	}

	public static float ChartLeftOffset(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return (float)((double)shp.Left + shp.Chart.ChartArea.Left);
	}

	public static float ChartTopOffset(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return (float)((double)shp.Top + shp.Chart.ChartArea.Top);
	}

	public static Rect TextRangesTopLeft(IList<TextRange2> listRanges, float leftOffset, float topOffset)
	{
		Func<TextRange2, float> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (TextRange2 A) => A.BoundTop);
		}
		else
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
			keySelector = _Closure_0024__.A;
		}
		IOrderedEnumerable<TextRange2> source = listRanges.OrderBy(keySelector);
		Func<TextRange2, float> keySelector2;
		if (_Closure_0024__.B == null)
		{
			keySelector2 = (_Closure_0024__.B = [SpecialName] (TextRange2 A) => A.BoundLeft);
		}
		else
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
			keySelector2 = _Closure_0024__.B;
		}
		TextRange2 textRange = source.ThenBy(keySelector2).ToList()[0];
		return GetObjectRectangle(textRange.BoundLeft + leftOffset, textRange.BoundTop + topOffset, 1f, 1f);
	}
}
