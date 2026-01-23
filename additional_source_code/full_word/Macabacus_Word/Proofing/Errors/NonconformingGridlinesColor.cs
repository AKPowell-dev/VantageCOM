using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingGridlinesColor : BaseColorError
{
	public NonconformingGridlinesColor(object shp, int intColor, PlotArea plot, Severity sev)
		: base(ErrorType.ColorPaletteChartGridlines, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		((BaseError)this).Title = XC.A(32712);
		((BaseError)this).Subtitle = XC.A(32773);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(29748));
		int b = ColorTranslator.ToOle(color);
		XlAxisType[] array = new XlAxisType[2]
		{
			XlAxisType.xlValue,
			XlAxisType.xlCategory
		};
		foreach (XlAxisType xlAxisType in array)
		{
			if (!Conversions.ToBoolean(base.Shape.Chart.get_HasAxis((object)xlAxisType, RuntimeHelpers.GetObjectValue(Missing.Value))))
			{
				continue;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Axis axis = (Axis)base.Shape.Chart.Axes(xlAxisType);
			if (axis.HasMajorGridlines)
			{
				A(axis.MajorGridlines, b);
			}
			if (axis.HasMinorGridlines)
			{
				A(axis.MinorGridlines, b);
			}
			axis = null;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			undoRecord.EndCustomRecord();
			undoRecord = null;
			return;
		}
	}

	private static void A(Gridlines A, int B)
	{
		A.Format.Line.ForeColor.RGB = B;
	}
}
