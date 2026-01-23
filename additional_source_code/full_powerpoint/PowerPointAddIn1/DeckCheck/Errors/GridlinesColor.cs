using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class GridlinesColor : BaseColorError
{
	private new readonly XlAxisType m_A;

	private new readonly XlAxisGroup m_A;

	private new readonly bool m_A;

	public GridlinesColor(Slide sld, Shape shp, int intColor, PlotArea plot, Severity sev, XlAxisType axisType, XlAxisGroup axisGroup, bool areMajor)
		: base(ErrorType.ColorPaletteChartGridlines, sev, sld, shp, intColor)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		base.PlotArea = plot;
		this.m_A = axisType;
		this.m_A = axisGroup;
		this.m_A = areMajor;
		string[] array = new string[3];
		object obj;
		if (this.m_A != XlAxisType.xlCategory)
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
			if (this.m_A != XlAxisType.xlSeriesAxis)
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
				obj = "";
			}
			else
			{
				obj = AH.A(26301);
			}
		}
		else
		{
			obj = AH.A(26314);
		}
		array[0] = (string)obj;
		array[1] = ((this.m_A == XlAxisGroup.xlSecondary) ? AH.A(26323) : "");
		array[2] = (this.m_A ? "" : AH.A(26342));
		string text = Strings.Join(array.Where([SpecialName] (string A) => Operators.CompareString(A, "", TextCompare: false) != 0).ToArray(), AH.A(14622));
		if (Operators.CompareString(text, "", TextCompare: false) != 0)
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
			text = string.Format(AH.A(26353), text);
		}
		((BaseError)this).Title = string.Format(AH.A(26366), text);
		((BaseError)this).Subtitle = AH.A(26403);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int b = ColorTranslator.ToOle(color);
		Axis axis = (Axis)base.Shape.Chart.Axes(this.m_A);
		if (this.m_A)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(axis.MajorGridlines, b);
		}
		else
		{
			A(axis.MinorGridlines, b);
		}
		axis = null;
	}

	private static void A(Gridlines A, int B)
	{
		A.Format.Line.ForeColor.RGB = B;
	}
}
