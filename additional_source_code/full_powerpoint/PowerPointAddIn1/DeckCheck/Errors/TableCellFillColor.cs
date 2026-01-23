using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TableCellFillColor : BaseColorError
{
	public TableCellFillColor(Slide sld, Shape shp, int intColor, List<Shape> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteFill, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		base.Shapes = listShapes;
		((BaseError)this).Title = AH.A(26041);
		((BaseError)this).Subtitle = AH.A(26062);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int rGB = ColorTranslator.ToOle(color);
		using List<Shape>.Enumerator enumerator = base.Shapes.GetEnumerator();
		while (enumerator.MoveNext())
		{
			enumerator.Current.Fill.ForeColor.RGB = rGB;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}
}
