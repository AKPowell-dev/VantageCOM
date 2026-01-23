using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TableCellMargins : BaseError
{
	[CompilerGenerated]
	private new List<CellMargins> A;

	private List<CellMargins> FixOptions
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public TableCellMargins(Slide sld, Shape shp, List<string> listLabels, List<Shape> listShapes, List<CellMargins> listFixes)
		: base(ErrorType.TableCellMargins, ((Settings)Main.Analysis.Options).TableCellMargins, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		base.Shapes = listShapes;
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(37563);
		((BaseError)this).Subtitle = AH.A(37614);
	}

	public override void FixAction(int i)
	{
		CellMargins cellMargins = FixOptions[i];
		NG.A.Application.StartNewUndoEntry();
		foreach (Shape shape in base.Shapes)
		{
			TextFrame2 textFrame = shape.TextFrame2;
			textFrame.MarginTop = cellMargins.Top;
			textFrame.MarginRight = cellMargins.Right;
			textFrame.MarginBottom = cellMargins.Bottom;
			textFrame.MarginLeft = cellMargins.Left;
			_ = null;
		}
	}
}
