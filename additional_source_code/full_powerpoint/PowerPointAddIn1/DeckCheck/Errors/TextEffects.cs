using System;
using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class TextEffects : BaseError
{
	public TextEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.TextEffects, Main.Analysis.Options.TextEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		A();
		((BaseError)this).TextRanges = listRanges;
	}

	public TextEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, ChartTitle title)
		: base(ErrorType.TextEffects, Main.Analysis.Options.TextEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.ChartTitle = title;
	}

	public TextEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, AxisTitle title)
		: base(ErrorType.TextEffects, Main.Analysis.Options.TextEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.AxisTitle = title;
	}

	public TextEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Legend legend)
		: base(ErrorType.TextEffects, Main.Analysis.Options.TextEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.Legend = legend;
	}

	public TextEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, DataTable tbl)
		: base(ErrorType.TextEffects, Main.Analysis.Options.TextEffects, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		A();
		base.DataTable = tbl;
	}

	private void A()
	{
		BaseError val = (BaseError)(object)this;
		Errors.TextEffects(ref val);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		if (base.AxisTitle != null)
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
					A(base.AxisTitle.Format.TextFrame2);
					return;
				}
			}
		}
		if (base.ChartTitle != null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					A(base.ChartTitle.Format.TextFrame2);
					return;
				}
			}
		}
		if (base.Legend != null)
		{
			A(base.Legend.Format.TextFrame2);
			return;
		}
		if (base.DataTable != null)
		{
			A(base.DataTable.Format.TextFrame2.TextRange.Font);
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = base.Shape;
		try
		{
			Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeD = shape.TextFrame2.ThreeD;
			threeD.ResetRotation();
			threeD.BevelTopType = MsoBevelType.msoBevelNone;
			threeD.BevelBottomType = MsoBevelType.msoBevelNone;
			threeD.Visible = MsoTriState.msoFalse;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				A(current.Font);
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		shape = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.TextFrame2 A)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.ThreeDFormat threeD = A.ThreeD;
			threeD.ResetRotation();
			threeD.BevelTopType = MsoBevelType.msoBevelNone;
			threeD.BevelBottomType = MsoBevelType.msoBevelNone;
			threeD.Visible = MsoTriState.msoFalse;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		foreach (TextRange2 item in A.TextRange.get_Runs(-1, -1))
		{
			this.A(item.Font);
		}
	}

	private void A(Font2 A)
	{
		B(A);
		C(A);
		D(A);
		E(A);
	}

	private void B(Font2 A)
	{
		try
		{
			A.Shadow.Visible = MsoTriState.msoFalse;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(Font2 A)
	{
		try
		{
			A.Reflection.Type = MsoReflectionType.msoReflectionTypeNone;
			A.Reflection.Size = 0f;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void D(Font2 A)
	{
		try
		{
			A.Glow.Radius = 0f;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void E(Font2 A)
	{
		try
		{
			A.SoftEdgeFormat = MsoSoftEdgeType.msoSoftEdgeTypeNone;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
