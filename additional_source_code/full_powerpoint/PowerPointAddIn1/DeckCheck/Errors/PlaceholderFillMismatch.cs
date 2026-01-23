using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderFillMismatch : BaseError
{
	[CompilerGenerated]
	private new Microsoft.Office.Interop.PowerPoint.FillFormat A;

	private Microsoft.Office.Interop.PowerPoint.FillFormat MasterFill
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

	public PlaceholderFillMismatch(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.FillFormat fill)
		: base(ErrorType.PlaceholderFillMismatch, Main.Analysis.Options.CheckPlaceholderFillMismatch, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(38454);
		((BaseError)this).Subtitle = AH.A(38505);
		MasterFill = fill;
		if (fill.Visible != MsoTriState.msoTrue || fill.Type == MsoFillType.msoFillSolid)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (fill.Type != MsoFillType.msoFillPatterned)
			{
				((BaseError)this).HasFix = false;
			}
			return;
		}
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		Microsoft.Office.Interop.PowerPoint.FillFormat fill = base.Shape.Fill;
		MsoFillType type = MasterFill.Type;
		if (type != MsoFillType.msoFillSolid)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (type == MsoFillType.msoFillPatterned)
			{
				fill.Patterned(MasterFill.Pattern);
			}
		}
		else
		{
			fill.Solid();
		}
		fill.ForeColor.RGB = MasterFill.ForeColor.RGB;
		fill.BackColor.RGB = MasterFill.BackColor.RGB;
		try
		{
			fill.Transparency = MasterFill.Transparency;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		fill.Visible = MasterFill.Visible;
		fill = null;
	}
}
