using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ShapeOutsideMargins : BaseError
{
	public ShapeOutsideMargins(object obj)
		: base(ErrorType.ShapeOutsideMargins, ((Settings)Main.Analysis.Options).ShapeOutsideMargins, RuntimeHelpers.GetObjectValue(obj), blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		if (obj is InlineShape)
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
			((BaseError)this).HasFix = false;
		}
		((BaseError)this).Title = XC.A(26282);
		((BaseError)this).Subtitle = XC.A(26325);
	}

	public override void FixAction()
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(26245));
		try
		{
			PageSetup pageSetup = base.Range.Document.PageSetup;
			double num = Math.Round(pageSetup.LeftMargin, 4);
			double num2 = Math.Round(pageSetup.LeftMargin + pageSetup.PageWidth, 4);
			pageSetup = null;
			if (base.Shape != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Shape shape = base.Shape;
					if (Math.Round(shape.Left, 4) < num)
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
						if (shape.Left >= 0f)
						{
							shape.Left = (float)num;
							goto IL_0150;
						}
					}
					if (Math.Round(shape.Left + shape.Width, 4) > num2)
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
						if (Math.Round(shape.Left + shape.Width, 4) <= Math.Round(base.Range.Document.PageSetup.PageWidth, 4))
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
							shape.Left = (float)(num2 - (double)shape.Width);
						}
					}
					goto IL_0150;
					IL_0150:
					shape = null;
					break;
				}
			}
			else
			{
				_ = base.InlineShape;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
