using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class ImageDistortion : BaseError
{
	public ImageDistortion(Microsoft.Office.Interop.Word.Shape shp, double dblHeight, double dblWidth)
		: base(ErrorType.ImageDistortion, ((Settings)Main.Analysis.Options).ImageDistortion, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.ImageDistortion(ref val, dblHeight, dblWidth);
	}

	public override void FixAction(int i)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(26204));
		if (base.InlineShape != null)
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
			InlineShape inlineShape = base.InlineShape;
			MsoTriState lockAspectRatio = inlineShape.LockAspectRatio;
			inlineShape.LockAspectRatio = MsoTriState.msoFalse;
			float num;
			if (i == 0)
			{
				float width = inlineShape.Width;
				inlineShape.ScaleWidth = 1f;
				num = width / inlineShape.Width;
			}
			else
			{
				float height = inlineShape.Height;
				inlineShape.ScaleHeight = 1f;
				num = height / inlineShape.Height;
			}
			inlineShape.ScaleHeight = num;
			inlineShape.ScaleWidth = num;
			inlineShape.LockAspectRatio = lockAspectRatio;
			inlineShape = null;
		}
		else if (base.Shape != null)
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
			Microsoft.Office.Interop.Word.Shape shape = base.Shape;
			MsoTriState lockAspectRatio = shape.LockAspectRatio;
			shape.LockAspectRatio = MsoTriState.msoFalse;
			float num;
			if (i == 0)
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
				float width2 = shape.Width;
				shape.ScaleWidth(1f, MsoTriState.msoTrue);
				num = width2 / shape.Width;
			}
			else
			{
				float height2 = shape.Height;
				shape.ScaleHeight(1f, MsoTriState.msoTrue);
				num = height2 / shape.Height;
			}
			shape.ScaleHeight(num, MsoTriState.msoTrue);
			shape.ScaleWidth(num, MsoTriState.msoTrue);
			shape.LockAspectRatio = lockAspectRatio;
			shape = null;
		}
		undoRecord.EndCustomRecord();
	}
}
