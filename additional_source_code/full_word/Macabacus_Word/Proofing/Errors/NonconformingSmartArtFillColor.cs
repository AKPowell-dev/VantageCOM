using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingSmartArtFillColor : BaseColorError
{
	public NonconformingSmartArtFillColor(object shp, int intColor, List<Microsoft.Office.Core.Shape> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteFill, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).OfficeShapes = listShapes;
		((BaseError)this).Title = XC.A(32326);
		((BaseError)this).Subtitle = XC.A(32377);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		int rGB = ColorTranslator.ToOle(color);
		using (IEnumerator<Microsoft.Office.Core.Shape> enumerator = ((BaseError)this).OfficeShapes.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				enumerator.Current.Fill.ForeColor.RGB = rGB;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
