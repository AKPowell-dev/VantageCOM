using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingSmartArtBorderColor : BaseColorError
{
	public NonconformingSmartArtBorderColor(object shp, int intColor, List<Microsoft.Office.Core.Shape> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteBorder, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).OfficeShapes = listShapes;
		((BaseError)this).Title = XC.A(30803);
		((BaseError)this).Subtitle = XC.A(30858);
	}

	public override void FixAction(Color color)
	{
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27753));
		int rGB = ColorTranslator.ToOle(color);
		IEnumerator<Microsoft.Office.Core.Shape> enumerator = default(IEnumerator<Microsoft.Office.Core.Shape>);
		try
		{
			enumerator = ((BaseError)this).OfficeShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.Line.ForeColor.RGB = rGB;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}
}
