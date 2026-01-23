using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SmartArtBorderColor : BaseColorError
{
	public SmartArtBorderColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, List<Microsoft.Office.Core.Shape> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteBorder, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).OfficeShapes = listShapes;
		((BaseError)this).Title = AH.A(24858);
		((BaseError)this).Subtitle = AH.A(24883);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		int rGB = ColorTranslator.ToOle(color);
		IEnumerator<Microsoft.Office.Core.Shape> enumerator = default(IEnumerator<Microsoft.Office.Core.Shape>);
		try
		{
			enumerator = ((BaseError)this).OfficeShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.Line.ForeColor.RGB = rGB;
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
				return;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
