using System.Collections.Generic;
using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SmartArtFillColor : BaseColorError
{
	public SmartArtFillColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, int intColor, List<Microsoft.Office.Core.Shape> listShapes, Severity sev)
		: base(ErrorType.ColorPaletteFill, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).OfficeShapes = listShapes;
		((BaseError)this).Title = AH.A(26041);
		((BaseError)this).Subtitle = AH.A(26062);
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
