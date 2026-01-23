using System;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Fix;

public sealed class Colors
{
	public static void RecolorChartFont(Func<ChartFormat> formatFunc, int intColor)
	{
		try
		{
			formatFunc().TextFrame2.TextRange.get_Characters(-1, -1).Font.Fill.ForeColor.RGB = intColor;
			try
			{
				if (formatFunc().TextFrame2.TextRange.Font.Fill.ForeColor.RGB == intColor)
				{
					return;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					throw new NotImplementedException(AH.A(47354));
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		finally
		{
		}
	}
}
