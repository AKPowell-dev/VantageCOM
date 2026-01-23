using System;
using System.Collections.Generic;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderMarginsMismatch
{
	public void Check(Slide sld, Shape shp, Shape placeholder)
	{
		List<string> list = new List<string>();
		TextFrame2 textFrame = placeholder.TextFrame2;
		try
		{
			TextFrame2 textFrame2 = shp.TextFrame2;
			if (textFrame.MarginLeft != textFrame2.MarginLeft)
			{
				while (true)
				{
					switch (1)
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
				list.Add(string.Format(AH.A(14294), textFrame2.MarginLeft));
			}
			if (textFrame.MarginRight != textFrame2.MarginRight)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				list.Add(string.Format(AH.A(14928), textFrame2.MarginRight));
			}
			if (textFrame.MarginTop != textFrame2.MarginTop)
			{
				list.Add(string.Format(AH.A(14263), textFrame2.MarginTop));
			}
			if (textFrame.MarginBottom != textFrame2.MarginBottom)
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
				list.Add(string.Format(AH.A(14963), textFrame2.MarginBottom));
			}
			if (list.Count > 0)
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderMarginsMismatch(sld, shp, string.Join(AH.A(14258), list.ToArray())));
			}
			textFrame2 = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		textFrame = null;
	}
}
