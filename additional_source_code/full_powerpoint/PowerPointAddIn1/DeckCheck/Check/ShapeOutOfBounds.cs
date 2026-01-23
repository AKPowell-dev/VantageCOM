using System;
using System.Collections.Generic;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ShapeOutOfBounds
{
	public void Check(Slide sld, Shape shp)
	{
		List<string> list = new List<string>();
		try
		{
			CustomLayout customLayout = sld.CustomLayout;
			Shape shape = shp;
			if (shape.Top < 0f)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (shape.Top + shape.Height > 0f)
				{
					goto IL_00b6;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			if (Math.Round(shape.Top + shape.Height, 4) > Math.Round(customLayout.Height, 4))
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
				if (Math.Round(shape.Top, 4) < Math.Round(customLayout.Height, 4))
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
					goto IL_00b6;
				}
			}
			goto IL_00dc;
			IL_00b6:
			list.Add(string.Format(AH.A(14263), shape.Top));
			goto IL_00dc;
			IL_00dc:
			if (shape.Left < 0f)
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
				if (shape.Left + shape.Width > 0f)
				{
					goto IL_0168;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			if (Math.Round(shape.Left + shape.Width, 4) > Math.Round(customLayout.Width, 4) && Math.Round(shape.Left, 4) < Math.Round(customLayout.Width, 4))
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
				goto IL_0168;
			}
			goto IL_018c;
			IL_0168:
			list.Add(string.Format(AH.A(14294), shape.Left));
			goto IL_018c;
			IL_018c:
			if (list.Count > 0)
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.ShapeOutOfBounds(sld, shp, string.Join(AH.A(14258), list.ToArray())));
			}
			customLayout = null;
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
	}
}
