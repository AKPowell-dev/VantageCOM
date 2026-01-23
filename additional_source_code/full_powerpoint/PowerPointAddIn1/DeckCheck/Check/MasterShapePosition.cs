using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;
using PowerPointAddIn1.MasterShapes;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MasterShapePosition
{
	public void Check(Slide sld, Shape shp)
	{
		string text = "";
		Shape value = null;
		try
		{
			text = shp.Tags[Base.TAG_ID];
			if (text.Length <= 0)
			{
				return;
			}
			if (Base.MyMasterShapes == null)
			{
				Base.C();
			}
			if (!Base.MyMasterShapes.TryGetValue(text, out value))
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
				Shape shape = value;
				if (Math.Round(shp.Top, 4) == Math.Round(shape.Top, 4))
				{
					if (Math.Round(shp.Left, 4) == Math.Round(shape.Left, 4))
					{
						goto IL_00da;
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.MasterShapePosition(sld, shp, shape.Top, shape.Left));
				goto IL_00da;
				IL_00da:
				shape = null;
				value = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
