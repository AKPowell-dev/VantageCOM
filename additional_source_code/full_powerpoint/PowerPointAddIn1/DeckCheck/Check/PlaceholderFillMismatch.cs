using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderFillMismatch
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.Shape placeholder)
	{
		Microsoft.Office.Interop.PowerPoint.FillFormat fill = placeholder.Fill;
		if (fill.Type != MsoFillType.msoFillGradient)
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
			if (fill.Visible == shp.Fill.Visible)
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
				if (fill.ForeColor.RGB == shp.Fill.ForeColor.RGB && fill.BackColor.RGB == shp.Fill.BackColor.RGB)
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
					if (fill.Transparency == shp.Fill.Transparency && fill.Type == shp.Fill.Type)
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
						if (fill.Type == MsoFillType.msoFillPatterned)
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
							if (shp.Fill.Type == MsoFillType.msoFillPatterned)
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
								if (fill.Pattern != shp.Fill.Pattern)
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
									goto IL_0133;
								}
							}
						}
						goto IL_0153;
					}
				}
			}
			goto IL_0133;
		}
		goto IL_0153;
		IL_0133:
		Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderFillMismatch(sld, shp, placeholder.Fill));
		goto IL_0153;
		IL_0153:
		fill = null;
	}
}
