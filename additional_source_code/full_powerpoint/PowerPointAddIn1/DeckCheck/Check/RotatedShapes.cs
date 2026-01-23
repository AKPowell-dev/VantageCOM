using System;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class RotatedShapes
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
		if (shape.Type != MsoShapeType.msoLine)
		{
			while (true)
			{
				switch (6)
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
			if (!A(shp))
			{
				if (shape.Type == MsoShapeType.msoAutoShape)
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
					try
					{
						float rotation = shape.Rotation;
						if (rotation == 0f)
						{
							goto IL_0153;
						}
						if (!(rotation < 2f))
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
							if (!(rotation > 358f))
							{
								goto IL_0153;
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
						Main.Analysis.Errors.Add(new RotatedShape(sld, shp, rotation, 0f));
						goto end_IL_00fd;
						IL_0153:
						if (rotation == 90f)
						{
							goto IL_019a;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						if (!(rotation < 92f) || !(rotation > 88f))
						{
							goto IL_019a;
						}
						Main.Analysis.Errors.Add(new RotatedShape(sld, shp, rotation, 90f));
						goto end_IL_00fd;
						IL_019a:
						if (rotation == 180f)
						{
							goto IL_01e4;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						if (!(rotation < 182f))
						{
							goto IL_01e4;
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
						if (!(rotation > 178f))
						{
							goto IL_01e4;
						}
						Main.Analysis.Errors.Add(new RotatedShape(sld, shp, rotation, 180f));
						goto end_IL_00fd;
						IL_01e4:
						if (rotation != 270f)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								if (!(rotation < 272f))
								{
									break;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									if (rotation > 268f)
									{
										Main.Analysis.Errors.Add(new RotatedShape(sld, shp, rotation, 270f));
									}
									break;
								}
								break;
							}
						}
						end_IL_00fd:;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				goto IL_0240;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (shape.Rotation == 0f)
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
			if (shape.Height > 0f && A(shape.Height) < 0.08f)
			{
				Main.Analysis.Errors.Add(new CrookedLine(sld, shp, blnHorizontal: true));
			}
			else if (shape.Width > 0f && A(shape.Width) < 0.08f)
			{
				Main.Analysis.Errors.Add(new CrookedLine(sld, shp, blnHorizontal: false));
			}
		}
		goto IL_0240;
		IL_0240:
		shape = null;
	}

	private float A(float A)
	{
		return clsPublish.PointsToInches(A);
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type == MsoShapeType.msoAutoShape && shape.AutoShapeType == MsoAutoShapeType.msoShapeMixed)
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
			if (shape.Line.EndArrowheadStyle != MsoArrowheadStyle.msoArrowheadTriangle)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (shape.Line.BeginArrowheadStyle != MsoArrowheadStyle.msoArrowheadTriangle)
				{
					goto IL_005a;
				}
			}
			return true;
		}
		goto IL_005a;
		IL_005a:
		shape = null;
		return false;
	}
}
