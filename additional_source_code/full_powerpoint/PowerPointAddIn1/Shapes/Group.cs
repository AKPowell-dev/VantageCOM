using System;
using System.Collections;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Group
{
	public static void Ungroup()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = default(Microsoft.Office.Interop.PowerPoint.ShapeRange);
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		clsRibbon clsRibbon = default(clsRibbon);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num4;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					new clsRibbon();
					goto IL_0008;
				case 377:
					{
						num = num2;
						switch (num3)
						{
						case 2:
							break;
						case 1:
							goto IL_011d;
						default:
							goto end_IL_0000;
						}
						break;
					}
					IL_011d:
					num4 = num + 1;
					num = 0;
					switch (num4)
					{
					case 1:
						break;
					case 2:
						goto IL_0008;
					case 3:
						goto IL_000f;
					case 4:
						goto IL_0032;
					case 5:
						goto IL_0039;
					case 6:
						goto IL_0053;
					case 7:
						goto IL_0058;
					case 8:
						goto IL_0079;
					case 9:
						goto IL_0085;
					case 10:
						goto IL_0097;
					case 11:
						goto IL_009a;
					case 12:
						goto IL_00b2;
					case 13:
						goto IL_00d4;
					case 14:
						goto IL_00de;
					case 15:
						goto IL_00e8;
					case 16:
						goto IL_00eb;
					case 18:
						goto end_IL_0000_2;
					default:
						goto end_IL_0000;
					case 17:
					case 19:
						goto end_IL_0000_3;
					}
					goto default;
					IL_0008:
					ProjectData.ClearProjectError();
					num3 = 2;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					shapeRange = NG.A.Application.ActiveWindow.Selection.ShapeRange;
					goto IL_0032;
					IL_0032:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0039;
					IL_0039:
					num2 = 5;
					enumerator = shapeRange.GetEnumerator();
					goto IL_009d;
					IL_009d:
					if (enumerator.MoveNext())
					{
						shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						goto IL_0053;
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
					goto IL_00b2;
					IL_0085:
					num2 = 9;
					shape2.Ungroup().Select();
					goto IL_0097;
					IL_00b2:
					num2 = 12;
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00d4;
					IL_009a:
					num2 = 11;
					goto IL_009d;
					IL_0079:
					num2 = 8;
					_ = shape2.Name;
					goto IL_0085;
					IL_00d4:
					num2 = 13;
					clsRibbon = new clsRibbon();
					goto IL_00de;
					IL_00de:
					num2 = 14;
					clsRibbon.GroupControlsReset();
					goto IL_00e8;
					IL_00e8:
					clsRibbon = null;
					goto IL_00eb;
					IL_00eb:
					num2 = 16;
					shapeRange = null;
					goto end_IL_0000_3;
					IL_0097:
					shape2 = null;
					goto IL_009a;
					IL_0053:
					num2 = 6;
					shape2 = shape;
					goto IL_0058;
					IL_0058:
					num2 = 7;
					if (shape2.Type == MsoShapeType.msoGroup)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0079;
					}
					goto IL_0097;
					end_IL_0000_2:
					break;
				}
				num2 = 18;
				Interaction.MsgBox(AH.A(81970), MsgBoxStyle.Exclamation, AH.A(5874));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 377;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void Regroup()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = default(Microsoft.Office.Interop.PowerPoint.ShapeRange);
		string text = default(string);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		string text2 = default(string);
		string[] array = default(string[]);
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape3 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		clsRibbon clsRibbon = default(clsRibbon);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				int num4;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					new clsRibbon();
					goto IL_0008;
				case 671:
					{
						num = num2;
						switch (num3)
						{
						case 2:
							break;
						case 1:
							goto IL_0213;
						default:
							goto end_IL_0000;
						}
						break;
					}
					IL_0213:
					num4 = num + 1;
					num = 0;
					switch (num4)
					{
					case 1:
						break;
					case 2:
						goto IL_0008;
					case 3:
						goto IL_000f;
					case 4:
						goto IL_002f;
					case 6:
						goto IL_0052;
					case 7:
						goto IL_005b;
					case 8:
						goto IL_005d;
					case 9:
						goto IL_0064;
					case 10:
						goto IL_0086;
					case 11:
						goto IL_008d;
					case 12:
						goto IL_00ae;
					case 13:
						goto IL_00d8;
					case 14:
						goto IL_00f2;
					case 15:
						goto IL_00fa;
					case 16:
						goto IL_0103;
					case 17:
						goto IL_0120;
					case 18:
						goto IL_0123;
					case 19:
						goto IL_013e;
					case 20:
						goto IL_0160;
					case 21:
						goto IL_0171;
					case 22:
						goto IL_018d;
					case 23:
						goto IL_01be;
					case 24:
						goto IL_01ca;
					case 25:
						goto IL_01cd;
					case 26:
						goto IL_01d7;
					case 27:
						goto IL_01e1;
					case 28:
						goto IL_01e4;
					case 30:
						goto end_IL_0000_2;
					default:
						goto end_IL_0000;
					case 5:
					case 29:
					case 31:
						goto end_IL_0000_3;
					}
					goto default;
					IL_0008:
					ProjectData.ClearProjectError();
					num3 = 2;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					shapeRange = NG.A.Application.ActiveWindow.Selection.ShapeRange;
					goto IL_002f;
					IL_002f:
					num2 = 4;
					if (shapeRange.Count < 2)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0052;
					IL_008d:
					num2 = 11;
					text = shape.Tags[AH.A(82029)];
					goto IL_00ae;
					IL_00ae:
					num2 = 12;
					if ((Operators.CompareString(text2, "", TextCompare: false) == 0) & (Operators.CompareString(text, "", TextCompare: false) != 0))
					{
						goto IL_00d8;
					}
					goto IL_0103;
					IL_00d8:
					num2 = 13;
					array = Strings.Split(text, AH.A(82052));
					goto IL_00f2;
					IL_0052:
					num2 = 6;
					text2 = "";
					goto IL_005b;
					IL_005b:
					num2 = 7;
					goto IL_005d;
					IL_005d:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0064;
					IL_0064:
					num2 = 9;
					enumerator = shapeRange.GetEnumerator();
					goto IL_0126;
					IL_0126:
					if (enumerator.MoveNext())
					{
						shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						goto IL_0086;
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
					goto IL_013e;
					IL_00f2:
					num2 = 14;
					_ = array[0];
					goto IL_00fa;
					IL_013e:
					num2 = 19;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0160;
					IL_0103:
					num2 = 16;
					shape.Tags.Delete(AH.A(82029));
					goto IL_0120;
					IL_0120:
					shape = null;
					goto IL_0123;
					IL_0160:
					num2 = 20;
					shapeRange.Group().Select();
					goto IL_0171;
					IL_0171:
					num2 = 21;
					if (Operators.CompareString(text2, "", TextCompare: false) != 0)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						goto IL_018d;
					}
					goto IL_01cd;
					IL_0123:
					num2 = 18;
					goto IL_0126;
					IL_018d:
					num2 = 22;
					shape3 = NG.A.Application.ActiveWindow.Selection.ShapeRange[1];
					goto IL_01be;
					IL_01be:
					num2 = 23;
					shape3.Name = text2;
					goto IL_01ca;
					IL_01ca:
					shape3 = null;
					goto IL_01cd;
					IL_01cd:
					num2 = 25;
					clsRibbon = new clsRibbon();
					goto IL_01d7;
					IL_01d7:
					num2 = 26;
					clsRibbon.GroupControlsReset();
					goto IL_01e1;
					IL_01e1:
					clsRibbon = null;
					goto IL_01e4;
					IL_01e4:
					num2 = 28;
					shapeRange = null;
					goto end_IL_0000_3;
					IL_00fa:
					num2 = 15;
					text2 = array[1];
					goto IL_0103;
					IL_0086:
					num2 = 10;
					shape = shape2;
					goto IL_008d;
					end_IL_0000_2:
					break;
				}
				num2 = 30;
				Interaction.MsgBox(AH.A(82057), MsgBoxStyle.Exclamation, AH.A(5874));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 671;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}
}
