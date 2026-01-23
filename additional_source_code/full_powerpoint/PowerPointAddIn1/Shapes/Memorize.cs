using System;
using System.Collections;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Memorize
{
	public static void RecordSizePosition()
	{
		ShapeRange shapeRange;
		try
		{
			shapeRange = Base.SelectedShapes();
			if (shapeRange.Count > 0)
			{
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator = shapeRange.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Shape shape = (Shape)enumerator.Current;
							shape.Tags.Add(AH.A(82719), shape.Height.ToString());
							shape.Tags.Add(AH.A(82754), shape.Width.ToString());
							shape.Tags.Add(AH.A(82787), shape.Top.ToString());
							shape.Tags.Add(AH.A(82816), shape.Left.ToString());
							shape = null;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00f4;
							}
							continue;
							end_IL_00f4:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					Base.LogActivity(AH.A(82847));
					break;
				}
			}
			else
			{
				A();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A();
			ProjectData.ClearProjectError();
		}
		shapeRange = null;
	}

	public static void RestoreSizePosition()
	{
		ShapeRange shapeRange;
		try
		{
			shapeRange = Base.SelectedShapes();
			if (shapeRange.Count > 0)
			{
				NG.A.Application.StartNewUndoEntry();
				{
					IEnumerator enumerator = shapeRange.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Shape shape = (Shape)enumerator.Current;
							try
							{
								shape.Height = Conversions.ToSingle(shape.Tags[AH.A(82719)]);
								shape.Width = Conversions.ToSingle(shape.Tags[AH.A(82754)]);
								shape.Top = Conversions.ToSingle(shape.Tags[AH.A(82787)]);
								shape.Left = Conversions.ToSingle(shape.Tags[AH.A(82816)]);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							shape = null;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				Base.LogActivity(AH.A(82908));
			}
			else
			{
				A();
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			A();
			ProjectData.ClearProjectError();
		}
		shapeRange = null;
	}

	private static void A()
	{
		Forms.WarningMessage(AH.A(82971));
	}
}
