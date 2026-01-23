using System;
using System.Collections;
using System.Drawing;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Proofing;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Colors;

public sealed class FillTransparency
{
	public static void Fix()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
		try
		{
			shapeRange = NG.A.Application.ActiveWindow.Selection.ShapeRange;
			try
			{
				NG.A.Application.StartNewUndoEntry();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = shapeRange.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.FillFormat fill = ((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current).Fill;
						if (fill.Transparency > 0f)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							if (fill.Visible == MsoTriState.msoTrue)
							{
								Color c = Fixes.ConvertToOpaqueColor(fill.ForeColor.RGB, fill.Transparency);
								try
								{
									fill.Transparency = 0f;
									fill.ForeColor.RGB = ColorTranslator.ToOle(c);
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
							}
						}
						fill = null;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00ec;
						}
						continue;
						end_IL_00ec:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(ex4.Message);
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(13321));
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		shapeRange = null;
	}
}
