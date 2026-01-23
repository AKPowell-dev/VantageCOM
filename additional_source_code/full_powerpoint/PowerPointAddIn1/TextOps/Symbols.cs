using System;
using System.Collections;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TextOps;

public sealed class Symbols
{
	public static void Insert(int num)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = NG.A.Application;
			int num2 = 0;
			try
			{
				char c = clsText.ConvertUnicodeToChar(num);
				Selection selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionText)
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
					application.StartNewUndoEntry();
					selection.TextRange2.InsertAfter(Conversions.ToString(c));
				}
				else
				{
					try
					{
						enumerator = selection.ShapeRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							if (shape.HasTextFrame != MsoTriState.msoTrue)
							{
								continue;
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
							if (num2 == 0)
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
								application.StartNewUndoEntry();
							}
							shape.TextFrame2.TextRange.InsertAfter(Conversions.ToString(c));
							num2 = checked(num2 + 1);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_00f1;
							}
							continue;
							end_IL_00f1:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (1)
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
				selection = null;
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(156167));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application = null;
			return;
		}
	}
}
