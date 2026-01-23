using System;
using System.Collections;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TextOps;

public sealed class Autofit
{
	public static void Toggle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = NG.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		int C = 0;
		try
		{
			if (selection.HasChildShapeRange)
			{
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
					try
					{
						enumerator = selection.ChildShapeRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, application, ref C);
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
					break;
				}
			}
			else
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = selection.ShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, application, ref C);
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		selection = null;
		application = null;
		if (C <= 0)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			Base.LogActivity(AH.A(153879));
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Application B, ref int C)
	{
		checked
		{
			if (A.Type != MsoShapeType.msoGroup)
			{
				if (A.HasTextFrame != MsoTriState.msoTrue)
				{
					return;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (C == 0)
					{
						B.StartNewUndoEntry();
					}
					Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = A.TextFrame2;
					textFrame.WordWrap = MsoTriState.msoTrue;
					if (textFrame.AutoSize == MsoAutoSize.msoAutoSizeShapeToFitText)
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
						textFrame.AutoSize = MsoAutoSize.msoAutoSizeNone;
					}
					else if (textFrame.AutoSize == MsoAutoSize.msoAutoSizeNone)
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
						textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
					}
					else
					{
						try
						{
							if (A.Fill.Visible == MsoTriState.msoTrue)
							{
								goto IL_00b6;
							}
							if (A.Line.Visible == MsoTriState.msoTrue)
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
								goto IL_00b6;
							}
							textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
							goto end_IL_0088;
							IL_00b6:
							textFrame.AutoSize = MsoAutoSize.msoAutoSizeNone;
							end_IL_0088:;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
							ProjectData.ClearProjectError();
						}
					}
					textFrame = null;
					C++;
					return;
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Autofit.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B, ref C);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
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
	}
}
