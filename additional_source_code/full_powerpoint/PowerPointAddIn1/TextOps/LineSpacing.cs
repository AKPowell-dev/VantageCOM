using System;
using System.Collections;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TextOps;

public sealed class LineSpacing
{
	public static void Increase()
	{
		A(A, AH.A(155566));
	}

	public static void Decrease()
	{
		A(B, AH.A(155609));
	}

	private static void A(Action<ParagraphFormat2> A, string B)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = NG.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		int D = 0;
		try
		{
			if (!selection.HasChildShapeRange)
			{
				{
					IEnumerator enumerator = selection.ShapeRange.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							LineSpacing.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, application, A, ref D);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_00d5;
							}
							continue;
							end_IL_00d5:
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
			}
			else
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = selection.ChildShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						LineSpacing.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, application, A, ref D);
					}
					while (true)
					{
						switch (3)
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
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (4)
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
		if (D > 0)
		{
			Base.LogActivity(B);
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Application B, Action<ParagraphFormat2> C, ref int D)
	{
		checked
		{
			if (A.Type != MsoShapeType.msoGroup)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (A.HasTextFrame == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									if (D == 0)
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
										B.StartNewUndoEntry();
									}
									C(A.TextFrame2.TextRange.ParagraphFormat);
									D++;
									return;
								}
							}
						}
						return;
					}
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					LineSpacing.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B, C, ref D);
				}
				while (true)
				{
					switch (1)
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
	}

	private static void A(ParagraphFormat2 A)
	{
		A.SpaceAfter += 1f;
	}

	private static void B(ParagraphFormat2 A)
	{
		A.SpaceAfter -= Math.Max(0f, 1f);
	}
}
