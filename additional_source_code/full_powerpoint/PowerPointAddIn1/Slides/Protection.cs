using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.MasterShapes;

namespace PowerPointAddIn1.Slides;

public sealed class Protection
{
	public static void LockSlides1()
	{
		A(A: false);
	}

	public static void LockSlides2()
	{
		A(A: true);
	}

	public static void LockSlides3()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		Application application = NG.A.Application;
		new List<int>();
		SlideRange slideRange;
		try
		{
			_ = application.ActivePresentation;
			slideRange = application.ActiveWindow.Selection.SlideRange;
			try
			{
				application.StartNewUndoEntry();
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = slideRange.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						Slide slide = (Slide)enumerator.Current;
						application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
						try
						{
							enumerator2 = slide.Shapes.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator2.Current);
								try
								{
									NewLateBinding.LateSet(objectValue, null, AH.A(69417), new object[1] { true }, null, null);
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
							}
							while (true)
							{
								switch (4)
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
									switch (1)
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
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0118;
						}
						continue;
						end_IL_0118:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				slideRange.Select();
				A();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				Forms.ErrorMessage(ex4.Message);
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			Forms.WarningMessage(AH.A(119072));
			ProjectData.ClearProjectError();
		}
		slideRange = null;
		application = null;
	}

	public static void LockSlidesAll(Microsoft.Office.Interop.PowerPoint.Presentation pres, bool blnImage)
	{
		SlideRange b = pres.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
		List<int> D = new List<int>();
		if (blnImage)
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
			B();
		}
		A(pres, b, blnImage, ref D);
		A();
		b = null;
		D = null;
	}

	private static void A(bool A)
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
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
			Application application = NG.A.Application;
			List<int> D = new List<int>();
			SlideRange slideRange;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			try
			{
				activePresentation = application.ActivePresentation;
				slideRange = application.ActiveWindow.Selection.SlideRange;
				try
				{
					if (A)
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
						B();
					}
					application.StartNewUndoEntry();
					Protection.A(activePresentation, slideRange, A, ref D);
					application.ActiveWindow.View.GotoSlide(D[0]);
					Helpers.SelectMultipleSlides(activePresentation, D);
					Protection.A();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					Forms.ErrorMessage(ex2.Message);
					ProjectData.ClearProjectError();
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(AH.A(119072));
				ProjectData.ClearProjectError();
			}
			slideRange = null;
			activePresentation = null;
			application = null;
			D = null;
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, SlideRange B, bool C, ref List<int> D)
	{
		IEnumerator enumerator = default(IEnumerator);
		CustomLayout pCustomLayout;
		Slide slide2;
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				if (C)
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
					pCustomLayout = Protection.A(slide);
				}
				else
				{
					pCustomLayout = Protection.B(slide);
				}
				slide2 = A.Slides.AddSlide(slide.SlideIndex, pCustomLayout);
				D.Add(slide2.SlideIndex);
				slide.Delete();
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		pCustomLayout = null;
		slide2 = null;
	}

	private static CustomLayout A(Slide A)
	{
		Application application = A.Application;
		A.CustomLayout.Copy();
		CustomLayout customLayout = A.CustomLayout.Design.SlideMaster.CustomLayouts.Paste();
		customLayout.Name = AH.A(119634);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.CustomLayout.Shapes.GetEnumerator();
			while (true)
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape;
				if (enumerator.MoveNext())
				{
					shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (!Numbers.IsSlideNumberPlaceholder(shape))
					{
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
							if (!Protection.A(shape))
							{
								break;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								break;
							}
							goto IL_00a7;
						}
						continue;
					}
					goto IL_00a7;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00bb;
					}
					continue;
					end_IL_00bb:
					break;
				}
				break;
				IL_00a7:
				shape.Visible = MsoTriState.msoFalse;
				break;
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
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = A.Shapes.Placeholders.GetEnumerator();
			while (true)
			{
				if (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
					if (!Numbers.IsSlideNumberPlaceholder(shape2))
					{
						continue;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						shape2.Visible = MsoTriState.msoFalse;
						break;
					}
					break;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0138;
					}
					continue;
					end_IL_0138:
					break;
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
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator3 = A.Shapes.GetEnumerator();
		try
		{
			while (enumerator3.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
				if (Base.A(shape3))
				{
					shape3.Visible = MsoTriState.msoFalse;
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_01a2;
				}
				continue;
				end_IL_01a2:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator3 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		string text = modFunctionsIO.PathGetTempFileName();
		A.Export(text, AH.A(119651), application.ActiveWindow.PointsToScreenPixelsX(A.CustomLayout.Width), application.ActiveWindow.PointsToScreenPixelsY(A.CustomLayout.Height));
		Microsoft.Office.Interop.PowerPoint.Shapes shapes = customLayout.Shapes;
		for (int i = shapes.Count; i >= 1; i = checked(i + -1))
		{
			if (Numbers.IsSlideNumberPlaceholder(shapes[i]))
			{
				continue;
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
			if (Protection.A(shapes[i]))
			{
				continue;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			shapes[i].Delete();
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			shapes.AddPicture2(text, MsoTriState.msoFalse, MsoTriState.msoTrue, 0f, 0f).ZOrder(MsoZOrderCmd.msoSendToBack);
			shapes = null;
			File.Delete(text);
			application = null;
			return customLayout;
		}
	}

	private static CustomLayout B(Slide A)
	{
		CustomLayout customLayout = A.CustomLayout;
		customLayout.Copy();
		CustomLayout customLayout2 = customLayout.Design.SlideMaster.CustomLayouts.Paste();
		_ = null;
		customLayout2.Name = AH.A(119634);
		Protection.A(customLayout2);
		Slide slide = A.Application.ActivePresentation.Slides.AddSlide(A.SlideIndex, A.CustomLayout);
		IEnumerator enumerator = A.Shapes.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (shape.Type == MsoShapeType.msoPlaceholder)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber)
					{
						continue;
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
				}
				try
				{
					shape.Copy();
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = slide.Shapes.Paste()[1];
					shape2.Copy();
					Microsoft.Office.Interop.PowerPoint.Shape shape3 = customLayout2.Shapes.Paste()[1];
					shape3.Top = shape.Top;
					shape3.Left = shape.Left;
					_ = null;
					shape2.Delete();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				finally
				{
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_014e;
				}
				continue;
				end_IL_014e:
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
		Protection.A(customLayout2);
		slide.Delete();
		slide = null;
		return customLayout2;
	}

	private static void A(CustomLayout A)
	{
		Microsoft.Office.Interop.PowerPoint.Placeholders placeholders = A.Shapes.Placeholders;
		for (int i = placeholders.Count; i >= 1; i = checked(i + -1))
		{
			if (Numbers.IsSlideNumberPlaceholder(placeholders[i]) || Protection.A(placeholders[i]))
			{
				continue;
			}
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
			placeholders[i].Delete();
		}
		placeholders = null;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
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
			if (A.TextFrame.HasText == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return Operators.CompareString(A.TextFrame2.TextRange.Text, AH.A(119658), TextCompare: false) == 0;
					}
				}
			}
		}
		return false;
	}

	private static void A()
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)14, AH.A(119665));
	}

	private static void B()
	{
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		if (activeWindow.ActivePane.ViewType != PpViewType.ppViewSlide)
		{
			foreach (Pane pane in activeWindow.Panes)
			{
				if (pane.ViewType != PpViewType.ppViewSlide)
				{
					continue;
				}
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
					pane.Activate();
					break;
				}
				break;
			}
		}
		activeWindow = null;
	}
}
