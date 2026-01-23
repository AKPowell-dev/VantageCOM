using System;
using System.Collections;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class ConvertToPicture
{
	public static void SelectedObjects()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
		Application application = NG.A.Application;
		application.StartNewUndoEntry();
		Slide slide;
		try
		{
			Selection selection = application.ActiveWindow.Selection;
			slide = selection.SlideRange[1];
			IEnumerator enumerator = selection.ShapeRange.GetEnumerator();
			try
			{
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
						try
						{
							MsoShapeType type = shape.Type;
							if (type <= MsoShapeType.msoEmbeddedOLEObject)
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
								if (type == MsoShapeType.msoChart || type == MsoShapeType.msoEmbeddedOLEObject)
								{
									goto IL_00ca;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0093;
									}
									continue;
									end_IL_0093:
									break;
								}
								continue;
							}
							if (type == MsoShapeType.msoLinkedOLEObject)
							{
								goto IL_00ca;
							}
							Microsoft.Office.Interop.PowerPoint.Shape shape2;
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								if (type != MsoShapeType.msoPlaceholder)
								{
									while (true)
									{
										switch (7)
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
								}
								else if (shape.PlaceholderFormat.ContainedType == MsoShapeType.msoChart)
								{
									shape2 = A(slide, shape);
									shape2.Cut();
									A(slide);
									NG.A.Application.ActiveWindow.View.Paste();
									shape2 = null;
								}
								break;
							}
							goto end_IL_0068;
							IL_00ca:
							if (shape.Type == MsoShapeType.msoChart)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									A(slide, shape);
									break;
								}
								continue;
							}
							int count = slide.Shapes.Count;
							shape2 = A(slide, shape);
							if (slide.Shapes.Count > count)
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
								shape2.Cut();
								A(slide);
								NG.A.Application.ActiveWindow.View.Paste();
							}
							shape2 = null;
							end_IL_0068:;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Forms.ErrorMessage(ex2.Message);
							ProjectData.ClearProjectError();
							break;
						}
						continue;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_01c6;
						}
						continue;
						end_IL_01c6:
						break;
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
			selection = null;
			A();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.WarningMessage(AH.A(73347));
			ProjectData.ClearProjectError();
		}
		application = null;
		slide = null;
	}

	public static void AllEmbeddedWorksheets()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			int i = 0;
			Application application = NG.A.Application;
			application.StartNewUndoEntry();
			try
			{
				try
				{
					enumerator = application.ActivePresentation.Slides.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Slide slide = (Slide)enumerator.Current;
						try
						{
							enumerator2 = slide.Shapes.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shp = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								ConvertEmbedded(slide, shp, ref i);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0091;
								}
								continue;
								end_IL_0091:
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
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00c7;
						}
						continue;
						end_IL_00c7:
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
				if (i > 0)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						Forms.SuccessMessage(AH.A(73404) + i + AH.A(73425));
						break;
					}
				}
				else
				{
					Forms.InfoMessage(AH.A(73534));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
			application = null;
			A();
			return;
		}
	}

	public static void ConvertEmbedded(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, ref int i)
	{
		MsoShapeType type = shp.Type;
		if (type != MsoShapeType.msoEmbeddedOLEObject)
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
			if (type != MsoShapeType.msoLinkedOLEObject)
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
				break;
			}
		}
		int count = sld.Shapes.Count;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A(sld, shp);
		if (sld.Shapes.Count > count)
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
			shape.Cut();
			A(sld.Application.ActiveWindow, sld);
		}
		shape = null;
		checked
		{
			i++;
		}
	}

	public static void AllCharts()
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
		int i = 0;
		Application application = NG.A.Application;
		application.StartNewUndoEntry();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.ActivePresentation.Slides.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					try
					{
						enumerator2 = slide.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shp = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							ConvertChart(slide, shp, ref i);
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
			if (i > 0)
			{
				Forms.SuccessMessage(AH.A(73404) + i + AH.A(73635));
			}
			else
			{
				Forms.InfoMessage(AH.A(73718));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application = null;
		A();
	}

	public static void ConvertChart(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, ref int i)
	{
		MsoShapeType type = shp.Type;
		checked
		{
			if (type != MsoShapeType.msoChart)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (type != MsoShapeType.msoPlaceholder || shp.PlaceholderFormat.ContainedType != MsoShapeType.msoChart)
						{
							return;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								A(sld, shp).Cut();
								A(sld.Application.ActiveWindow, sld);
								i++;
								return;
							}
						}
					}
				}
			}
			A(sld, shp);
			i++;
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		B.Copy();
		Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes.PasteSpecial(PpPasteDataType.ppPasteEnhancedMetafile)[1];
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
		int zOrderPosition = B.ZOrderPosition;
		shape2.Left = B.Left;
		shape2.Top = B.Top;
		shape2.Name = B.Name;
		do
		{
			shape2.ZOrder(MsoZOrderCmd.msoSendBackward);
		}
		while (shape2.ZOrderPosition != zOrderPosition);
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
			int count = B.Tags.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				shape2.Tags.Add(B.Tags.Name(i), B.Tags.Value(i));
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				shape2 = null;
				B.Delete();
				return shape;
			}
		}
	}

	private static void A(DocumentWindow A, Slide B)
	{
		View view = A.View;
		view.GotoSlide(B.SlideIndex);
		ConvertToPicture.A(B);
		view.Paste();
		_ = null;
	}

	private static void A(Slide A)
	{
		A.Shapes[A.Shapes.Count].Select();
	}

	private static void A()
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(73793));
	}
}
