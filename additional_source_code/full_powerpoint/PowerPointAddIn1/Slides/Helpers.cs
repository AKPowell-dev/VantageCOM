using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Slides;

public sealed class Helpers
{
	public static readonly string TAG_SLIDE_TYPE = AH.A(118835);

	public static SlideType GetSlideType(Slide sld)
	{
		Microsoft.Office.Interop.PowerPoint.Shape slideTypeShape = GetSlideTypeShape(sld.CustomLayout.Shapes);
		SlideType result;
		if (slideTypeShape != null)
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
			string value = slideTypeShape.Tags[TAG_SLIDE_TYPE].ToString();
			slideTypeShape = null;
			result = (SlideType)Conversions.ToInteger(value);
		}
		else
		{
			try
			{
				result = (SlideType)Conversions.ToInteger(sld.Tags[TAG_SLIDE_TYPE].ToString());
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result = SlideType.Content;
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	public static SlideType GetLayoutType(CustomLayout lay)
	{
		Microsoft.Office.Interop.PowerPoint.Shape slideTypeShape = GetSlideTypeShape(lay.Shapes);
		if (slideTypeShape != null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					string value = slideTypeShape.Tags[TAG_SLIDE_TYPE].ToString();
					slideTypeShape = null;
					return (SlideType)Conversions.ToInteger(value);
				}
				}
			}
		}
		return SlideType.Content;
	}

	public static void DesignateSlideAsType(Slide sld, SlideType sldType)
	{
		try
		{
			A(sld);
			DesignateLayoutAsType(sld.CustomLayout, sldType);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void DesignateLayoutAsType(CustomLayout lay, SlideType sldType)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = lay.Shapes.AddShape(MsoAutoShapeType.msoShapeOctagon, 0f, 0f, 1f, 1f);
			shape.Visible = MsoTriState.msoFalse;
			shape.Name = AH.A(118808);
			shape.Fill.Visible = MsoTriState.msoFalse;
			shape.Line.Visible = MsoTriState.msoFalse;
			Tags tags = shape.Tags;
			string tAG_SLIDE_TYPE = TAG_SLIDE_TYPE;
			int num = (int)sldType;
			tags.Add(tAG_SLIDE_TYPE, num.ToString());
			shape.ZOrder(MsoZOrderCmd.msoSendToBack);
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Slide A)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape slideTypeShape = GetSlideTypeShape(A.CustomLayout.Shapes);
			if (slideTypeShape == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				slideTypeShape.Delete();
				slideTypeShape = null;
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape GetSlideTypeShape(Microsoft.Office.Interop.PowerPoint.Shapes shps)
	{
		Microsoft.Office.Interop.PowerPoint.Shape result = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shps.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
					if (shape2.Visible == MsoTriState.msoFalse && Operators.CompareString(shape2.Tags[TAG_SLIDE_TYPE], "", TextCompare: false) != 0)
					{
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
							result = shape;
							break;
						}
						break;
					}
					shape2 = null;
					continue;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0067;
					}
					continue;
					end_IL_0067:
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return result;
	}

	public static bool IsSpecialSlide(Slide sld)
	{
		return A(GetSlideType(sld));
	}

	public static bool IsSpecialLayout(CustomLayout lay)
	{
		return A(GetLayoutType(lay));
	}

	private static bool A(SlideType A)
	{
		if ((uint)(A - 1) > 4u)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if ((uint)(A - 7) > 3u)
			{
				return false;
			}
		}
		return true;
	}

	public static CustomLayout GetLayout(Microsoft.Office.Interop.PowerPoint.Presentation pres, SlideType type)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		CustomLayout customLayout = null;
		IEnumerator enumerator = pres.Designs.GetEnumerator();
		try
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Design design = (Design)enumerator.Current;
					try
					{
						enumerator2 = design.SlideMaster.CustomLayouts.GetEnumerator();
						while (true)
						{
							if (enumerator2.MoveNext())
							{
								CustomLayout customLayout2 = (CustomLayout)enumerator2.Current;
								shape = GetSlideTypeShape(customLayout2.Shapes);
								if (shape == null)
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
								try
								{
									if (Conversions.ToInteger(shape.Tags[TAG_SLIDE_TYPE].ToString()) == (int)type)
									{
										customLayout = customLayout2;
										break;
									}
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								continue;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_00b4;
								}
								continue;
								end_IL_00b4:
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
					shape = null;
					if (customLayout == null)
					{
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00e5;
						}
						continue;
						end_IL_00e5:
						break;
					}
					break;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_00fe;
					}
					continue;
					end_IL_00fe:
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
		return customLayout;
	}

	public static void SelectMultipleSlides(Microsoft.Office.Interop.PowerPoint.Presentation pres, IEnumerable<int> listIndicies)
	{
		Pane pane = pres.Application.ActiveWindow.Panes[1];
		if (pane.ViewType == PpViewType.ppViewThumbnails)
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
			if (pane.Active != MsoTriState.msoTrue)
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
				pane.Activate();
			}
		}
		pane = null;
		pres.Slides.Range(listIndicies.ToArray()).Select();
	}

	internal static List<Slide> A(Application A)
	{
		List<Slide> list = new List<Slide>();
		try
		{
			DocumentWindow activeWindow = A.ActiveWindow;
			Selection selection = activeWindow.Selection;
			object obj;
			if (selection == null)
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
				obj = null;
			}
			else
			{
				obj = selection.SlideRange;
			}
			SlideRange slideRange = (SlideRange)obj;
			PpSelectionType? obj2;
			if (selection == null)
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
				obj2 = null;
			}
			else
			{
				obj2 = selection.Type;
			}
			if (!object.Equals(obj2, PpSelectionType.ppSelectionNone))
			{
				goto IL_00a4;
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
			if (slideRange == null)
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
			}
			else if (slideRange.Count > 0)
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
				goto IL_00a4;
			}
			View view = activeWindow.View;
			object obj3;
			if (view == null)
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
				obj3 = null;
			}
			else
			{
				obj3 = view.Slide;
			}
			if (obj3 is Slide item)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					list.Add(item);
					break;
				}
			}
			goto end_IL_0006;
			IL_00a4:
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = slideRange.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Slide item2 = (Slide)enumerator.Current;
					list.Add(item2);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_00d1;
					}
					continue;
					end_IL_00d1:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			end_IL_0006:;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Slide slide = null;
			View view = null;
			Selection selection = null;
			DocumentWindow activeWindow = null;
		}
		return list;
	}

	public static int GetSlideIndex()
	{
		bool flag = false;
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		if (activeWindow.Selection.Type == PpSelectionType.ppSelectionNone)
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
			try
			{
				_ = activeWindow.Selection.SlideRange.Count;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				flag = true;
				ProjectData.ClearProjectError();
			}
		}
		if (activeWindow.Panes[1].ViewType == PpViewType.ppViewSlideSorter || activeWindow.Panes[1].ViewType == PpViewType.ppViewThumbnails)
		{
			if (flag)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						if (activeWindow.Panes[1].ViewType == PpViewType.ppViewSlideSorter)
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
							activeWindow.ViewType = PpViewType.ppViewSlide;
							activeWindow.ViewType = PpViewType.ppViewSlideSorter;
						}
						else
						{
							activeWindow.ViewType = PpViewType.ppViewSlide;
							activeWindow.ViewType = PpViewType.ppViewNormal;
						}
						return activeWindow.Selection.SlideRange.SlideIndex;
					}
				}
			}
			return activeWindow.Selection.SlideRange[1].SlideIndex;
		}
		return activeWindow.Selection.SlideRange.SlideIndex;
	}

	public static string GetSlideTitle(Slide sld)
	{
		if (sld.Shapes.HasTitle == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return sld.Shapes.Title.TextFrame.TextRange.Text;
				}
			}
		}
		return "";
	}
}
