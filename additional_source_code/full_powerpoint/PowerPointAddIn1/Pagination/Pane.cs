using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Pagination;

public sealed class Pane
{
	public static readonly string PANE_TITLE = AH.A(100389);

	public static readonly string RIBBON_CONTROL = AH.A(100406);

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	private static clsDisplay Dpi
	{
		[CompilerGenerated]
		get
		{
			return Pane.m_A;
		}
		[CompilerGenerated]
		set
		{
			Pane.m_A = value;
		}
	}

	private static Dictionary<int, CustomTaskPane> TaskPanes
	{
		[CompilerGenerated]
		get
		{
			return Pane.m_A;
		}
		[CompilerGenerated]
		set
		{
			Pane.m_A = value;
		}
	} = null;

	public static void Toggle(bool blnPressed)
	{
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Expected O, but got Unknown
		if (blnPressed)
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
			if (!Access.AllowPowerPointOperation((PlanType)5, (Restriction)2, false))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						A();
						return;
					}
				}
			}
		}
		if (!A())
		{
			A();
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		bool flag = false;
		if (blnPressed)
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
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null)
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
			if (TaskPanes.TryGetValue(hWND, out value))
			{
				value.Visible = blnPressed;
				flag = true;
			}
		}
		if (blnPressed)
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
			if (!flag)
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
				ctpPagination control = new ctpPagination();
				value = NG.A.CustomTaskPanes.Add(control, PANE_TITLE, activeWindow);
				clsPanes.A(value, Dpi);
				value.VisibleChanged += A;
				value.Visible = true;
				if (TaskPanes == null)
				{
					TaskPanes = new Dictionary<int, CustomTaskPane>();
				}
				TaskPanes.Add(hWND, value);
			}
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		if (blnPressed)
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
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		}
		activeWindow = null;
		value = null;
		application = null;
	}

	private static bool A()
	{
		bool flag = false;
		bool result;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
		try
		{
			activePresentation = NG.A.Application.ActivePresentation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			result = false;
			ProjectData.ClearProjectError();
			goto IL_0132;
		}
		if (!activePresentation.Final)
		{
			int num = activePresentation.Slides.Count;
			while (true)
			{
				Slide slide;
				if (num >= 1)
				{
					slide = activePresentation.Slides[num];
					if (Helpers.GetSlideType(slide) != SlideType.Blank)
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
						if (slide.Shapes.Count != 0)
						{
							goto IL_00ec;
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
						if (Helpers.IsSpecialSlide(slide))
						{
							goto IL_00ec;
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
					if (!flag)
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
						if (Forms.OkCancelMessage2(AH.A(99930)) != DialogResult.OK)
						{
							slide = null;
							activePresentation = null;
							result = false;
							break;
						}
					}
					flag = true;
					slide.Delete();
					goto IL_00ec;
				}
				if (activePresentation.Slides.Count < 3)
				{
					Forms.WarningMessage(AH.A(100148));
					activePresentation = null;
					result = false;
				}
				else
				{
					activePresentation = null;
					result = true;
				}
				break;
				IL_00ec:
				slide = null;
				num = checked(num + -1);
			}
		}
		else
		{
			MarkedFinalWarning();
			result = false;
		}
		goto IL_0132;
		IL_0132:
		return result;
	}

	public static void MarkedFinalWarning()
	{
		Forms.WarningMessage(AH.A(100278));
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfPaginate wpfPaginate2 = Pane.A((ctpPagination)customTaskPane.Control);
		if (customTaskPane.Visible)
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
			wpfPaginate2.ShowPane();
		}
		else if (!KG.A)
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
			wpfPaginate2.HidePane();
			PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.A();
		}
		customTaskPane = null;
		wpfPaginate2 = null;
	}

	private static wpfPaginate A(ctpPagination A)
	{
		return (wpfPaginate)A.ElementHost1.Child;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	private static void A()
	{
		KG.A.InvalidateControl(RIBBON_CONTROL);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Dictionary<int, CustomTaskPane> TaskPanes = Pane.TaskPanes;
		bool num = clsPanes.PresentationCloseFinal(ref TaskPanes, A);
		Pane.TaskPanes = TaskPanes;
		if (!num)
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
			Pane.A();
			return;
		}
	}
}
