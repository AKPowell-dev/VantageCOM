using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.UI;

public sealed class Pane
{
	private static readonly string m_A = AH.A(57272);

	private static readonly string m_B = AH.A(57293);

	private static readonly string m_C = AH.A(57320);

	[CompilerGenerated]
	private static wpfPane m_A;

	[CompilerGenerated]
	private static BaseError m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	internal static wpfPane TaskPane
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

	internal static BaseError ActiveItem
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

	internal static clsDisplay Dpi
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

	internal static bool RefreshOnShow
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
	} = false;

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

	internal static void A(bool A, bool B)
	{
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		if (A && !Access.AllowPowerPointOperation((PlanType)5, (Restriction)2, false))
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					D();
					return;
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		if (A)
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
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null)
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
			if (TaskPanes.TryGetValue(hWND, out value))
			{
				if (A)
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
					if (RefreshOnShow)
					{
						wpfPane wpfPane2 = (wpfPane)((ctpDeckCheck)value.Control).ElementHost1.Child;
						wpfPane2.WholePresentation = B;
						wpfPane2.ShowLoadingState();
						if (value.Visible)
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
							wpfPane2.StartCheckAsynch();
							RefreshOnShow = false;
						}
						wpfPane2 = null;
					}
				}
				value.Visible = A;
			}
		}
		ctpDeckCheck control;
		if (A && value == null)
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
			RefreshOnShow = true;
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)12, AH.A(56620));
			control = new ctpDeckCheck(B);
			value = NG.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
			value.Width = checked((int)Math.Round(380.0 * Dpi.X));
			value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
			value.VisibleChanged += Pane.A;
			value.Visible = true;
			if (TaskPanes == null)
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
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hWND, value);
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(Pane.B));
		if (A)
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
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(Pane.B));
		}
		activeWindow = null;
		application = null;
		value = null;
		control = null;
	}

	internal static void B()
	{
		if (clsRibbon.CallbackSlideView(ShowWarning: true))
		{
			A(NG.A.Application.ActivePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)), B: true);
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)12, AH.A(56620));
		}
	}

	internal static void C()
	{
		if (!clsRibbon.CallbackSlideView(ShowWarning: true))
		{
			return;
		}
		Selection selection = NG.A.Application.ActiveWindow.Selection;
		PpSelectionType type = selection.Type;
		if ((uint)(type - 1) <= 2u)
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
			A(selection.SlideRange, B: false);
		}
		else
		{
			try
			{
				A(selection.SlideRange, B: false);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.WarningMessage(AH.A(13552));
				ProjectData.ClearProjectError();
			}
		}
		selection = null;
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)12, AH.A(56657));
	}

	private static void A(SlideRange A, bool B)
	{
		RefreshOnShow = true;
		Pane.A(A: true, B);
		D();
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane obj = (CustomTaskPane)A;
		ctpDeckCheck ctpDeckCheck = (ctpDeckCheck)obj.Control;
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (obj.Visible)
		{
			if (RefreshOnShow)
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
				((wpfPane)ctpDeckCheck.ElementHost1.Child).StartCheckAsynch();
				RefreshOnShow = false;
			}
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).AddEventHandler(application, new EApplication_PresentationBeforeCloseEventHandler(Pane.A));
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(Pane.A));
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).AddEventHandler(application, new EApplication_SlideSelectionChangedEventHandler(Pane.A));
		}
		else if (!KG.A)
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
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).RemoveEventHandler(application, new EApplication_PresentationBeforeCloseEventHandler(Pane.A));
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(Pane.A));
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56735)).RemoveEventHandler(application, new EApplication_SlideSelectionChangedEventHandler(Pane.A));
			CloseCallout();
			D();
		}
		ctpDeckCheck = null;
		application = null;
	}

	internal static bool A()
	{
		bool result;
		try
		{
			if (TaskPanes != null)
			{
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
					CustomTaskPane value = null;
					if (TaskPanes.TryGetValue(NG.A.Application.ActiveWindow.HWND, out value))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							result = value.Visible;
							break;
						}
					}
					else
					{
						result = false;
					}
					break;
				}
			}
			else
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void D()
	{
		KG.A.InvalidateControl(Pane.m_B);
		KG.A.InvalidateControl(Pane.m_C);
	}

	internal static void A(SlideRange A)
	{
		CustomTaskPane value = null;
		if (A.Count != 1)
		{
			return;
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
			if (!TaskPanes.TryGetValue(A.Application.ActiveWindow.HWND, out value))
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
				((wpfPane)((ctpDeckCheck)value.Control).ElementHost1.Child).ScrollSpy(A[1]);
				value = null;
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		CloseCallout();
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		E();
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Presentation A)
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			D();
			return;
		}
	}

	private static void E()
	{
		checked
		{
			for (int i = TaskPanes.Count - 1; i >= 0; i += -1)
			{
				CustomTaskPane value = TaskPanes.ElementAt(i).Value;
				try
				{
					_ = value.Title;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					NG.A.CustomTaskPanes.Remove(value);
					TaskPanes.Remove(TaskPanes.ElementAt(i).Key);
					ProjectData.ClearProjectError();
				}
				value = null;
			}
		}
	}

	public static void CloseCallout()
	{
		try
		{
			IEnumerable<wpfCallout> source = System.Windows.Application.Current.Windows.OfType<wpfCallout>();
			if (source.Any())
			{
				source.ElementAt(0).Close();
			}
			source = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal static void F()
	{
		A(TaskPane.WarningsView.lbxResults);
	}

	internal static void A(ListBox A)
	{
		Pane.FocusListBoxItem(A);
	}

	internal static void A(int A)
	{
		if (A < 1)
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
			string text;
			if (A <= 1)
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
				text = AH.A(56778);
			}
			else
			{
				text = string.Format(AH.A(57022), A);
			}
			Forms.WarningMessage(text);
			return;
		}
	}
}
