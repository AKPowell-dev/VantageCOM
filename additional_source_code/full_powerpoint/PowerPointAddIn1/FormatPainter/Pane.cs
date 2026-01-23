using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.FormatPainter;

public sealed class Pane
{
	private static readonly string m_A = AH.A(138778);

	private static readonly string B = AH.A(138807);

	[CompilerGenerated]
	private static Properties m_A;

	[CompilerGenerated]
	private static Properties B;

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

	public static Properties CopiedProperties
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

	public static Properties CurrentProperties
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	} = null;

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

	public static void Toggle(bool blnPressed)
	{
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Expected O, but got Unknown
		Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		bool flag = false;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null && TaskPanes.TryGetValue(hWND, out value))
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
			value.Visible = blnPressed;
			flag = true;
		}
		if (blnPressed && !flag)
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
			ctpFormatPainter control = new ctpFormatPainter();
			value = NG.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
			clsPanes.A(value, Dpi);
			value.VisibleChanged += A;
			value.Visible = true;
			if (TaskPanes == null)
			{
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hWND, value);
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

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		ctpFormatPainter ctpFormatPainter = (ctpFormatPainter)customTaskPane.Control;
		Application application = NG.A.Application;
		ctpFormatPainter.A.Visible = customTaskPane.Visible;
		if (customTaskPane.Visible)
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
			try
			{
				Selection selection = application.ActiveWindow.Selection;
				ShapeRange shapeRange = Base.SelectedShapes(selection);
				FormatTree a = ctpFormatPainter.A;
				if (CopiedProperties == null)
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
					if (IsSingleShapeSelected(selection))
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
						CopiedProperties = new Properties(shapeRange[1]);
						a.PopulateProperties();
						a.btnCopy.IsEnabled = true;
					}
				}
				else
				{
					a.PopulateProperties();
					if (selection.Type != PpSelectionType.ppSelectionShapes)
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
						if (selection.Type != PpSelectionType.ppSelectionText)
						{
							a.btnApply.IsEnabled = false;
							a.btnCopy.IsEnabled = false;
							goto IL_0142;
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
					}
					a.btnApply.IsEnabled = true;
					a.btnCopy.IsEnabled = shapeRange.Count == 1;
				}
				goto IL_0142;
				IL_0142:
				a = null;
				selection = null;
				shapeRange = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		else if (!KG.A)
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
			PB.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			Pane.A();
		}
		customTaskPane = null;
		ctpFormatPainter = null;
		application = null;
	}

	public static bool IsSingleShapeSelected(Selection sel)
	{
		if (sel.Type == PpSelectionType.ppSelectionText)
		{
			return true;
		}
		if (sel.Type == PpSelectionType.ppSelectionShapes)
		{
			return Base.SelectedShapes(sel).Count == 1;
		}
		bool result = default(bool);
		return result;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	public static CustomTaskPane GetPaneByHwnd(int hwnd)
	{
		CustomTaskPane value = null;
		try
		{
			if (TaskPanes != null)
			{
				TaskPanes.TryGetValue(hwnd, out value);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return value;
	}

	private static void A()
	{
		KG.A.InvalidateControl(Pane.B);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Dictionary<int, CustomTaskPane> TaskPanes = Pane.TaskPanes;
		bool num = clsPanes.PresentationCloseFinal(ref TaskPanes, A);
		Pane.TaskPanes = TaskPanes;
		if (num)
		{
			Pane.A();
		}
	}
}
