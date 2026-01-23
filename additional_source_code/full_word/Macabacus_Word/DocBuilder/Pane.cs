using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.DocBuilder;

public sealed class Pane
{
	private static readonly string m_A = XC.A(21825);

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

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

	public static void Toggle(bool blnShow)
	{
		if (blnShow)
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
					if (Licensing.AllowDocBuilderOperation())
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								try
								{
									Base.InspectDocument(PC.A.Application.ActiveDocument, blnManual: true);
									return;
								}
								catch (InvalidTemplateException ex)
								{
									ProjectData.SetProjectError(ex);
									InvalidTemplateException a = ex;
									A(a);
									B();
									ProjectData.ClearProjectError();
									return;
								}
								catch (Exception ex2)
								{
									ProjectData.SetProjectError(ex2);
									Exception ex3 = ex2;
									A(ex3);
									B();
									clsReporting.LogException(ex3);
									ProjectData.ClearProjectError();
									return;
								}
							}
						}
					}
					B();
					return;
				}
			}
		}
		A();
	}

	private static void A(Exception A)
	{
		Forms.ErrorMessage(A.Message);
	}

	public static void Show(List<BaseQuestion> listQuestions)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		Window activeWindow = PC.A.Application.ActiveWindow;
		CustomTaskPane value = null;
		Dpi = new clsDisplay();
		int hwnd = activeWindow.Hwnd;
		if (TaskPanes != null && TaskPanes.TryGetValue(hwnd, out value))
		{
			_ = (wpfDocWizard)((ctpDocWizard)value.Control).ElementHost1.Child;
			value.Visible = true;
		}
		ctpDocWizard control;
		if (value == null)
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
			control = new ctpDocWizard();
			value = PC.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
			clsPanes.A(value, Dpi);
			value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
			value.VisibleChanged += A;
			value.Visible = true;
			if (TaskPanes == null)
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
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hwnd, value);
		}
		((wpfDocWizard)((ctpDocWizard)value.Control).ElementHost1.Child).LoadQuestionnaire(listQuestions);
		activeWindow = null;
		value = null;
		control = null;
	}

	private static void A()
	{
		CustomTaskPane value = null;
		int hwnd = PC.A.Application.ActiveWindow.Hwnd;
		if (TaskPanes != null && TaskPanes.TryGetValue(hwnd, out value))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			value.Visible = false;
		}
		value = null;
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		_ = (ctpDocWizard)customTaskPane.Control;
		if (customTaskPane.Visible)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(21825));
		}
		else if (!NC.B)
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
			Pane.B();
			N.Settings.TaskPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
		}
		customTaskPane = null;
	}

	public static bool IsVisible()
	{
		return clsPanes.IsVisible(TaskPanes);
	}

	internal static void B()
	{
		NC.A.InvalidateControl(XC.A(21848));
	}
}
