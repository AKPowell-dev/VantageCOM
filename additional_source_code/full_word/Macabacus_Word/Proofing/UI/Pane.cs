using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.UI;

public sealed class Pane
{
	private static readonly string m_A = XC.A(38567);

	private static readonly string m_B = XC.A(38584);

	private static wpfPane m_A;

	private static BaseError m_A;

	private static Selection m_A;

	private static bool m_A;

	[CompilerGenerated]
	private static bool m_B;

	private static Dictionary<int, CustomTaskPane> m_A = null;

	public static wpfPane TaskPane
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static BaseError ActiveItem
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static Selection Selection
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static bool WholeDocument
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	private static bool RefreshOnShow
	{
		[CompilerGenerated]
		get
		{
			return Pane.m_B;
		}
		[CompilerGenerated]
		set
		{
			Pane.m_B = value;
		}
	} = false;

	private static Dictionary<int, CustomTaskPane> A
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static void Toggle(bool blnShow)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		Microsoft.Office.Interop.Word.Window activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hwnd = activeWindow.Hwnd;
		if (Pane.A != null && Pane.A.TryGetValue(hwnd, out value))
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
			if (blnShow && RefreshOnShow)
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
				wpfPane wpfPane2 = (wpfPane)((ctpProofing)value.Control).ElementHost1.Child;
				wpfPane2.ShowLoadingState();
				if (value.Visible)
				{
					wpfPane2.StartCheckAsynch();
					RefreshOnShow = false;
				}
				wpfPane2 = null;
			}
			value.Visible = blnShow;
		}
		ctpProofing control;
		if (blnShow)
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
			if (value == null)
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
				if (Selection == null)
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
					Selection = application.Selection;
					WholeDocument = true;
					RefreshOnShow = true;
					clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)12, XC.A(38472));
				}
				control = new ctpProofing();
				value = PC.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
				value.Width = 400;
				value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				value.VisibleChanged += A;
				value.Visible = true;
				if (Pane.A == null)
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
					Pane.A = new Dictionary<int, CustomTaskPane>();
				}
				Pane.A.Add(hwnd, value);
			}
		}
		activeWindow = null;
		application = null;
		value = null;
		control = null;
	}

	public static void CheckDocument()
	{
		A(PC.A.Application.Selection, B: true);
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)12, XC.A(38472));
	}

	public static void CheckSelection()
	{
		Selection selection = PC.A.Application.ActiveWindow.Selection;
		if (selection.Type != WdSelectionType.wdNoSelection)
		{
			A(selection, B: false);
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)12, XC.A(38501));
		}
		else
		{
			System.Windows.Forms.MessageBox.Show(XC.A(38532), XC.A(2438), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		selection = null;
	}

	private static bool A(Selection A)
	{
		return Conversions.ToBoolean(A.get_Information(WdInformation.wdWithInTable));
	}

	private static void A(Selection A, bool B)
	{
		Selection = A;
		WholeDocument = B;
		RefreshOnShow = true;
		Toggle(blnShow: true);
		Pane.A();
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane obj = (CustomTaskPane)A;
		ctpProofing ctpProofing = (ctpProofing)obj.Control;
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		if (obj.Visible)
		{
			if (RefreshOnShow)
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
				((wpfPane)ctpProofing.ElementHost1.Child).StartCheckAsynch();
				RefreshOnShow = false;
			}
			new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).AddEventHandler(application, new ApplicationEvents4_DocumentBeforeCloseEventHandler(Pane.A));
		}
		else if (!NC.B)
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
			new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).RemoveEventHandler(application, new ApplicationEvents4_DocumentBeforeCloseEventHandler(Pane.A));
			CloseCallout();
			Pane.A();
		}
		ctpProofing = null;
		application = null;
	}

	public static bool IsVisible()
	{
		bool result;
		try
		{
			if (Pane.A != null)
			{
				CustomTaskPane value = null;
				if (Pane.A.TryGetValue(PC.A.Application.ActiveWindow.Hwnd, out value))
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
						result = value.Visible;
						break;
					}
				}
				else
				{
					result = false;
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

	private static void A()
	{
		NC.A.InvalidateControl(Pane.m_B);
	}

	private static void A(Document A, ref bool B)
	{
		CloseCallout();
	}

	private static void A(Document A)
	{
		B();
	}

	private static void B(Document A)
	{
		Pane.A();
	}

	private static void B()
	{
		checked
		{
			for (int i = Pane.A.Count - 1; i >= 0; i += -1)
			{
				CustomTaskPane value = Pane.A.ElementAt(i).Value;
				try
				{
					_ = value.Title;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					PC.A.CustomTaskPanes.Remove(value);
					Pane.A.Remove(Pane.A.ElementAt(i).Key);
					ProjectData.ClearProjectError();
				}
				value = null;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return;
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

	public static void RefocusActiveListBoxItem()
	{
		FocusListBoxItem(TaskPane.lbxResults);
	}

	public static void FocusListBoxItem(System.Windows.Controls.ListBox lbx)
	{
		Pane.FocusListBoxItem(lbx);
	}
}
