using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2.UI;

public sealed class Pane
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<CustomTaskPane, bool> A;

		public static Func<CustomTaskPane, wpfLibrary> A;

		public static Func<wpfLibrary, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(CustomTaskPane A)
		{
			return A.Control is ctpLibrary2;
		}

		[SpecialName]
		internal wpfLibrary A(CustomTaskPane A)
		{
			return ((ctpLibrary2)A.Control).A;
		}

		[SpecialName]
		internal bool A(wpfLibrary A)
		{
			return A != null;
		}
	}

	[CompilerGenerated]
	internal sealed class XE
	{
		public wpfLibrary A;

		[SpecialName]
		internal bool A(CustomTaskPane A)
		{
			return object.Equals(((ctpLibrary2)A.Control).A, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class YE
	{
		public CustomTaskPane A;

		[SpecialName]
		internal bool A(KeyValuePair<int, CustomTaskPane> A)
		{
			return object.Equals(A.Value, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class ZE
	{
		public bool A;

		public Predicate<wpfLibrary> A;

		public ZE(ZE A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(wpfLibrary A)
		{
			if (this.A)
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
						return A.AllGroups == null;
					}
				}
			}
			return true;
		}
	}

	[CompilerGenerated]
	internal sealed class AF
	{
		public wpfLibrary A;

		public AF(AF A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(wpfLibrary A)
		{
			return object.Equals(A.AllGroups, this.A.AllGroups);
		}
	}

	private static readonly string m_A = VH.A(87667);

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static bool m_A;

	public static bool PDFsContentIsEnabled = false;

	public static bool ModelsContentIsEnabled = false;

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

	public static clsDisplay Dpi
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

	private static bool RefreshOnShow
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

	public static void Toggle(bool blnShow)
	{
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003f: Expected O, but got Unknown
		Window activeWindow = MH.A.Application.ActiveWindow;
		CustomTaskPane value = null;
		int hwnd = activeWindow.Hwnd;
		if (blnShow)
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
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null && TaskPanes.TryGetValue(hwnd, out value))
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
			if (blnShow)
			{
				wpfLibrary wpfLibrary2 = (wpfLibrary)((ctpLibrary2)value.Control).ElementHost1.Child;
				wpfLibrary2.G();
				if (RefreshOnShow)
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
					if (value.Visible)
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
						wpfLibrary2.A(A: true);
						RefreshOnShow = false;
					}
				}
				wpfLibrary2 = null;
			}
			value.Visible = blnShow;
		}
		ctpLibrary2 control;
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
				RefreshOnShow = true;
				control = new ctpLibrary2();
				value = MH.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
				clsPanes.A(value, Dpi, K.Settings.ContentInsertPaneWidth);
				value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				value.VisibleChanged += A;
				value.Visible = true;
				if (TaskPanes == null)
				{
					TaskPanes = new Dictionary<int, CustomTaskPane>();
				}
				TaskPanes.Add(hwnd, value);
			}
		}
		activeWindow = null;
		value = null;
		control = null;
	}

	private static IEnumerable<CustomTaskPane> A()
	{
		CustomTaskPaneCollection customTaskPanes = MH.A.CustomTaskPanes;
		Func<CustomTaskPane, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (CustomTaskPane A) => A.Control is ctpLibrary2);
		}
		else
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
			predicate = _Closure_0024__.A;
		}
		return customTaskPanes.Where(predicate);
	}

	private static IEnumerable<wpfLibrary> A()
	{
		IEnumerable<CustomTaskPane> source = A();
		Func<CustomTaskPane, wpfLibrary> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (CustomTaskPane A) => ((ctpLibrary2)A.Control).A);
		}
		else
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
			selector = _Closure_0024__.A;
		}
		IEnumerable<wpfLibrary> source2 = source.Select(selector);
		Func<wpfLibrary, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (wpfLibrary A) => A != null);
		}
		else
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
			predicate = _Closure_0024__.A;
		}
		return source2.Where(predicate);
	}

	internal static void A(wpfLibrary A)
	{
		try
		{
			CustomTaskPane A2 = Pane.A().FirstOrDefault([SpecialName] (CustomTaskPane customTaskPane) => object.Equals(((ctpLibrary2)customTaskPane.Control).A, A));
			if (A2 == null)
			{
				return;
			}
			MH.A.CustomTaskPanes.Remove(A2);
			if (TaskPanes == null)
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
				KeyValuePair<int, CustomTaskPane> keyValuePair = TaskPanes.FirstOrDefault([SpecialName] (KeyValuePair<int, CustomTaskPane> keyValuePair2) => object.Equals(keyValuePair2.Value, A2));
				if (keyValuePair.Value == null)
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
					TaskPanes.Remove(keyValuePair.Key);
					return;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(ref List<wpfLibrary> A, bool B)
	{
		ZE a = default(ZE);
		ZE CS_0024_003C_003E8__locals5 = new ZE(a);
		CS_0024_003C_003E8__locals5.A = B;
		AF aF = default(AF);
		while (true)
		{
			aF = new AF(aF);
			AF aF2 = aF;
			List<wpfLibrary> obj = A;
			Predicate<wpfLibrary> match;
			if (CS_0024_003C_003E8__locals5.A != null)
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
				match = CS_0024_003C_003E8__locals5.A;
			}
			else
			{
				match = (CS_0024_003C_003E8__locals5.A = [SpecialName] (wpfLibrary wpfLibrary2) =>
				{
					if (CS_0024_003C_003E8__locals5.A)
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
								return wpfLibrary2.AllGroups == null;
							}
						}
					}
					return true;
				});
			}
			aF2.A = obj.Find(match);
			if (aF.A == null)
			{
				break;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0064;
				}
				continue;
				end_IL_0064:
				break;
			}
			aF.A.LoadContent();
			A.Remove(aF.A);
			A.RemoveAll(aF.A);
			aF.A = null;
		}
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		wpfLibrary wpfLibrary2 = (wpfLibrary)((ctpLibrary2)customTaskPane.Control).ElementHost1.Child;
		if (customTaskPane.Visible)
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
			if (wpfLibrary2.HiddenOnBackstaging)
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
				wpfLibrary2.HiddenOnBackstaging = false;
				goto IL_010e;
			}
		}
		if (customTaskPane.Visible)
		{
			if (!RefreshOnShow)
			{
				if (!Load.RefreshRequired())
				{
					goto IL_00ac;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			List<wpfLibrary> A2 = Pane.A().ToList();
			Pane.A(ref A2, B: true);
			Pane.A(ref A2, B: false);
			A2 = null;
			RefreshOnShow = false;
			goto IL_00ac;
		}
		if (!KH.A)
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
			Pane.A();
			K.Settings.ContentInsertPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			wpfLibrary2.CleanUp();
		}
		else
		{
			wpfLibrary2.HiddenOnBackstaging = true;
		}
		goto IL_010e;
		IL_00ac:
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)6, VH.A(87476));
		goto IL_010e;
		IL_010e:
		customTaskPane = null;
		wpfLibrary2 = null;
	}

	public static bool IsVisible()
	{
		bool result;
		try
		{
			if (TaskPanes != null)
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
					CustomTaskPane value = null;
					if (TaskPanes.TryGetValue(MH.A.Application.ActiveWindow.Hwnd, out value))
					{
						while (true)
						{
							switch (6)
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

	private static void A()
	{
		KH.A.InvalidateControl(VH.A(87515));
	}

	public static void Show(bool blnShapes, bool blnImages, bool blnCharts, bool blnText, bool blnTables, bool blnModels = false, bool blnPDFs = false)
	{
		try
		{
			MySettings settings = K.Settings;
			settings.ContentInsertShowShapes = blnShapes;
			settings.ContentInsertShowImages = blnImages;
			settings.ContentInsertShowCharts = blnCharts;
			settings.ContentInsertShowText = blnText;
			if (PDFsContentIsEnabled)
			{
				settings.ContentInsertShowPDFs = blnPDFs;
			}
			if (ModelsContentIsEnabled)
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
				settings.ContentInsertShowModels = blnModels;
			}
			settings.ContentInsertShowTables = blnTables;
			settings = null;
			RefreshOnShow = true;
			Toggle(blnShow: true);
			A();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			Forms.ErrorMessage(VH.A(87542) + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	internal static void B()
	{
		if (TaskPanes == null)
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
			CustomTaskPane value = null;
			try
			{
				Window activeWindow = MH.A.Application.ActiveWindow;
				if (!TaskPanes.TryGetValue(activeWindow.Hwnd, out value))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				if (!value.Visible)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				Toggle(blnShow: false);
				Toggle(blnShow: true);
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
				return;
			}
			finally
			{
				value = null;
				Window activeWindow = null;
			}
		}
	}

	public static void SoftDisable()
	{
		if (TaskPanes == null)
		{
			return;
		}
		checked
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
				try
				{
					Application application = MH.A.Application;
					if (application.Windows.Count > 0)
					{
						TaskPanes[application.ActiveWindow.Hwnd].Visible = false;
					}
					application = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					Dictionary<int, CustomTaskPane> taskPanes = TaskPanes;
					for (int i = taskPanes.Count - 1; i >= 0; i += -1)
					{
						try
						{
							int key = taskPanes.ElementAt(i).Key;
							MH.A.CustomTaskPanes.Remove(taskPanes[key]);
							taskPanes.Remove(key);
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
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
						taskPanes = null;
						break;
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				TaskPanes = null;
				return;
			}
		}
	}
}
