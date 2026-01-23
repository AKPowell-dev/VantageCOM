using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.UI;

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
	internal sealed class S
	{
		public wpfLibrary A;

		[SpecialName]
		internal bool A(CustomTaskPane A)
		{
			return object.Equals(((ctpLibrary2)A.Control).A, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class T
	{
		public CustomTaskPane A;

		[SpecialName]
		internal bool A(KeyValuePair<int, CustomTaskPane> A)
		{
			return object.Equals(A.Value, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class U
	{
		public bool A;

		public Predicate<wpfLibrary> A;

		public U(U A)
		{
			if (A == null)
			{
				return;
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
					switch (7)
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
	internal sealed class V
	{
		public wpfLibrary A;

		public V(V A)
		{
			if (A == null)
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

	private static readonly string m_A = XC.A(9826);

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static bool m_A;

	public static bool PDFsContentIsEnabled = false;

	public static bool DocsContentIsEnabled = false;

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
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		if (PC.A.Application.Documents.Count == 0)
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
					Forms.WarningMessage(XC.A(7399));
					A();
					return;
				}
			}
		}
		Window activeWindow = PC.A.Application.ActiveWindow;
		CustomTaskPane value = null;
		int hwnd = activeWindow.Hwnd;
		if (blnShow)
		{
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null)
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
			if (TaskPanes.TryGetValue(hwnd, out value))
			{
				if (blnShow)
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
					wpfLibrary wpfLibrary2 = (wpfLibrary)((ctpLibrary2)value.Control).ElementHost1.Child;
					wpfLibrary2.G();
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
						if (value.Visible)
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
							wpfLibrary2.A(A: true);
							RefreshOnShow = false;
						}
					}
					wpfLibrary2 = null;
				}
				value.Visible = blnShow;
			}
		}
		ctpLibrary2 control;
		if (blnShow && value == null)
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
			RefreshOnShow = true;
			control = new ctpLibrary2();
			value = PC.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
			clsPanes.A(value, Dpi, N.Settings.ContentInsertPaneWidth);
			value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
			value.VisibleChanged += A;
			value.Visible = true;
			if (TaskPanes == null)
			{
				TaskPanes = new Dictionary<int, CustomTaskPane>();
			}
			TaskPanes.Add(hwnd, value);
		}
		activeWindow = null;
		value = null;
		control = null;
	}

	private static IEnumerable<CustomTaskPane> A()
	{
		CustomTaskPaneCollection customTaskPanes = PC.A.CustomTaskPanes;
		Func<CustomTaskPane, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (CustomTaskPane A) => A.Control is ctpLibrary2);
		}
		else
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
				switch (3)
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
				PC.A.CustomTaskPanes.Remove(A2);
				if (TaskPanes == null)
				{
					return;
				}
				KeyValuePair<int, CustomTaskPane> keyValuePair = TaskPanes.FirstOrDefault([SpecialName] (KeyValuePair<int, CustomTaskPane> keyValuePair2) => object.Equals(keyValuePair2.Value, A2));
				if (keyValuePair.Value == null)
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
		U a = default(U);
		U CS_0024_003C_003E8__locals2 = new U(a);
		CS_0024_003C_003E8__locals2.A = B;
		V v = default(V);
		while (true)
		{
			v = new V(v);
			v.A = A.Find([SpecialName] (wpfLibrary wpfLibrary2) =>
			{
				if (CS_0024_003C_003E8__locals2.A)
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
							return wpfLibrary2.AllGroups == null;
						}
					}
				}
				return true;
			});
			if (v.A == null)
			{
				break;
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
				v.A.LoadContent();
				A.Remove(v.A);
				A.RemoveAll(v.A);
				v.A = null;
				break;
			}
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
			if (wpfLibrary2.HiddenOnBackstaging)
			{
				wpfLibrary2.HiddenOnBackstaging = false;
				goto IL_011a;
			}
		}
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
			if (!RefreshOnShow)
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
				if (!Load.RefreshRequired())
				{
					goto IL_00b6;
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
			List<wpfLibrary> A2 = Pane.A().ToList();
			Pane.A(ref A2, B: true);
			Pane.A(ref A2, B: false);
			A2 = null;
			RefreshOnShow = false;
			goto IL_00b6;
		}
		if (!NC.B)
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
			Pane.A();
			N.Settings.ContentInsertPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			wpfLibrary2.CleanUp();
		}
		else
		{
			wpfLibrary2.HiddenOnBackstaging = true;
		}
		goto IL_011a;
		IL_011a:
		customTaskPane = null;
		wpfLibrary2 = null;
		return;
		IL_00b6:
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(9639));
		goto IL_011a;
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
					if (TaskPanes.TryGetValue(PC.A.Application.ActiveWindow.Hwnd, out value))
					{
						while (true)
						{
							switch (1)
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
		NC.A.InvalidateControl(XC.A(9678));
	}

	public static void Show(bool blnShapes, bool blnImages, bool blnCharts, bool blnText, bool blnDocs = false, bool blnPDFs = false)
	{
		try
		{
			MySettings settings = N.Settings;
			settings.ContentInsertShowShapes = blnShapes;
			settings.ContentInsertShowImages = blnImages;
			settings.ContentInsertShowCharts = blnCharts;
			settings.ContentInsertShowText = blnText;
			if (PDFsContentIsEnabled)
			{
				settings.ContentInsertShowPDFs = blnPDFs;
			}
			if (DocsContentIsEnabled)
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
				settings.ContentInsertShowDocs = blnDocs;
			}
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
			Forms.ErrorMessage(XC.A(9701) + ex2.Message);
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
			try
			{
				Window activeWindow = PC.A.Application.ActiveWindow;
				if (!TaskPanes.TryGetValue(activeWindow.Hwnd, out value))
				{
					while (true)
					{
						switch (3)
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
						switch (4)
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
}
