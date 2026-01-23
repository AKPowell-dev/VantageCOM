using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Library2.UI;

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
	internal sealed class MD
	{
		public wpfLibrary A;

		[SpecialName]
		internal bool A(CustomTaskPane A)
		{
			return object.Equals(((ctpLibrary2)A.Control).A, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class ND
	{
		public CustomTaskPane A;

		[SpecialName]
		internal bool A(KeyValuePair<int, CustomTaskPane> A)
		{
			return object.Equals(A.Value, this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class OD
	{
		public bool A;

		public Func<wpfLibrary, bool> A;

		public OD(OD A)
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
				return A.AllGroups == null;
			}
			return true;
		}
	}

	[CompilerGenerated]
	internal sealed class PD
	{
		public wpfLibrary A;

		public PD(PD A)
		{
			if (A == null)
			{
				return;
			}
			while (true)
			{
				switch (6)
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

	private static readonly string m_A = AH.A(64918);

	private static readonly string m_B = AH.A(64933);

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static bool m_A;

	public static bool PDFsContentIsEnabled = false;

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
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Expected O, but got Unknown
		Application application = NG.A.Application;
		DocumentWindow activeWindow = application.ActiveWindow;
		CustomTaskPane value = null;
		int hWND = activeWindow.HWND;
		if (blnShow)
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
			Dpi = new clsDisplay();
		}
		if (TaskPanes != null)
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
			if (TaskPanes.TryGetValue(hWND, out value))
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
				if (blnShow)
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
								switch (4)
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
		if (blnShow)
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
				RefreshOnShow = true;
				control = new ctpLibrary2();
				value = NG.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
				clsPanes.A(value, Dpi, PB.Settings.ContentInsertPaneWidth);
				value.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				value.VisibleChanged += A;
				value.Visible = true;
				if (TaskPanes == null)
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
					TaskPanes = new Dictionary<int, CustomTaskPane>();
				}
				TaskPanes.Add(hWND, value);
			}
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		if (blnShow)
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
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(application, new EApplication_PresentationCloseFinalEventHandler(A));
		}
		activeWindow = null;
		application = null;
		value = null;
		control = null;
	}

	private static IEnumerable<CustomTaskPane> A()
	{
		return NG.A.CustomTaskPanes.Where([SpecialName] (CustomTaskPane A) => A.Control is ctpLibrary2);
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
				switch (7)
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
			NG.A.CustomTaskPanes.Remove(A2);
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
				if (keyValuePair.Value != null)
				{
					TaskPanes.Remove(keyValuePair.Key);
				}
				return;
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
		OD a = default(OD);
		OD CS_0024_003C_003E8__locals2 = new OD(a);
		CS_0024_003C_003E8__locals2.A = B;
		PD pD = default(PD);
		while (true)
		{
			pD = new PD(pD);
			pD.A = A.FirstOrDefault([SpecialName] (wpfLibrary wpfLibrary2) => !CS_0024_003C_003E8__locals2.A || wpfLibrary2.AllGroups == null);
			if (pD.A != null)
			{
				pD.A.LoadContent();
				A.Remove(pD.A);
				A.RemoveAll(pD.A);
				pD.A = null;
				continue;
			}
			break;
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
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				wpfLibrary2.HiddenOnBackstaging = false;
				goto IL_0108;
			}
		}
		if (customTaskPane.Visible)
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
			if (!RefreshOnShow)
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
				if (!Load.RefreshRequired())
				{
					goto IL_00b4;
				}
			}
			List<wpfLibrary> A2 = Pane.A().ToList();
			Pane.A(ref A2, B: true);
			Pane.A(ref A2, B: false);
			A2 = null;
			RefreshOnShow = false;
			goto IL_00b4;
		}
		if (!KG.A)
		{
			Pane.A();
			PB.Settings.ContentInsertPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
			wpfLibrary2.CleanUp();
		}
		else
		{
			wpfLibrary2.HiddenOnBackstaging = true;
		}
		goto IL_0108;
		IL_00b4:
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(64754));
		goto IL_0108;
		IL_0108:
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
					if (TaskPanes.TryGetValue(NG.A.Application.ActiveWindow.HWND, out value))
					{
						while (true)
						{
							switch (4)
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
		KG.A.InvalidateControl(Pane.m_B);
	}

	public static void ShowFromRibbon(bool blnSlides, bool blnShapes, bool blnImages, bool blnCharts, bool blnText, bool blnDecks, bool blnVideos, bool blnPDFs = false)
	{
		try
		{
			Show(blnSlides, blnShapes, blnImages, blnCharts, blnText, blnDecks, blnVideos, blnPDFs);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			Forms.ErrorMessage(AH.A(64793) + ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	public static void Show(bool blnSlides, bool blnShapes, bool blnImages, bool blnCharts, bool blnText, bool blnDecks, bool blnVideos, bool blnPDFs = false)
	{
		MySettings settings = PB.Settings;
		settings.ContentInsertShowSlides = blnSlides;
		settings.ContentInsertShowShapes = blnShapes;
		settings.ContentInsertShowImages = blnImages;
		settings.ContentInsertShowVideos = blnVideos;
		settings.ContentInsertShowCharts = blnCharts;
		settings.ContentInsertShowText = blnText;
		if (PDFsContentIsEnabled)
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
			settings.ContentInsertShowPDFs = blnPDFs;
		}
		settings.ContentInsertShowDecks = blnDecks;
		settings = null;
		RefreshOnShow = true;
		Toggle(blnShow: true);
		A();
	}

	internal static void B()
	{
		if (TaskPanes == null)
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
			CustomTaskPane value = null;
			try
			{
				DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
				if (!TaskPanes.TryGetValue(activeWindow.HWND, out value))
				{
					while (true)
					{
						switch (7)
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
				DocumentWindow activeWindow = null;
			}
		}
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
			switch (6)
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
