using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Explorer;

public sealed class Pane
{
	public static readonly string PANE_TITLE = AH.A(116233);

	private static readonly string m_A = AH.A(116250);

	public static bool IsPresentationClosing = false;

	public static string ClosedPresentationName = "";

	private static ObservableCollection<PresentationItem> m_A = null;

	[CompilerGenerated]
	private static clsDisplay m_A;

	[CompilerGenerated]
	private static Dictionary<int, CustomTaskPane> m_A;

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static bool m_B;

	[CompilerGenerated]
	private static Cache m_A;

	public static ObservableCollection<PresentationItem> AllPresentations
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
			new ViewModel().NotifyPropertyChanged(AH.A(116200));
			_ = null;
		}
	}

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

	public static Dictionary<int, CustomTaskPane> PanesCollection
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

	public static bool IsOpen
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

	public static bool Initialized
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

	public static Cache CachedObjects
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
		if (blnPressed)
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
			if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
			{
				C();
				return;
			}
		}
		Application application = NG.A.Application;
		IsOpen = blnPressed;
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
			CachedObjects = new Cache();
			if (PB.Settings.ExplorerShowAllPresentations)
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
				LoadPresentations(application);
			}
			if (PanesCollection == null)
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
				PanesCollection = new Dictionary<int, CustomTaskPane>();
			}
		}
		checked
		{
			for (int i = PanesCollection.Count - 1; i >= 0; i += -1)
			{
				CustomTaskPane value = PanesCollection.ElementAt(i).Value;
				try
				{
					_ = value.Title;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					NG.A.CustomTaskPanes.Remove(value);
					PanesCollection.Remove(PanesCollection.ElementAt(i).Key);
					ProjectData.ClearProjectError();
					continue;
				}
				CustomTaskPane customTaskPane = value;
				if (!blnPressed)
				{
					customTaskPane.VisibleChanged -= B;
				}
				try
				{
					customTaskPane.Visible = blnPressed;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (!blnPressed)
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
					Events.Disable(application);
					if (!blnPressed)
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
						customTaskPane.VisibleChanged += B;
					}
				}
				customTaskPane = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
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
					A(application);
					B(application);
					clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)7, AH.A(116233));
				}
				else
				{
					C(application);
					if (!PanesCollection.Any())
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
						PanesCollection = null;
					}
				}
				application = null;
				CustomTaskPane value = null;
				return;
			}
		}
	}

	public static void LoadPresentations(Application ppApp)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation;
		try
		{
			presentation = ppApp.ActivePresentation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			presentation = null;
			ProjectData.ClearProjectError();
		}
		if (AllPresentations == null)
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
			AllPresentations = new ObservableCollection<PresentationItem>();
		}
		else
		{
			AllPresentations.Clear();
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ppApp.Presentations.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current;
				if (presentation2.Windows.Count <= 0)
				{
					continue;
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
				PresentationItem presentationItem = new PresentationItem(presentation2);
				if (presentation2 != presentation)
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
					presentationItem.Opacity = Constants.INACTIVE_FILE_OPACITY;
				}
				AllPresentations.Add(presentationItem);
				presentationItem = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00bf;
				}
				continue;
				end_IL_00bf:
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
		presentation = null;
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		IEnumerator<PresentationItem> enumerator = default(IEnumerator<PresentationItem>);
		try
		{
			enumerator = AllPresentations.GetEnumerator();
			while (enumerator.MoveNext())
			{
				PresentationItem current = enumerator.Current;
				if (current.Presentation == A)
				{
					current.Opacity = 1.0;
				}
				else
				{
					current.Opacity = Constants.INACTIVE_FILE_OPACITY;
				}
			}
		}
		finally
		{
			if (enumerator != null)
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	private static void A(Application A)
	{
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Presentations.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Pane.A((Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current, B: true);
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
					return;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B(Application A)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).AddEventHandler(A, new EApplication_AfterPresentationOpenEventHandler(CreatePane));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).AddEventHandler(A, new EApplication_AfterNewPresentationEventHandler(CreatePane));
	}

	public static void RemoveEvents(Application ppApp)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(63418)).RemoveEventHandler(ppApp, new EApplication_AfterPresentationOpenEventHandler(CreatePane));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105550)).RemoveEventHandler(ppApp, new EApplication_AfterNewPresentationEventHandler(CreatePane));
	}

	public static void CreatePane(Microsoft.Office.Interop.PowerPoint.Presentation Pres)
	{
		A(Pres, B: false);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, bool B)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Expected O, but got Unknown
		if (A.Windows.Count <= 0)
		{
			return;
		}
		if (Dpi == null)
		{
			Dpi = new clsDisplay();
		}
		if (!B)
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
			if (PB.Settings.ExplorerShowAllPresentations)
			{
				IEnumerator<PresentationItem> enumerator = default(IEnumerator<PresentationItem>);
				try
				{
					enumerator = AllPresentations.GetEnumerator();
					while (enumerator.MoveNext())
					{
						enumerator.Current.Opacity = Constants.INACTIVE_FILE_OPACITY;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0072;
						}
						continue;
						end_IL_0072:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				AllPresentations.Add(new PresentationItem(A));
			}
		}
		try
		{
			IEnumerator enumerator2 = A.Windows.GetEnumerator();
			try
			{
				while (enumerator2.MoveNext())
				{
					DocumentWindow documentWindow = (DocumentWindow)enumerator2.Current;
					if (PanesCollection.ContainsKey(documentWindow.HWND))
					{
						continue;
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
					ctpExplorer2 control = new ctpExplorer2();
					CustomTaskPane customTaskPane = NG.A.CustomTaskPanes.Add(control, PANE_TITLE, documentWindow);
					Pane.A((ctpExplorer2)customTaskPane.Control).Initialize(A);
					CustomTaskPane customTaskPane2 = customTaskPane;
					try
					{
						customTaskPane2.DockPosition = (MsoCTPDockPosition)PB.Settings.ExplorerPanePosn;
						try
						{
							customTaskPane2.Width = checked((int)Math.Round((double)PB.Settings.ExplorerPaneWidth * Dpi.X));
						}
						catch (ArgumentException ex)
						{
							ProjectData.SetProjectError(ex);
							ArgumentException ex2 = ex;
							customTaskPane2.Width = 350;
							ProjectData.ClearProjectError();
						}
						customTaskPane2.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					customTaskPane2.VisibleChanged += Pane.B;
					customTaskPane2.Visible = true;
					customTaskPane2 = null;
					PanesCollection.Add(documentWindow.HWND, customTaskPane);
					customTaskPane = null;
					control = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				IDisposable disposable = enumerator2 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			clsReporting.LogException(ex6);
			ProjectData.ClearProjectError();
		}
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
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
			Pane.A((ctpExplorer2)customTaskPane.Control).Connect();
		}
		else if (!KG.A)
		{
			if (IsPresentationClosing)
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
				Pane.B();
				NG.A.CustomTaskPanes.Remove(customTaskPane);
				IsPresentationClosing = false;
			}
			else
			{
				try
				{
					PB.Settings.ExplorerPaneWidth = checked((int)Math.Round((double)customTaskPane.Width / Dpi.X));
					PB.Settings.ExplorerPanePosn = (int)customTaskPane.DockPosition;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				Pane.A();
			}
		}
		customTaskPane = null;
	}

	private static void B(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		checked
		{
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
				Pane.A((ctpExplorer2)customTaskPane.Control).Connect();
			}
			else if (!KG.A)
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
				for (int i = PanesCollection.Count - 1; i >= 0; i += -1)
				{
					CustomTaskPane value = PanesCollection.ElementAt(i).Value;
					try
					{
						Operators.CompareString(value.Title, PANE_TITLE, TextCompare: false);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						value.VisibleChanged -= Pane.B;
						NG.A.CustomTaskPanes.Remove(value);
						PanesCollection.Remove(PanesCollection.ElementAt(i).Key);
						ProjectData.ClearProjectError();
					}
					value = null;
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
				bool flag = default(bool);
				if (Events.explorerDoNotClose)
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
					Events.explorerDoNotClose = false;
				}
				else if (Events.blnBeforeClose)
				{
					Events.blnBeforeClose = false;
				}
				else
				{
					flag = true;
				}
				if (flag)
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
					try
					{
						PB.Settings.ExplorerPaneWidth = (int)Math.Round((double)customTaskPane.Width / Dpi.X);
						PB.Settings.ExplorerPanePosn = unchecked((int)customTaskPane.DockPosition);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					Pane.A();
				}
			}
			customTaskPane = null;
		}
	}

	private static void A()
	{
		IEnumerator<CustomTaskPane> enumerator = default(IEnumerator<CustomTaskPane>);
		try
		{
			enumerator = NG.A.CustomTaskPanes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomTaskPane current = enumerator.Current;
				try
				{
					CustomTaskPane customTaskPane = current;
					if (IsExplorerPane(current) && customTaskPane.Visible)
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
						customTaskPane.VisibleChanged -= B;
						customTaskPane.Visible = false;
						customTaskPane.VisibleChanged += B;
						Events.Disable(NG.A.Application);
					}
					customTaskPane = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		C(NG.A.Application);
		IsOpen = false;
		C();
	}

	private static void C(Application A)
	{
		RemoveEvents(A);
		JG.A(AllPresentations);
		AllPresentations = null;
	}

	private static void B()
	{
		checked
		{
			if (AllPresentations != null)
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
				try
				{
					ObservableCollection<PresentationItem> allPresentations = AllPresentations;
					int num = allPresentations.Count - 1;
					while (true)
					{
						if (num >= 0)
						{
							if (Operators.CompareString(((BaseItem)allPresentations[num]).Label, ClosedPresentationName, TextCompare: false) == 0)
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
								allPresentations.RemoveAt(num);
								break;
							}
							num += -1;
							continue;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
						break;
					}
					allPresentations = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			ClosedPresentationName = "";
		}
	}

	private static void C()
	{
		KG.A.InvalidateControl(Pane.m_A);
	}

	public static bool IsExplorerPane(CustomTaskPane ctp)
	{
		return Operators.CompareString(ctp.Title, PANE_TITLE, TextCompare: false) == 0;
	}

	internal static wpfExplorer A(ctpExplorer2 A)
	{
		return (wpfExplorer)A.ElementHost1.Child;
	}
}
