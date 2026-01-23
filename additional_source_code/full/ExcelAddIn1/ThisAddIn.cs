using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using ExcelAddIn1.Audit.Check.UI;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Keyboard;
using ExcelAddIn1.Library2.UI;
using ExcelAddIn1.UndoRedo;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.Config;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Pitchly.Retrieval;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace ExcelAddIn1;

[StartupObject(0)]
[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
public sealed class ThisAddIn : AddInBase
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<int?> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int? A()
		{
			int? result;
			try
			{
				Microsoft.Office.Interop.Excel.Window activeWindow = MH.A.Application.ActiveWindow;
				int value;
				if (activeWindow == null)
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
					value = 0;
				}
				else
				{
					value = activeWindow.Hwnd;
				}
				return value;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = 0;
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private clsRouter m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("CustomTaskPanes")]
	private CustomTaskPaneCollection m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("VstoSmartTags")]
	private SmartTagCollection m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Application")]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	private Microsoft.Office.Interop.Excel.Application m_A;

	internal virtual CustomTaskPaneCollection CustomTaskPanes
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual SmartTagCollection VstoSmartTags
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual Microsoft.Office.Interop.Excel.Application Application
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			AppEvents_WorkbookOpenEventHandler handler = A;
			Microsoft.Office.Interop.Excel.Application application = this.m_A;
			if (application != null)
			{
				new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175844)).RemoveEventHandler(application, handler);
			}
			this.m_A = value;
			application = this.m_A;
			if (application != null)
			{
				new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(175844)).AddEventHandler(application, handler);
			}
		}
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
		: base((Microsoft.Office.Tools.Factory)factory, serviceProvider, VH.A(206832), VH.A(206843))
	{
		base.Startup += A;
		base.Shutdown += B;
		MH.A = factory;
	}

	private void A(object A, EventArgs B)
	{
		Events.Add();
		Microsoft.Office.Interop.Excel.Application application = Application;
		XmlDocument xmlSettings = clsUtilities.StartupProcedures(application.Version, Conversions.ToString(application.Build), application.LanguageSettings, application.UserControl, VH.A(169659));
		application = null;
		Base.PaneReopener = ExcelAddIn1.Library2.UI.Pane.B;
		Base.DisableInterruptKey = this.A;
		Func<int?> getOfficeAppActiveWinHwnd;
		if (_Closure_0024__.A == null)
		{
			getOfficeAppActiveWinHwnd = (_Closure_0024__.A = [SpecialName] () =>
			{
				int? result;
				try
				{
					Microsoft.Office.Interop.Excel.Window activeWindow = MH.A.Application.ActiveWindow;
					int value;
					if (activeWindow == null)
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
						value = 0;
					}
					else
					{
						value = activeWindow.Hwnd;
					}
					return value;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					result = 0;
					ProjectData.ClearProjectError();
				}
				return result;
			});
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			getOfficeAppActiveWinHwnd = _Closure_0024__.A;
		}
		UIFormsExtensions.GetOfficeAppActiveWinHwnd = getOfficeAppActiveWinHwnd;
		try
		{
			ThisAddIn.B();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			clsUtilities.StartupProcedures(xmlSettings);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		xmlSettings = null;
		try
		{
			if (System.Windows.Application.Current != null)
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
				new System.Windows.Application().ShutdownMode = ShutdownMode.OnExplicitShutdown;
				return;
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
	}

	private void B(object A, EventArgs B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				wpfAudit instanceRunningAnalysis;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 203:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_002f;
						case 4:
							goto IL_0036;
						case 5:
							goto IL_003d;
						case 6:
							goto IL_0050;
						case 7:
							goto IL_0057;
						case 8:
							goto IL_005e;
						case 9:
							goto IL_0066;
						case 10:
							goto IL_006e;
						case 11:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 12:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_006e:
					num2 = 10;
					K.Settings.Save();
					break;
					IL_0007:
					num2 = 2;
					instanceRunningAnalysis = wpfAudit.InstanceRunningAnalysis;
					if (instanceRunningAnalysis == null)
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
					}
					else
					{
						instanceRunningAnalysis.A(A: true);
					}
					goto IL_002f;
					IL_0050:
					num2 = 6;
					NavAid.Disable();
					goto IL_0057;
					IL_005e:
					num2 = 8;
					ExcelAddIn1.View.Gridlines.AutoHide(blnEnabled: false);
					goto IL_0066;
					IL_0066:
					num2 = 9;
					Core.Disable();
					goto IL_006e;
					IL_0057:
					num2 = 7;
					Shortcuts.Remove();
					goto IL_005e;
					IL_002f:
					num2 = 3;
					Events.Remove();
					goto IL_0036;
					IL_0036:
					num2 = 4;
					C();
					goto IL_003d;
					IL_003d:
					num2 = 5;
					if (NavAid.Enabled)
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
						goto IL_0050;
					}
					goto IL_0057;
					end_IL_0000_2:
					break;
				}
				num2 = 11;
				DataManagerV4.CancelFullRetrievalIfRunning(true);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 203;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
	{
		return new RibbonXL();
	}

	private void A()
	{
		try
		{
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (application.EnableCancelKey == XlEnableCancelKey.xlDisabled)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			application.EnableCancelKey = XlEnableCancelKey.xlDisabled;
			application = null;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	protected override object RequestComAddInAutomationService()
	{
		if (this.m_A == null)
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
			this.m_A = new clsRouter();
		}
		return this.m_A;
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (A.ReadOnly)
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
			if (Access.UserHasAccess(A.FullName))
			{
				return;
			}
			try
			{
				Application.DisplayAlerts = false;
				A.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				System.Windows.Forms.MessageBox.Show(VH.A(206567), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
			finally
			{
				Application.DisplayAlerts = true;
			}
		}
	}

	private static void B()
	{
		if (MH.A.Application.Visible)
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
					clsUtilities.InitializeMacabacus();
					return;
				}
			}
		}
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(206745)).AddEventHandler(MH.A.Application, new AppEvents_WindowActivateEventHandler(A));
	}

	private static void C()
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(206745)).RemoveEventHandler(MH.A.Application, new AppEvents_WindowActivateEventHandler(A));
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A, Microsoft.Office.Interop.Excel.Window B)
	{
		C();
		clsUtilities.InitializeMacabacus();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	protected override void Initialize()
	{
		base.Initialize();
		Application = GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), VH.A(124603));
		MH.A = this;
		System.Windows.Forms.Application.EnableVisualStyles();
		D();
		I();
		J();
		E();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	protected override void FinishInitialization()
	{
		OnStartup();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	protected override void InitializeDataBindings()
	{
		G();
		F();
		H();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void D()
	{
		if (base.DataHost == null || !base.DataHost.IsCacheInitialized)
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
			base.DataHost.FillCachedData(this);
			return;
		}
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void E()
	{
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void F()
	{
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[DebuggerNonUserCode]
	private void A(string A)
	{
		base.DataHost.StartCaching(this, A);
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[DebuggerNonUserCode]
	private void B(string A)
	{
		base.DataHost.StopCaching(this, A);
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[DebuggerNonUserCode]
	private bool A(string A)
	{
		return base.DataHost.IsCached(this, A);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void G()
	{
		BeginInit();
		CustomTaskPanes.BeginInit();
		VstoSmartTags.BeginInit();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void H()
	{
		VstoSmartTags.EndInit();
		CustomTaskPanes.EndInit();
		EndInit();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void I()
	{
		CustomTaskPanes = MH.A.CreateCustomTaskPaneCollection(null, null, VH.A(206774), VH.A(206774), this);
		VstoSmartTags = MH.A.CreateSmartTagCollection(null, null, VH.A(206805), VH.A(206805), this);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void J()
	{
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[DebuggerNonUserCode]
	private bool B(string A)
	{
		return base.DataHost.NeedsFill(this, A);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	protected override void OnShutdown()
	{
		VstoSmartTags.Dispose();
		CustomTaskPanes.Dispose();
		base.OnShutdown();
	}
}
