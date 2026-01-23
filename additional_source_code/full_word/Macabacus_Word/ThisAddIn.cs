using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Config;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Pitchly.Retrieval;
using MacabacusMacros.UI.FormsExtensions;
using Macabacus_Word.Library2.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace Macabacus_Word;

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
				Microsoft.Office.Interop.Word.Window activeWindow = PC.A.Application.ActiveWindow;
				int value;
				if (activeWindow == null)
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

	[AccessedThroughProperty("wdApp")]
	[CompilerGenerated]
	private Microsoft.Office.Interop.Word.Application m_A;

	private clsRouter m_A;

	[AccessedThroughProperty("CustomTaskPanes")]
	[CompilerGenerated]
	private CustomTaskPaneCollection m_A;

	[AccessedThroughProperty("VstoSmartTags")]
	[CompilerGenerated]
	private SmartTagCollection m_A;

	[AccessedThroughProperty("Application")]
	[CompilerGenerated]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	private Microsoft.Office.Interop.Word.Application m_B;

	private virtual Microsoft.Office.Interop.Word.Application wdApp
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

	internal virtual Microsoft.Office.Interop.Word.Application Application
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			ApplicationEvents4_WindowSelectionChangeEventHandler handler = A;
			Microsoft.Office.Interop.Word.Application application = this.m_B;
			if (application != null)
			{
				new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(application, handler);
			}
			this.m_B = value;
			application = this.m_B;
			if (application == null)
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
				new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).AddEventHandler(application, handler);
				return;
			}
		}
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider)
		: base((Microsoft.Office.Tools.Factory)factory, serviceProvider, XC.A(42768), XC.A(42779))
	{
		base.Startup += [SpecialName] (object A, EventArgs B) =>
		{
			this.A();
		};
		base.Shutdown += [SpecialName] (object A, EventArgs B) =>
		{
			this.B();
		};
		PC.A = factory;
	}

	private void A()
	{
		Microsoft.Office.Interop.Word.Application application = Application;
		XmlDocument xmlSettings = clsUtilities.StartupProcedures(application.Version, application.Build, application.LanguageSettings, application.UserControl, XC.A(18421));
		application = null;
		Base.PaneReopener = Pane.B;
		Base.DisableInterruptKey = C;
		Func<int?> getOfficeAppActiveWinHwnd;
		if (_Closure_0024__.A == null)
		{
			getOfficeAppActiveWinHwnd = (_Closure_0024__.A = [SpecialName] () =>
			{
				int? result;
				try
				{
					Microsoft.Office.Interop.Word.Window activeWindow = PC.A.Application.ActiveWindow;
					int value;
					if (activeWindow == null)
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
			getOfficeAppActiveWinHwnd = _Closure_0024__.A;
		}
		UIFormsExtensions.GetOfficeAppActiveWinHwnd = getOfficeAppActiveWinHwnd;
		try
		{
			clsUtilities.InitializeMacabacus();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		clsUtilities.StartupProcedures(xmlSettings);
		xmlSettings = null;
		try
		{
			if (System.Windows.Application.Current == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					new System.Windows.Application().ShutdownMode = ShutdownMode.OnExplicitShutdown;
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		clsEvents.A();
	}

	private void B()
	{
		N.Settings.Save();
		clsEvents.B();
		DataManagerV4.CancelFullRetrievalIfRunning(true);
	}

	protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
	{
		return new RibbonWD();
	}

	protected override object RequestComAddInAutomationService()
	{
		if (this.m_A == null)
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
			this.m_A = new clsRouter();
		}
		return this.m_A;
	}

	private void C()
	{
		try
		{
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			if (application.EnableCancelKey == WdEnableCancelKey.wdCancelDisabled)
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
			application.EnableCancelKey = WdEnableCancelKey.wdCancelDisabled;
			application = null;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void A(Microsoft.Office.Interop.Word.Document A)
	{
	}

	private void B(Microsoft.Office.Interop.Word.Document A)
	{
	}

	private void A(Selection A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		clsGlobals clsGlobals2 = default(clsGlobals);
		IRibbonUI ribbonUI = default(IRibbonUI);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 284:
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
							goto IL_000f;
						case 4:
							goto IL_001c;
						case 5:
							goto IL_001e;
						case 6:
							goto IL_0027;
						case 7:
							goto IL_0033;
						case 8:
							goto IL_003c;
						case 9:
							goto IL_0046;
						case 10:
							goto IL_0051;
						case 11:
							goto IL_005c;
						case 12:
							goto IL_0067;
						case 13:
							goto IL_0072;
						case 14:
							goto IL_007d;
						case 15:
							goto IL_0088;
						case 16:
							goto IL_0093;
						case 17:
							goto IL_009e;
						case 18:
							goto IL_00a9;
						case 19:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 20:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00a9:
					num2 = 18;
					clsGlobals2.StyleCycleIndex = 0;
					break;
					IL_0007:
					num2 = 2;
					ribbonUI = NC.A;
					goto IL_000f;
					IL_000f:
					num2 = 3;
					ribbonUI.InvalidateControl(clsColors.LAST_FILL_COLOR_BUTTON);
					goto IL_001c;
					IL_001c:
					ribbonUI = null;
					goto IL_001e;
					IL_001e:
					num2 = 5;
					if (NC.A == null)
					{
						goto IL_0027;
					}
					goto IL_0033;
					IL_0027:
					num2 = 6;
					NC.A = new clsGlobals();
					goto IL_0033;
					IL_0033:
					num2 = 7;
					clsGlobals2 = NC.A;
					goto IL_003c;
					IL_003c:
					num2 = 8;
					clsGlobals2.FontColorCycle = 0;
					goto IL_0046;
					IL_0046:
					num2 = 9;
					clsGlobals2.FillColorCycle = 0;
					goto IL_0051;
					IL_0051:
					num2 = 10;
					clsGlobals2.BorderColorCycle = 0;
					goto IL_005c;
					IL_005c:
					num2 = 11;
					clsGlobals2.StyleCycle1 = 0;
					goto IL_0067;
					IL_0067:
					num2 = 12;
					clsGlobals2.StyleCycle2 = 0;
					goto IL_0072;
					IL_0072:
					num2 = 13;
					clsGlobals2.StyleCycle3 = 0;
					goto IL_007d;
					IL_007d:
					num2 = 14;
					clsGlobals2.StyleCycle4 = 0;
					goto IL_0088;
					IL_0088:
					num2 = 15;
					clsGlobals2.StyleCycle5 = 0;
					goto IL_0093;
					IL_0093:
					num2 = 16;
					clsGlobals2.StyleCycle6 = 0;
					goto IL_009e;
					IL_009e:
					num2 = 17;
					clsGlobals2.StyleCycle = 0;
					goto IL_00a9;
					end_IL_0000_2:
					break;
				}
				clsGlobals2 = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 284;
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	protected override void Initialize()
	{
		base.Initialize();
		Application = GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), XC.A(42687));
		PC.A = this;
		System.Windows.Forms.Application.EnableVisualStyles();
		D();
		I();
		J();
		E();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	protected override void FinishInitialization()
	{
		OnStartup();
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	protected override void InitializeDataBindings()
	{
		G();
		F();
		H();
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	private void D()
	{
		if (base.DataHost == null)
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
			if (base.DataHost.IsCacheInitialized)
			{
				base.DataHost.FillCachedData(this);
			}
			return;
		}
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	private void E()
	{
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
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

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Advanced)]
	private bool A(string A)
	{
		return base.DataHost.IsCached(this, A);
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	private void G()
	{
		BeginInit();
		CustomTaskPanes.BeginInit();
		VstoSmartTags.BeginInit();
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void H()
	{
		VstoSmartTags.EndInit();
		CustomTaskPanes.EndInit();
		EndInit();
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	private void I()
	{
		CustomTaskPanes = PC.A.CreateCustomTaskPaneCollection(null, null, XC.A(42710), XC.A(42710), this);
		VstoSmartTags = PC.A.CreateSmartTagCollection(null, null, XC.A(42741), XC.A(42741), this);
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	private void J()
	{
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	[DebuggerNonUserCode]
	private bool B(string A)
	{
		return base.DataHost.NeedsFill(this, A);
	}

	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	protected override void OnShutdown()
	{
		VstoSmartTags.Dispose();
		CustomTaskPanes.Dispose();
		base.OnShutdown();
	}

	[SpecialName]
	[CompilerGenerated]
	private void A(object A, EventArgs B)
	{
		this.A();
	}

	[SpecialName]
	[CompilerGenerated]
	private void B(object A, EventArgs B)
	{
		this.B();
	}
}
