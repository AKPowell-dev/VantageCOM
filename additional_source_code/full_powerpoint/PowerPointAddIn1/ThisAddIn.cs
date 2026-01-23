using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Security.Permissions;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Pitchly.Retrieval;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.Library2.UI;

namespace PowerPointAddIn1;

[StartupObject(0)]
[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
public sealed class ThisAddIn : AddInBase
{
	[AccessedThroughProperty("CustomTaskPanes")]
	[CompilerGenerated]
	private CustomTaskPaneCollection m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Application")]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	private Microsoft.Office.Interop.PowerPoint.Application m_A;

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

	internal virtual Microsoft.Office.Interop.PowerPoint.Application Application
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public ThisAddIn(Factory factory, IServiceProvider serviceProvider)
		: base(factory, serviceProvider, AH.A(170559), AH.A(170570))
	{
		base.Startup += A;
		base.Shutdown += B;
		NG.A = factory;
	}

	private void A(object A, EventArgs B)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = Application;
		XmlDocument xmlSettings = clsUtilities.StartupProcedures(application.Version, application.Build, application.LanguageSettings, true, AH.A(116727));
		application = null;
		Base.PaneReopener = PowerPointAddIn1.Library2.UI.Pane.B;
		UIFormsExtensions.GetOfficeAppActiveWinHwnd = [SpecialName] () =>
		{
			int? result;
			try
			{
				DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
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
					value = activeWindow.HWND;
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
		};
		clsEvents.Add();
		clsUtilities.StartupProcedures(xmlSettings);
		try
		{
			if (System.Windows.Application.Current == null)
			{
				new System.Windows.Application().ShutdownMode = ShutdownMode.OnExplicitShutdown;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(object A, EventArgs B)
	{
		Events.Disable(Application);
		PB.Settings.Save();
		clsEvents.Remove();
		DataManagerV4.CancelFullRetrievalIfRunning(true);
	}

	protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
	{
		return new RibbonPP();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	protected override void Initialize()
	{
		base.Initialize();
		Application = GetHostItem<Microsoft.Office.Interop.PowerPoint.Application>(typeof(Microsoft.Office.Interop.PowerPoint.Application), AH.A(170505));
		NG.A = this;
		System.Windows.Forms.Application.EnableVisualStyles();
		A();
		F();
		G();
		B();
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
		D();
		C();
		E();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void A()
	{
		if (base.DataHost == null)
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
			if (!base.DataHost.IsCacheInitialized)
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
				base.DataHost.FillCachedData(this);
				return;
			}
		}
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void B()
	{
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void C()
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
	private void D()
	{
		BeginInit();
		CustomTaskPanes.BeginInit();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void E()
	{
		CustomTaskPanes.EndInit();
		EndInit();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void F()
	{
		CustomTaskPanes = NG.A.CreateCustomTaskPaneCollection(null, null, AH.A(170528), AH.A(170528), this);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
	[DebuggerNonUserCode]
	private void G()
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
		CustomTaskPanes.Dispose();
		base.OnShutdown();
	}
}
