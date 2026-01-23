using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

[DesignerGenerated]
public sealed class wpfLabelPoints : Window, IComponentConnector
{
	[CompilerGenerated]
	[AccessedThroughProperty("cbxLast")]
	private ComboBox A;

	[AccessedThroughProperty("cbxFirst")]
	[CompilerGenerated]
	private ComboBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxMaximum")]
	private ComboBox C;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxMinimum")]
	private ComboBox D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkColor")]
	private CheckBox A;

	[AccessedThroughProperty("chkBold")]
	[CompilerGenerated]
	private CheckBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxRotation")]
	private ComboBox E;

	[CompilerGenerated]
	[AccessedThroughProperty("radValue")]
	private RadioButton A;

	[CompilerGenerated]
	[AccessedThroughProperty("radSeriesName")]
	private RadioButton B;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool A;

	internal virtual ComboBox cbxLast
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual ComboBox cbxFirst
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual ComboBox cbxMaximum
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal virtual ComboBox cbxMinimum
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal virtual CheckBox chkColor
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual CheckBox chkBold
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual ComboBox cbxRotation
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal virtual RadioButton radValue
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	internal virtual RadioButton radSeriesName
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.A = value;
			button = this.A;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public wpfLabelPoints()
	{
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (A)
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
			A = true;
			Uri resourceLocator = new Uri(VH.A(67963), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			cbxLast = (ComboBox)target;
			return;
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					cbxFirst = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					cbxMaximum = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxMinimum = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkColor = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			chkBold = (CheckBox)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxRotation = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					radValue = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					radSeriesName = (RadioButton)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 10:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				btnOk = (Button)target;
				return;
			}
		case 11:
			btnCancel = (Button)target;
			break;
		default:
			A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
