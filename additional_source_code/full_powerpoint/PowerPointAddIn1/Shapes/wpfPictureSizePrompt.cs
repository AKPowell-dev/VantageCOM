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

namespace PowerPointAddIn1.Shapes;

[DesignerGenerated]
public sealed class wpfPictureSizePrompt : Window, IComponentConnector
{
	[AccessedThroughProperty("radWidth")]
	[CompilerGenerated]
	private RadioButton A;

	[AccessedThroughProperty("radHeight")]
	[CompilerGenerated]
	private RadioButton B;

	[AccessedThroughProperty("radBoth")]
	[CompilerGenerated]
	private RadioButton C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkApplyAll")]
	private CheckBox A;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button A;

	private bool A;

	internal virtual RadioButton radWidth
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

	internal virtual RadioButton radHeight
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

	internal virtual RadioButton radBoth
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

	internal virtual CheckBox chkApplyAll
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
				switch (1)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	public wpfPictureSizePrompt()
	{
		base.Loaded += wpfPictureSizePrompt_Loaded;
		base.Closing += wpfPictureSizePrompt_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void wpfPictureSizePrompt_Loaded(object sender, RoutedEventArgs e)
	{
		switch (PB.Settings.PictureSizePrompt)
		{
		case 1:
			radWidth.IsChecked = true;
			break;
		case 2:
			radHeight.IsChecked = true;
			break;
		case 0:
			radBoth.IsChecked = true;
			break;
		}
	}

	private void wpfPictureSizePrompt_Closing(object sender, CancelEventArgs e)
	{
		bool flag = true;
		bool flag2 = flag;
		bool? isChecked = radWidth.IsChecked;
		bool? obj;
		if (!isChecked.HasValue)
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
			obj = null;
		}
		else
		{
			obj = flag2 == (isChecked == true);
		}
		isChecked = obj;
		if (isChecked == true)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					PB.Settings.PictureSizePrompt = 1;
					return;
				}
			}
		}
		flag2 = flag;
		isChecked = radHeight.IsChecked;
		bool? obj2;
		if (!isChecked.HasValue)
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
			obj2 = null;
		}
		else
		{
			obj2 = flag2 == (isChecked == true);
		}
		isChecked = obj2;
		if (isChecked == true)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					PB.Settings.PictureSizePrompt = 2;
					return;
				}
			}
		}
		flag2 = flag;
		isChecked = radBoth.IsChecked;
		if ((isChecked.HasValue ? new bool?(flag2 == (isChecked == true)) : ((bool?)null)) != true)
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
			PB.Settings.PictureSizePrompt = 0;
			return;
		}
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(AH.A(91836), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			radWidth = (RadioButton)target;
			return;
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					radHeight = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radBoth = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkApplyAll = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
