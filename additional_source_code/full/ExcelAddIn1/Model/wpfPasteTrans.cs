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

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfPasteTrans : Window, IComponentConnector
{
	[AccessedThroughProperty("optLinks")]
	[CompilerGenerated]
	private RadioButton A;

	[AccessedThroughProperty("optFormats")]
	[CompilerGenerated]
	private RadioButton B;

	[CompilerGenerated]
	[AccessedThroughProperty("optExact")]
	private RadioButton C;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button B;

	private bool A;

	internal virtual RadioButton optLinks
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = this.A;
			if (radioButton != null)
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
				radioButton.GotFocus -= value2;
			}
			this.A = value;
			radioButton = this.A;
			if (radioButton != null)
			{
				radioButton.GotFocus += value2;
			}
		}
	}

	internal virtual RadioButton optFormats
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = this.B;
			if (radioButton != null)
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
				radioButton.GotFocus -= value2;
			}
			this.B = value;
			radioButton = this.B;
			if (radioButton == null)
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
				radioButton.GotFocus += value2;
				return;
			}
		}
	}

	internal virtual RadioButton optExact
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
			RoutedEventHandler value2 = CheckOnFocus;
			RadioButton radioButton = C;
			if (radioButton != null)
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
				radioButton.GotFocus -= value2;
			}
			C = value;
			radioButton = C;
			if (radioButton == null)
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
				radioButton.GotFocus += value2;
				return;
			}
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
				button.Click -= value2;
			}
			this.A = value;
			button = this.A;
			if (button != null)
			{
				button.Click += value2;
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

	public wpfPasteTrans()
	{
		base.Loaded += wpfPasteTrans_Loaded;
		int try0018_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				string pasteTranspose;
				switch (try0018_dispatch)
				{
				default:
					num2 = 1;
					InitializeComponent();
					goto IL_0020;
				case 271:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0018;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0020;
						case 3:
							goto IL_002f;
						case 4:
							goto IL_0036;
						case 6:
							goto IL_0088;
						case 8:
							goto IL_009d;
						case 10:
							goto end_IL_0018_2;
						default:
							goto end_IL_0018;
						case 5:
						case 7:
						case 9:
						case 11:
						case 12:
							goto end_IL_0018_3;
						}
						goto default;
					}
					IL_009d:
					num2 = 8;
					optExact.IsChecked = true;
					goto end_IL_0018_3;
					IL_0020:
					num2 = 2;
					base.Icon = Forms.GetIcon();
					goto IL_002f;
					IL_002f:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0036;
					IL_0036:
					num2 = 4;
					pasteTranspose = K.Settings.PasteTranspose;
					if (Operators.CompareString(pasteTranspose, VH.A(91683), TextCompare: false) == 0)
					{
						goto IL_0088;
					}
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
					if (Operators.CompareString(pasteTranspose, VH.A(91694), TextCompare: false) != 0)
					{
						break;
					}
					goto IL_009d;
					IL_0088:
					num2 = 6;
					optLinks.IsChecked = true;
					goto end_IL_0018_3;
					end_IL_0018_2:
					break;
				}
				num2 = 10;
				optFormats.IsChecked = true;
				break;
				end_IL_0018:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0018_dispatch = 271;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0018_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	private void wpfPasteTrans_Loaded(object sender, RoutedEventArgs e)
	{
		optLinks.Focus();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		if (optLinks.IsChecked == true)
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
			K.Settings.PasteTranspose = VH.A(91683);
		}
		else if (optExact.IsChecked == true)
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
			K.Settings.PasteTranspose = VH.A(91694);
		}
		else
		{
			K.Settings.PasteTranspose = VH.A(91705);
		}
		base.DialogResult = true;
	}

	private void CheckOnFocus(object sender, RoutedEventArgs e)
	{
		((RadioButton)sender).IsChecked = true;
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
			Uri resourceLocator = new Uri(VH.A(91720), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			optLinks = (RadioButton)target;
			return;
		}
		if (connectionId == 2)
		{
			optFormats = (RadioButton)target;
			return;
		}
		if (connectionId == 3)
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
					optExact = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 4)
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
		if (connectionId == 5)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
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
