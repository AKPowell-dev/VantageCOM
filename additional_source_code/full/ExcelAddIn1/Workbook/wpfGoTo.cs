using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

[DesignerGenerated]
public sealed class wpfGoTo : System.Windows.Window, IComponentConnector
{
	private List<Range> A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxCells")]
	private System.Windows.Controls.ListBox A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button A;

	private bool A;

	internal virtual System.Windows.Controls.ListBox lbxCells
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
			SelectionChangedEventHandler value2 = lbxCells_SelectionChanged;
			System.Windows.Controls.ListBox listBox = this.A;
			if (listBox != null)
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
				listBox.SelectionChanged -= value2;
			}
			this.A = value;
			listBox = this.A;
			if (listBox == null)
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
				listBox.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			Button button = this.A;
			if (button != null)
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

	public wpfGoTo(List<Range> listRanges)
	{
		base.Loaded += wpfGoTo_Loaded;
		base.Closing += wpfGoTo_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.A = listRanges;
	}

	private void wpfGoTo_Loaded(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		List<string> list = default(List<string>);
		List<Range>.Enumerator enumerator = default(List<Range>.Enumerator);
		Range current = default(Range);
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
				case 297:
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
							goto IL_0010;
						case 4:
							goto IL_0029;
						case 5:
							goto IL_0081;
						case 6:
							goto IL_00a1;
						case 7:
							goto IL_00b0;
						case 8:
							goto IL_00c1;
						case 9:
							goto IL_00c6;
						case 10:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 11:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00c1:
					num2 = 8;
					list = null;
					goto IL_00c6;
					IL_0007:
					num2 = 2;
					list = new List<string>();
					goto IL_0010;
					IL_0010:
					num2 = 3;
					enumerator = this.A.GetEnumerator();
					goto IL_0083;
					IL_0083:
					if (enumerator.MoveNext())
					{
						current = enumerator.Current;
						goto IL_0029;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_00a1;
					IL_00c6:
					num2 = 9;
					base.MinHeight = base.ActualHeight;
					break;
					IL_0029:
					num2 = 4;
					list.Add(current.Worksheet.Name + VH.A(7827) + current.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
					goto IL_0081;
					IL_0081:
					num2 = 5;
					goto IL_0083;
					IL_00a1:
					num2 = 6;
					((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
					goto IL_00b0;
					IL_00b0:
					num2 = 7;
					lbxCells.ItemsSource = list;
					goto IL_00c1;
					end_IL_0000_2:
					break;
				}
				num2 = 10;
				base.MinWidth = base.ActualWidth;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 297;
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

	private void lbxCells_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
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
				case 123:
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
							goto IL_0024;
						case 4:
							goto IL_0031;
						case 5:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 6:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0031:
					num2 = 4;
					range.Application.Goto(range, RuntimeHelpers.GetObjectValue(Missing.Value));
					break;
					IL_0007:
					num2 = 2;
					range = this.A[lbxCells.SelectedIndex];
					goto IL_0024;
					IL_0024:
					num2 = 3;
					range.Worksheet.Activate();
					goto IL_0031;
					end_IL_0000_2:
					break;
				}
				num2 = 5;
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 123;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = false;
		Close();
	}

	private void wpfGoTo_Closing(object sender, CancelEventArgs e)
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
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					break;
				case 46:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
							goto end_IL_0000_3;
						}
						goto default;
					}
					end_IL_0000_2:
					break;
				}
				num2 = 2;
				this.A = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 46;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
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
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A = true;
			Uri resourceLocator = new Uri(VH.A(186532), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					lbxCells = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnClose = (Button)target;
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
