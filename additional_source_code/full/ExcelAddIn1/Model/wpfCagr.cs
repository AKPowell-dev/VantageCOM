using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using ExcelAddIn1.View;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfCagr : System.Windows.Window, IComponentConnector
{
	private Microsoft.Office.Interop.Excel.Application m_A;

	private static readonly string m_A = VH.A(91517);

	[CompilerGenerated]
	[AccessedThroughProperty("txtAddress")]
	private TextBox m_A;

	[AccessedThroughProperty("btnRangeEdit")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("cbxInterval")]
	[CompilerGenerated]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button C;

	private bool m_A;

	internal virtual TextBox txtAddress
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

	internal virtual Button btnRangeEdit
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
			RoutedEventHandler value2 = btnRangeEdit_Click;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual ComboBox cbxInterval
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	public wpfCagr()
	{
		base.Loaded += wpfCagr_Loaded;
		base.Closing += wpfCagr_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = MH.A.Application;
		this.m_A.ScreenUpdating = true;
	}

	private void wpfCagr_Loaded(object sender, RoutedEventArgs e)
	{
		btnOk.Focus();
		A((Range)base.Tag);
	}

	private void btnRangeEdit_Click(object sender, RoutedEventArgs e)
	{
		string text = "";
		try
		{
			text = QuickCagr.RelativeAddress((Range)base.Tag, (Worksheet)this.m_A.ActiveSheet, blnAbsolute: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Range range = (Range)this.m_A.InputBox(VH.A(62623), VH.A(40448), text, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), 8);
			if (range != null)
			{
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
					if (Operators.ConditionalCompareObjectGreater(range.Rows.CountLarge, 1, TextCompare: false))
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
						if (Operators.ConditionalCompareObjectGreater(range.Columns.CountLarge, 1, TextCompare: false))
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								Forms.WarningMessage(VH.A(91213));
								break;
							}
							break;
						}
					}
					if (range.Areas.Count > 1)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(VH.A(91290));
							break;
						}
					}
					else
					{
						txtAddress.Text = QuickCagr.RelativeAddress(range, (Worksheet)this.m_A.ActiveSheet, blnAbsolute: true);
						B((Range)base.Tag);
						base.Tag = range;
						A(range);
					}
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
		Focus();
		btnOk.Focus();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		Range range = null;
		try
		{
			range = (Range)base.Tag;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range == null)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(VH.A(91359));
					return;
				}
			}
		}
		range = null;
		base.DialogResult = true;
	}

	private void wpfCagr_Closing(object sender, CancelEventArgs e)
	{
		B((Range)base.Tag);
		this.m_A.ScreenUpdating = false;
		this.m_A = null;
	}

	private void A(Range A)
	{
		try
		{
			FormatConditions formatConditions = A.FormatConditions;
			formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), wpfCagr.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			object instance = NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null);
			NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { NavAid.PATTERN_COLOR_GREEN }, null, null, OptimisticSet: false, RValueBase: true);
			NewLateBinding.LateSetComplex(instance, null, VH.A(36212), new object[1] { XlPattern.xlPatternGray50 }, null, null, OptimisticSet: false, RValueBase: true);
			instance = null;
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Range A)
	{
		try
		{
			FormatConditions formatConditions = A.FormatConditions;
			for (int i = formatConditions.Count; i >= 1; i = checked(i + -1))
			{
				try
				{
					FormatCondition formatCondition = (FormatCondition)formatConditions.Item(i);
					if (formatCondition.Type == 2)
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
						if (Operators.CompareString(formatCondition.Formula1, wpfCagr.m_A, TextCompare: false) == 0)
						{
							formatCondition.Delete();
							break;
						}
					}
					formatCondition = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			formatConditions = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_A)
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
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(91406), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					txtAddress = (TextBox)target;
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
					btnRangeEdit = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					cbxInterval = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (6)
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
		this.m_A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
