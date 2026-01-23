using System;
using System.CodeDom.Compiler;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format.ContextualPopups;

[DesignerGenerated]
public sealed class wpfPlotArea : System.Windows.Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ObservableCollection<ColorItem> m_A;

	private int m_A;

	private int B;

	private SolidColorBrush m_A;

	private Color m_A;

	private XmlNode m_A;

	private object m_A;

	[AccessedThroughProperty("bdrColors")]
	[CompilerGenerated]
	private System.Windows.Controls.Border m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("bdrPatterns")]
	private System.Windows.Controls.Border B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkColor")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPattern")]
	private System.Windows.Controls.CheckBox B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLineStyle")]
	private System.Windows.Controls.CheckBox C;

	[AccessedThroughProperty("chkLineWeight")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox D;

	private bool m_A;

	public ObservableCollection<ColorItem> Colors
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(146234));
		}
	}

	public int RowsCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(146247));
		}
	}

	public int ColumnsCount
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			A(VH.A(146266));
		}
	}

	public SolidColorBrush CurrentColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(146291));
		}
	}

	public Color SelectedColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.Border bdrColors
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

	internal virtual System.Windows.Controls.Border bdrPatterns
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

	internal virtual System.Windows.Controls.CheckBox chkColor
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
			RoutedEventHandler value2 = ToggleColors;
			RoutedEventHandler value3 = ToggleColors;
			System.Windows.Controls.CheckBox checkBox = this.m_A;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_A = value;
			checkBox = this.m_A;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPattern
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
			RoutedEventHandler value2 = TogglePatterns;
			RoutedEventHandler value3 = TogglePatterns;
			System.Windows.Controls.CheckBox checkBox = B;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			B = value;
			checkBox = B;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLineStyle
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

	internal virtual System.Windows.Controls.CheckBox chkLineWeight
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

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
		}
	}

	public wpfPlotArea(XmlNode nd, int intOleColor, object obj)
	{
		base.Closing += wpfColors_Closing;
		base.Loaded += wpfColors_Loaded;
		InitializeComponent();
		this.m_A = nd;
		this.m_A = RuntimeHelpers.GetObjectValue(obj);
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private void wpfColors_Closing(object sender, CancelEventArgs e)
	{
		try
		{
			if (this.m_A is PlotArea)
			{
				((PlotArea)this.m_A).Format.Fill.ForeColor.RGB = 0;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Colors = null;
		SelectedColor = default(Color);
		CurrentColor = null;
	}

	private void wpfColors_Loaded(object sender, RoutedEventArgs e)
	{
		try
		{
			CurrentColor = new SolidColorBrush(System.Windows.Media.Colors.YellowGreen);
			Colors = new ObservableCollection<ColorItem>();
			XmlNode xmlNode = this.m_A.SelectSingleNode(VH.A(146234));
			RowsCount = Conversions.ToInteger(xmlNode.Attributes[VH.A(2877)].Value);
			ColumnsCount = Conversions.ToInteger(xmlNode.Attributes[VH.A(2862)].Value);
			foreach (XmlNode item in xmlNode.SelectNodes(VH.A(55331)))
			{
				Colors.Add(new ColorItem(item, blnChecked: false));
			}
			xmlNode = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Interaction.MsgBox(ex2.Message);
			ProjectData.ClearProjectError();
		}
		base.Deactivated += PopupDeactivated;
	}

	private void MouseEnterColor(object sender, MouseEventArgs e)
	{
		System.Windows.Controls.Border obj = (System.Windows.Controls.Border)sender;
		obj.BorderBrush = obj.Background;
	}

	private void MouseLeaveColor(object sender, MouseEventArgs e)
	{
		((System.Windows.Controls.Border)sender).BorderBrush = null;
	}

	private void ColorSelected(object sender, MouseButtonEventArgs e)
	{
		System.Windows.Controls.Border border = (System.Windows.Controls.Border)sender;
		SelectedColor = ((SolidColorBrush)border.DataContext).Color;
		border = null;
		Close();
	}

	private void MakeTransparent(object sender, RoutedEventArgs e)
	{
		_ = (System.Windows.Controls.Border)sender;
		SelectedColor = System.Windows.Media.Colors.Transparent;
		Close();
	}

	private void ToggleColors(object sender, RoutedEventArgs e)
	{
		A();
		if (chkColor.IsChecked != true)
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
			bdrColors.Visibility = Visibility.Visible;
			return;
		}
	}

	private void TogglePatterns(object sender, RoutedEventArgs e)
	{
		A();
		if (chkColor.IsChecked != true)
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
			bdrPatterns.Visibility = Visibility.Visible;
			return;
		}
	}

	private void A()
	{
		bdrColors.Visibility = Visibility.Collapsed;
		bdrPatterns.Visibility = Visibility.Collapsed;
	}

	private void PopupDeactivated(object sender, EventArgs e)
	{
		Close();
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
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(146316), UriKind.Relative);
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
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					bdrColors = (System.Windows.Controls.Border)target;
					return;
				}
			}
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
					bdrPatterns = (System.Windows.Controls.Border)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkColor = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 6:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				chkPattern = (System.Windows.Controls.CheckBox)target;
				return;
			}
		case 7:
			chkLineStyle = (System.Windows.Controls.CheckBox)target;
			break;
		case 8:
			chkLineWeight = (System.Windows.Controls.CheckBox)target;
			break;
		default:
			this.m_A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 2)
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
			((System.Windows.Controls.Border)target).MouseEnter += MouseEnterColor;
			((System.Windows.Controls.Border)target).MouseLeave += MouseLeaveColor;
			((System.Windows.Controls.Border)target).MouseUp += ColorSelected;
		}
		if (connectionId != 4)
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
			((System.Windows.Controls.Border)target).MouseEnter += MouseEnterColor;
			((System.Windows.Controls.Border)target).MouseLeave += MouseLeaveColor;
			((System.Windows.Controls.Border)target).MouseUp += ColorSelected;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
