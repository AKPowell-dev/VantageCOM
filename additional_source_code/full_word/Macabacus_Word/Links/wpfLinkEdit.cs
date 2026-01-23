using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

[DesignerGenerated]
public sealed class wpfLinkEdit : System.Windows.Window, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class AB
	{
		public string A;

		public wpfLinkEdit A;

		[SpecialName]
		internal void A()
		{
			this.A.cbxFiles.Items.Add(this.A);
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private List<object> m_A;

	private Edit.EditedShapes m_A;

	private bool m_A;

	private Document m_A;

	private Dictionary<string, WorkbookStruct> m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private List<Workbook> m_A;

	private bool m_B;

	private readonly string m_A;

	private readonly string m_B;

	private ObservableCollection<string> m_A;

	private ObservableCollection<ChartNameInfo> m_A;

	private int m_A;

	[AccessedThroughProperty("cbxFiles")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("btnApply")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBrowse")]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("grpRange")]
	private GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxRanges")]
	private ComboBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnViewRange")]
	private Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("grpChart")]
	private GroupBox m_B;

	[AccessedThroughProperty("cbxCharts")]
	[CompilerGenerated]
	private ComboBox m_C;

	[AccessedThroughProperty("btnViewChart")]
	[CompilerGenerated]
	private Button D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkGraphic")]
	private RadioButton m_A;

	[AccessedThroughProperty("chkPicture")]
	[CompilerGenerated]
	private RadioButton m_B;

	[AccessedThroughProperty("chkTable")]
	[CompilerGenerated]
	private RadioButton m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkEmbedded")]
	private RadioButton D;

	[AccessedThroughProperty("chkChart")]
	[CompilerGenerated]
	private RadioButton E;

	[AccessedThroughProperty("chkText")]
	[CompilerGenerated]
	private RadioButton F;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button F;

	private bool m_C;

	public Edit.EditedShapes ReturnValue
	{
		get
		{
			return this.m_A;
		}
		set
		{
		}
	}

	public ObservableCollection<string> Ranges
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(11706));
		}
	}

	public ObservableCollection<ChartNameInfo> Charts
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(11719));
		}
	}

	internal virtual ComboBox cbxFiles
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
			SelectionChangedEventHandler value2 = cbxFiles_SelectionChanged;
			KeyEventHandler value3 = cbxFiles_PreviewKeyDown;
			ComboBox comboBox = this.m_A;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
				comboBox.PreviewKeyDown -= value3;
			}
			this.m_A = value;
			comboBox = this.m_A;
			if (comboBox == null)
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
				comboBox.SelectionChanged += value2;
				comboBox.PreviewKeyDown += value3;
				return;
			}
		}
	}

	internal virtual Button btnApply
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
			RoutedEventHandler value2 = btnApply_Click;
			Button button = this.m_A;
			if (button != null)
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

	internal virtual Button btnBrowse
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
			RoutedEventHandler value2 = btnBrowse_Click;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual GroupBox grpRange
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

	internal virtual ComboBox cbxRanges
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
			this.m_B = value;
		}
	}

	internal virtual Button btnViewRange
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnViewRange_Click;
			Button button = this.m_C;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual GroupBox grpChart
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
			this.m_B = value;
		}
	}

	internal virtual ComboBox cbxCharts
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual Button btnViewChart
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnViewChart_Click;
			Button button = this.D;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.D = value;
			button = this.D;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual RadioButton chkGraphic
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

	internal virtual RadioButton chkPicture
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
			this.m_B = value;
		}
	}

	internal virtual RadioButton chkTable
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual RadioButton chkEmbedded
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

	internal virtual RadioButton chkChart
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	internal virtual RadioButton chkText
	{
		[CompilerGenerated]
		get
		{
			return this.F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.F = value;
		}
	}

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = E;
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
			E = value;
			button = E;
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
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			F = value;
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
				switch (2)
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

	public wpfLinkEdit(List<object> listObjects)
	{
		//IL_06d0: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_06dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_06fe: Expected I4, but got Unknown
		//IL_05d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_05da: Unknown result type (might be due to invalid IL or missing references)
		//IL_05dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_05de: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0613: Expected I4, but got Unknown
		//IL_0230: Unknown result type (might be due to invalid IL or missing references)
		//IL_0235: Unknown result type (might be due to invalid IL or missing references)
		//IL_0238: Unknown result type (might be due to invalid IL or missing references)
		//IL_0245: Unknown result type (might be due to invalid IL or missing references)
		//IL_0247: Unknown result type (might be due to invalid IL or missing references)
		//IL_0251: Unknown result type (might be due to invalid IL or missing references)
		//IL_0253: Unknown result type (might be due to invalid IL or missing references)
		//IL_0258: Unknown result type (might be due to invalid IL or missing references)
		//IL_025a: Unknown result type (might be due to invalid IL or missing references)
		//IL_025d: Unknown result type (might be due to invalid IL or missing references)
		//IL_025f: Invalid comparison between Unknown and I4
		//IL_0226: Unknown result type (might be due to invalid IL or missing references)
		//IL_022c: Unknown result type (might be due to invalid IL or missing references)
		//IL_046e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0473: Unknown result type (might be due to invalid IL or missing references)
		//IL_0475: Unknown result type (might be due to invalid IL or missing references)
		//IL_0477: Unknown result type (might be due to invalid IL or missing references)
		//IL_047a: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b0: Expected I4, but got Unknown
		//IL_0261: Unknown result type (might be due to invalid IL or missing references)
		//IL_0265: Invalid comparison between Unknown and I4
		base.Loaded += wpfLinkEdit_Loaded;
		base.Closing += wpfLinkEdit_Closing;
		this.m_A = true;
		this.m_B = false;
		this.m_A = XC.A(12960);
		this.m_B = XC.A(12989);
		this.m_A = -1;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		bool flag = false;
		bool flag2 = false;
		List<string> list = new List<string>();
		List<ImportType> list2 = new List<ImportType>();
		new List<string>();
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		int num4 = 0;
		int num5 = 0;
		Ranges = new ObservableCollection<string>();
		Charts = new ObservableCollection<ChartNameInfo>();
		this.m_A = listObjects;
		this.m_A = new Dictionary<string, WorkbookStruct>();
		this.m_A = PC.A.Application.ActiveDocument;
		this.m_A = new List<Workbook>();
		btnApply.Visibility = System.Windows.Visibility.Collapsed;
		try
		{
			this.m_A = A();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		int count = default(int);
		if (this.m_A != null)
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
			using (List<object>.Enumerator enumerator = this.m_A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					object[] array;
					bool[] array2;
					object value = NewLateBinding.LateGet(null, typeof(Common), XC.A(13018), array = new object[1] { objectValue }, null, null, array2 = new bool[1] { true });
					if (array2[0])
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
						objectValue = RuntimeHelpers.GetObjectValue(array[0]);
					}
					if (!Conversions.ToBoolean(value))
					{
						continue;
					}
					Type typeFromHandle = typeof(Common);
					string memberName = XC.A(11777);
					object[] obj = new object[1] { objectValue };
					array = obj;
					bool[] obj2 = new bool[1] { true };
					array2 = obj2;
					object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
					if (array2[0])
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
						objectValue = RuntimeHelpers.GetObjectValue(array[0]);
					}
					_003F val;
					if (obj3 == null)
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
						val = default(Link);
					}
					else
					{
						val = (Link)obj3;
					}
					Link val2 = (Link)val;
					list.Add(val2.Source);
					list2.Add(val2.Type);
					ImportType type = val2.Type;
					if (type - 6 > 2)
					{
						if ((int)type != 12)
						{
							flag = true;
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
					}
					flag2 = true;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_0283;
					}
					continue;
					end_IL_0283:
					break;
				}
			}
			count = this.m_A.Count;
		}
		list = list.Distinct().ToList();
		using (List<string>.Enumerator enumerator2 = list.GetEnumerator())
		{
			while (enumerator2.MoveNext())
			{
				string current = enumerator2.Current;
				A(current, null);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_02e5;
				}
				continue;
				end_IL_02e5:
				break;
			}
		}
		if (this.m_A != null)
		{
			Workbooks workbooks = this.m_A.Workbooks;
			try
			{
				IEnumerator enumerator3 = default(IEnumerator);
				try
				{
					enumerator3 = workbooks.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						Workbook workbook = (Workbook)enumerator3.Current;
						if (this.m_A.ContainsKey(workbook.FullName))
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
							B(workbook.FullName, workbook);
						}
						else
						{
							A(workbook.FullName, workbook);
						}
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0383;
						}
						continue;
						end_IL_0383:
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			MC.A(workbooks);
			workbooks = null;
		}
		if (count == 1)
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
			if (flag)
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
				grpRange.IsEnabled = true;
			}
			else
			{
				grpChart.IsEnabled = true;
			}
		}
		if (flag)
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
			chkGraphic.IsEnabled = false;
		}
		list2 = list2.Distinct().ToList();
		int count2 = list2.Count;
		checked
		{
			if (flag)
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
				if (flag2)
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
					A(A: false, B: true, C: false, D: true, E: true, F: false);
					using (List<ImportType>.Enumerator enumerator4 = list2.GetEnumerator())
					{
						while (enumerator4.MoveNext())
						{
							ImportType current2 = enumerator4.Current;
							switch (unchecked(current2 - 1))
							{
							case 10:
							case 11:
								num++;
								break;
							case 0:
							case 5:
								num2++;
								break;
							case 2:
							case 7:
								num3++;
								break;
							case 4:
							case 6:
								num4++;
								break;
							case 3:
								num5++;
								break;
							}
							if (num2 == count2)
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
								chkPicture.IsChecked = true;
							}
							else if (num == count2)
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
								chkGraphic.IsChecked = true;
							}
							else if (num3 == count2)
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
								chkEmbedded.IsChecked = true;
							}
							else if (num4 == count2)
							{
								chkChart.IsChecked = true;
							}
							else if (num5 == count2)
							{
								chkText.IsChecked = true;
							}
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0588;
							}
							continue;
							end_IL_0588:
							break;
						}
					}
					goto IL_0750;
				}
			}
		}
		if (flag)
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
			A(A: false, B: true, C: true, D: true, E: true, F: true);
			if (count2 == 1)
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
				ImportType val3 = list2[0];
				switch (val3 - 1)
				{
				case 10:
					chkGraphic.IsChecked = true;
					break;
				case 0:
					chkPicture.IsChecked = true;
					break;
				case 1:
					chkTable.IsChecked = true;
					break;
				case 2:
					chkEmbedded.IsChecked = true;
					break;
				case 4:
					chkChart.IsChecked = true;
					break;
				case 3:
					chkText.IsChecked = true;
					A(A: false, B: false, C: false, D: false, E: false, F: false);
					break;
				}
			}
		}
		else
		{
			A(A: true, B: true, C: false, D: true, E: true, F: false);
			if (count2 == 1)
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
				ImportType val4 = list2[0];
				switch (val4 - 6)
				{
				case 6:
					chkGraphic.IsChecked = true;
					break;
				case 0:
					chkPicture.IsChecked = true;
					break;
				case 2:
					chkEmbedded.IsChecked = true;
					break;
				case 1:
					chkChart.IsChecked = true;
					break;
				}
			}
		}
		goto IL_0750;
		IL_0750:
		this.m_A = false;
		if (list.Count == 1)
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
			cbxFiles.SelectedIndex = 0;
		}
		if (this.m_A == null)
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
			new ComAwareEventInfo(typeof(AppEvents_Event), XC.A(11865)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
			return;
		}
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

	private void A(bool A, bool B, bool C, bool D, bool E, bool F)
	{
		chkGraphic.IsEnabled = B;
		chkPicture.IsEnabled = B;
		chkTable.IsEnabled = C;
		chkEmbedded.IsEnabled = D;
		chkChart.IsEnabled = E;
		chkText.IsEnabled = F;
	}

	private void wpfLinkEdit_Loaded(object sender, RoutedEventArgs e)
	{
		base.MinHeight = base.ActualHeight;
		base.MaxHeight = base.ActualHeight;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void wpfLinkEdit_Closing(object sender, CancelEventArgs e)
	{
		//IL_011b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0679: Unknown result type (might be due to invalid IL or missing references)
		//IL_0681: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_01fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0203: Unknown result type (might be due to invalid IL or missing references)
		//IL_0237: Unknown result type (might be due to invalid IL or missing references)
		//IL_031b: Unknown result type (might be due to invalid IL or missing references)
		//IL_027b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0369: Unknown result type (might be due to invalid IL or missing references)
		//IL_0354: Unknown result type (might be due to invalid IL or missing references)
		//IL_0355: Unknown result type (might be due to invalid IL or missing references)
		//IL_035a: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_034b: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_040d: Unknown result type (might be due to invalid IL or missing references)
		//IL_041a: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_05cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_042e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0537: Unknown result type (might be due to invalid IL or missing references)
		//IL_053e: Expected O, but got Unknown
		//IL_0511: Unknown result type (might be due to invalid IL or missing references)
		if (base.DialogResult == true)
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
			base.Cursor = Cursors.Wait;
			if (this.m_A == null)
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
				try
				{
					this.m_A = A();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					try
					{
						this.m_A = this.B();
						this.m_A.Visible = true;
						this.m_B = true;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						C(XC.A(11732));
						e.Cancel = true;
						ProjectData.ClearProjectError();
						goto IL_06b1;
					}
					ProjectData.ClearProjectError();
				}
			}
			Microsoft.Office.Interop.Excel.Application a = this.m_A;
			if (!a.Visible)
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
				a.Visible = true;
			}
			a.ScreenUpdating = false;
			a.EnableEvents = false;
			a = null;
			List<object> list = new List<object>();
			List<bool> list2 = new List<bool>();
			List<string> list3 = new List<string>();
			int num = checked(this.m_A.Count - 1);
			int num2 = 0;
			Link val;
			Link val2;
			CustomXMLPart part;
			while (true)
			{
				object obj;
				if (num2 <= num)
				{
					obj = null;
					val = default(Link);
					try
					{
						obj = RuntimeHelpers.GetObjectValue(this.m_A[num2]);
						Type typeFromHandle = typeof(Macabacus_Word.CustomXML);
						string memberName = XC.A(3113);
						object[] obj2 = new object[1] { obj };
						object[] array = obj2;
						bool[] obj3 = new bool[1] { true };
						bool[] array2 = obj3;
						object obj4 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj2, null, null, obj3);
						if (array2[0])
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
							obj = RuntimeHelpers.GetObjectValue(array[0]);
						}
						part = (CustomXMLPart)obj4;
						object obj5 = NewLateBinding.LateGet(null, typeof(Common), XC.A(11777), array = new object[1] { obj }, null, null, array2 = new bool[1] { true });
						if (array2[0])
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
							obj = RuntimeHelpers.GetObjectValue(array[0]);
						}
						val = ((obj5 != null) ? ((Link)obj5) : default(Link));
						val2 = default(Link);
						bool flag = Base.SourceIsRange(val);
						if (chkPicture.IsChecked == true)
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
							val2.Type = (ImportType)(flag ? 1 : 6);
						}
						else if (chkGraphic.IsChecked == true)
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
							int num3;
							if (!flag)
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
								num3 = 12;
							}
							else
							{
								num3 = 11;
							}
							val2.Type = (ImportType)num3;
						}
						else if (chkEmbedded.IsChecked == true)
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
							val2.Type = (ImportType)(flag ? 3 : 8);
						}
						else if (chkChart.IsChecked == true)
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
							int num4;
							if (!flag)
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
								num4 = 7;
							}
							else
							{
								num4 = 5;
							}
							val2.Type = (ImportType)num4;
						}
						else if (chkTable.IsChecked == true)
						{
							val2.Type = (ImportType)2;
						}
						else if (chkText.IsChecked == true)
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
							val2.Type = (ImportType)4;
						}
						else
						{
							val2.Type = val.Type;
						}
						val2.Source = modFunctionsStr.BlankTo(A(), val.Source);
						bool num5 = (flag ? cbxRanges : cbxCharts).SelectedIndex > -1;
						int num6;
						if (num5)
						{
							if (!flag)
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
								num6 = ((!modFunctionsStr.IsBlank(((ChartNameInfo)cbxCharts.SelectedItem).LinkName)) ? 1 : 0);
							}
							else
							{
								num6 = (cbxRanges.SelectedItem.ToString().StartsWith(XC.A(6385)) ? 1 : 0);
							}
						}
						else
						{
							num6 = 0;
						}
						bool flag2 = (byte)num6 != 0;
						if (!num5)
						{
							goto IL_040b;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						if (flag2)
						{
							goto IL_040b;
						}
						bool B = false;
						Workbook workbook = A(val2.Source, ref B);
						if (B)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								e.Cancel = true;
								workbook = null;
								break;
							}
							break;
						}
						if (flag)
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
							if (Operators.CompareString(cbxRanges.SelectedItem.ToString(), this.m_A, TextCompare: false) == 0)
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
								Microsoft.Office.Interop.Excel.Range rangeSelection = workbook.Windows[1].RangeSelection;
								Worksheet worksheet = rangeSelection.Worksheet;
								val2.Name = Base.SourceRangeId(rangeSelection).Name;
								val2.ParentId = Base.SourceSheetId(worksheet).Name;
								worksheet = null;
								rangeSelection = null;
							}
							else
							{
								val2.Name = cbxRanges.SelectedItem.ToString();
								val2.ParentId = Edit.RangeNameToParentId(workbook, val2.Name);
							}
							goto IL_05bb;
						}
						ChartNameInfo val3 = (ChartNameInfo)cbxCharts.SelectedItem;
						Microsoft.Office.Interop.Excel.Chart chart = A(workbook);
						if (chart == null)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								C(XC.A(11800));
								e.Cancel = true;
								break;
							}
							break;
						}
						string name;
						if (!val3.IsSelectedChart)
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
							name = val3.NameForLink();
						}
						else
						{
							name = Edit.ChartToLinkName(chart);
						}
						val2.Name = name;
						val2.ParentId = Edit.GetChartParentId(chart);
						chart = null;
						val3 = null;
						goto IL_05bb;
						IL_05bd:
						CustomXML.UpdatePart(part, null, val2.Source, val2, val.Address, val.LastUpdate, this.m_A.Application.UserName, val2.ParentId);
						list2.Add(item: false);
						goto IL_061f;
						IL_040b:
						val2.Name = val.Name;
						val2.ParentId = val.ParentId;
						goto IL_05bd;
						IL_05bb:
						workbook = null;
						goto IL_05bd;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						list2.Add(item: true);
						list3.Add(ex6.Message);
						ProjectData.ClearProjectError();
						goto IL_061f;
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				this.m_A.Objects = list;
				this.m_A.IsError = list2;
				this.m_A.Errors = list3;
				break;
				IL_061f:
				list.Add(RuntimeHelpers.GetObjectValue(obj));
				num2 = checked(num2 + 1);
			}
			list = null;
			list2 = null;
			list3 = null;
			val = default(Link);
			val2 = default(Link);
			part = null;
			if (!this.m_B)
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
				Microsoft.Office.Interop.Excel.Application a2 = this.m_A;
				a2.ScreenUpdating = true;
				a2.EnableEvents = true;
				_ = null;
			}
			goto IL_06b1;
		}
		goto IL_06be;
		IL_06b1:
		base.Cursor = Cursors.Arrow;
		goto IL_06be;
		IL_06be:
		if (e.Cancel)
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
			Microsoft.Office.Interop.Word.Windows windows = this.m_A.Windows;
			object Index = 1;
			windows[ref Index].Activate();
			if (this.m_A != null)
			{
				try
				{
					this.m_A.DisplayAlerts = false;
					foreach (Workbook item in this.m_A)
					{
						try
						{
							item.Saved = true;
							item.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
					}
					Workbook workbook = null;
					this.m_A.DisplayAlerts = true;
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					ProjectData.ClearProjectError();
				}
				new ComAwareEventInfo(typeof(AppEvents_Event), XC.A(11865)).RemoveEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
				if (this.m_B)
				{
					this.m_A.Quit();
				}
				MC.A(this.m_A);
				this.m_A = null;
			}
			Properties.EditLinksWidth = base.Width;
			this.m_A = null;
			this.m_A = null;
			this.m_A = null;
			return;
		}
	}

	private Workbook A(string A, ref bool B)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		B = false;
		Workbook result;
		try
		{
			Workbook workbook = this.m_A[Path.GetFileName(A)].Workbook;
			if (workbook != null)
			{
				result = workbook;
				goto IL_030e;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Workbook workbook = null;
		}
		Workbooks workbooks = this.m_A.Workbooks;
		try
		{
			Workbook workbook = workbooks[Path.GetFileName(A)];
			if (workbook == null)
			{
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
					throw new Exception();
				}
			}
			if (object.Equals(workbook.FullName, A))
			{
				result = workbook;
			}
			else if (!workbook.Saved)
			{
				MessageBoxResult messageBoxResult = MessageBox.Show(XC.A(11890) + workbook.Name + XC.A(11969), XC.A(2438), MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);
				if (messageBoxResult != MessageBoxResult.Cancel)
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
					if (messageBoxResult == MessageBoxResult.Yes)
					{
						workbook.Save();
					}
					goto IL_0141;
				}
				B = true;
				result = workbook;
			}
			else
			{
				if (MessageBox.Show(XC.A(11890) + workbook.Name + XC.A(12219), XC.A(2438), MessageBoxButton.OKCancel, MessageBoxImage.Exclamation) != MessageBoxResult.Cancel)
				{
					goto IL_0141;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					B = true;
					result = workbook;
					break;
				}
			}
			goto end_IL_0045;
			IL_0141:
			this.m_A.DisplayAlerts = false;
			workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			this.m_A.DisplayAlerts = true;
			result = workbooks.Open(A, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			end_IL_0045:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				Workbook workbook = workbooks.Open(A, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				this.m_A.Add(workbook);
				result = workbook;
				ProjectData.ClearProjectError();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				throw new Exception(XC.A(12363));
			}
		}
		finally
		{
			MC.A(workbooks);
			workbooks = null;
			Workbook workbook = null;
		}
		goto IL_030e;
		IL_030e:
		return result;
	}

	private void A(Workbook A)
	{
		string fullName = A.FullName;
		if (this.m_A.ContainsKey(fullName))
		{
			B(fullName, A);
		}
		else
		{
			this.A(fullName, A);
		}
	}

	private void cbxFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxFiles.SelectedIndex > -1)
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
					this.m_A = cbxFiles.SelectedIndex;
					btnApply.Visibility = System.Windows.Visibility.Collapsed;
					B();
					return;
				}
			}
		}
		Ranges.Clear();
		Charts.Clear();
		btnApply.Visibility = System.Windows.Visibility.Visible;
	}

	private void cbxFiles_PreviewKeyDown(object sender, KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Return)
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
					if (key == Key.Escape)
					{
						cbxFiles.SelectedIndex = this.m_A;
						e.Handled = true;
					}
					return;
				}
			}
		}
		if (cbxFiles.SelectedIndex == -1)
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
			B();
		}
		e.Handled = true;
	}

	private void btnApply_Click(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void btnBrowse_Click(object sender, RoutedEventArgs e)
	{
		//IL_00f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_010f: Unknown result type (might be due to invalid IL or missing references)
		FileDialog fileDialog = ((Microsoft.Office.Interop.Word._Application)PC.A.Application).get_FileDialog(MsoFileDialogType.msoFileDialogOpen);
		fileDialog.Title = XC.A(12416);
		fileDialog.Filters.Clear();
		fileDialog.Filters.Add(XC.A(12447), XC.A(12476), RuntimeHelpers.GetObjectValue(Missing.Value));
		fileDialog.AllowMultiSelect = false;
		fileDialog.ButtonName = XC.A(12515);
		fileDialog.Show();
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		_ = null;
		if (selectedItems.Count == 1)
		{
			string text = Conversions.ToString(selectedItems.Cast<object>().ElementAtOrDefault(0));
			cbxFiles.Text = text;
			cbxFiles.Focus();
			if (!this.m_A.ContainsKey(text))
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
				WorkbookStruct value = new WorkbookStruct
				{
					FullName = text,
					Workbook = null
				};
				this.m_A.Add(text, value);
			}
			B();
		}
		selectedItems = null;
	}

	private string A()
	{
		ComboBox comboBox = cbxFiles;
		if (comboBox.SelectedIndex > -1)
		{
			return comboBox.SelectedItem.ToString();
		}
		return comboBox.Text;
	}

	private void B()
	{
		//IL_05dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0127: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0339: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_035e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0117: Unknown result type (might be due to invalid IL or missing references)
		//IL_011c: Unknown result type (might be due to invalid IL or missing references)
		//IL_010d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0113: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c0: Unknown result type (might be due to invalid IL or missing references)
		if (this.m_A)
		{
			return;
		}
		Link val2 = default(Link);
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
			Workbook workbook = null;
			Name name = null;
			bool flag = true;
			bool flag2 = false;
			this.m_A = true;
			Ranges.Clear();
			Charts.Clear();
			base.Cursor = Cursors.Wait;
			base.IsEnabled = false;
			if (this.m_A != null)
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
				if (this.m_A.Count == 1)
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
					Type typeFromHandle = typeof(Common);
					string memberName = XC.A(11777);
					List<object> a;
					object[] obj = new object[1] { (a = this.m_A)[0] };
					object[] array = obj;
					bool[] obj2 = new bool[1] { true };
					bool[] array2 = obj2;
					object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
					if (array2[0])
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
						a[0] = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
					}
					_003F val;
					if (obj3 == null)
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
						val = default(Link);
					}
					else
					{
						val = (Link)obj3;
					}
					val2 = (Link)val;
					flag2 = true;
				}
			}
			if (flag2)
			{
				bool flag3 = Base.SourceIsRange(val2);
				try
				{
					_ = this.m_A.Name;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					this.m_A = B();
					this.m_B = true;
					ProjectData.ClearProjectError();
				}
				this.m_A.EnableEvents = false;
				this.m_A.ScreenUpdating = false;
				try
				{
					workbook = A();
					if (workbook == null)
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
						string text = A();
						Workbooks workbooks = this.m_A.Workbooks;
						try
						{
							workbook = workbooks[Path.GetFileName(text)];
							if (workbook == null)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									throw new Exception();
								}
							}
							if (Operators.CompareString(workbook.FullName, text, TextCompare: false) != 0)
							{
								workbook = null;
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							try
							{
								try
								{
									workbook = this.m_A.Workbooks.Open(text, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									workbook.Windows[1].Visible = false;
									this.m_A.Add(workbook);
									Interaction.AppActivate(PC.A.Application.Caption);
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
							ProjectData.ClearProjectError();
						}
						MC.A(workbooks);
						workbooks = null;
					}
					if (workbook != null)
					{
						if (flag3)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								List<string> list = new List<string>();
								if (val2.Name.StartsWith(Base.LINK_PREFIX))
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
									try
									{
										name = workbook.Names.Item(val2.Name, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										if (name == null)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												throw new Exception();
											}
										}
										list.Add(name.RefersTo.ToString());
									}
									catch (Exception ex9)
									{
										ProjectData.SetProjectError(ex9);
										Exception ex10 = ex9;
										flag = false;
										ProjectData.ClearProjectError();
									}
									finally
									{
										name = null;
									}
								}
								else
								{
									list.Add(val2.Name);
								}
								if (workbook.Windows[1].Visible)
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
									list.Add(this.m_A);
								}
								list.AddRange((IEnumerable<string>)Edit.GetWorkbookNames(workbook, NC.A.ShowAllNames));
								if (!list.Any())
								{
									break;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									using (List<string>.Enumerator enumerator = list.GetEnumerator())
									{
										while (enumerator.MoveNext())
										{
											string current = enumerator.Current;
											Ranges.Add(current);
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_046c;
											}
											continue;
											end_IL_046c:
											break;
										}
									}
									if (!flag)
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
										if (Operators.CompareString(Ranges[0].ToString(), this.m_A, TextCompare: false) != 0 || Operators.CompareString(workbook.FullName, val2.Source, TextCompare: false) == 0)
										{
											break;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											break;
										}
									}
									cbxRanges.SelectedIndex = 0;
									break;
								}
								break;
							}
						}
						else
						{
							List<ChartNameInfo> availableCharts = Edit.GetAvailableCharts(val2, workbook);
							if (availableCharts.Count > 0)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									using (List<ChartNameInfo>.Enumerator enumerator2 = availableCharts.GetEnumerator())
									{
										while (enumerator2.MoveNext())
										{
											ChartNameInfo current2 = enumerator2.Current;
											Charts.Add(current2);
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_0541;
											}
											continue;
											end_IL_0541:
											break;
										}
									}
									cbxCharts.SelectedIndex = 0;
									break;
								}
							}
						}
					}
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
				finally
				{
					List<string> list = null;
					List<ChartNameInfo> availableCharts = null;
				}
				this.m_A.EnableEvents = true;
				this.m_A.ScreenUpdating = true;
			}
			this.m_A = false;
			cbxRanges.IsEnabled = cbxRanges.Items.Count > 1;
			base.IsEnabled = true;
			base.Cursor = Cursors.Arrow;
			val2 = default(Link);
			workbook = null;
			return;
		}
	}

	private void btnViewRange_Click(object sender, RoutedEventArgs e)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_014a: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		Microsoft.Office.Interop.Excel.Range range = null;
		Link val = default(Link);
		if (this.m_A == null)
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
			if (this.m_A.Count != 1)
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
				Type typeFromHandle = typeof(Common);
				string memberName = XC.A(11777);
				List<object> a;
				object[] obj = new object[1] { (a = this.m_A)[0] };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
				if (array2[0])
				{
					a[0] = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
				}
				val = ((obj3 != null) ? ((Link)obj3) : default(Link));
				if (this.m_A == null)
				{
					this.m_A = B();
					new ComAwareEventInfo(typeof(AppEvents_Event), XC.A(11865)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
					this.m_B = true;
				}
				workbook = A();
				if (workbook == null)
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
					workbook = A(this.m_A.Workbooks);
				}
				if (workbook == null)
				{
					return;
				}
				range = A(workbook, val);
				Worksheet worksheet;
				if (range != null)
				{
					this.m_A.ScreenUpdating = false;
					try
					{
						worksheet = range.Worksheet;
						A(workbook, worksheet);
						A(range);
						Environment.MakeExcelVisible(this.m_A);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						C(ex2.Message);
						ProjectData.ClearProjectError();
					}
					this.m_A.ScreenUpdating = true;
				}
				else
				{
					C(XC.A(12528));
				}
				range = null;
				worksheet = null;
				workbook = null;
				return;
			}
		}
	}

	private void btnViewChart_Click(object sender, RoutedEventArgs e)
	{
		//IL_00c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		if (this.m_A != null)
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
			if (this.m_A.Count != 1)
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
				break;
			}
			object[] array;
			List<object> a;
			bool[] array2;
			object obj = NewLateBinding.LateGet(null, typeof(Common), XC.A(11777), array = new object[1] { (a = this.m_A)[0] }, null, null, array2 = new bool[1] { true });
			if (array2[0])
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
				a[0] = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
			}
			_003F val;
			if (obj == null)
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
				val = default(Link);
			}
			else
			{
				val = (Link)obj;
			}
		}
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
			this.m_A = B();
			new ComAwareEventInfo(typeof(AppEvents_Event), XC.A(11865)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
			this.m_B = true;
		}
		Workbook workbook2 = A();
		if (workbook2 == null)
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
			workbook2 = A(this.m_A.Workbooks);
		}
		workbook = workbook2;
		if (workbook == null)
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
			Microsoft.Office.Interop.Excel.Chart chart = null;
			try
			{
				chart = A(workbook);
				if (chart != null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						A(workbook, (Worksheet)null);
						Charts.GoToChart(chart, this.m_A);
						Environment.MakeExcelVisible(this.m_A);
						break;
					}
				}
				else
				{
					C(XC.A(12599));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(string.Format(XC.A(12668), ex2.Message));
				ProjectData.ClearProjectError();
			}
			finally
			{
				MC.A(chart);
				chart = null;
			}
			workbook = null;
			return;
		}
	}

	private Microsoft.Office.Interop.Excel.Range A(Workbook A, Link B)
	{
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Excel.Range result;
		if (object.Equals(cbxRanges.SelectedItem.ToString(), this.m_A))
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
			try
			{
				result = A.Windows[1].RangeSelection;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = null;
				ProjectData.ClearProjectError();
			}
			finally
			{
			}
		}
		else
		{
			try
			{
				result = A.Names.Item(B.Name, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).RefersToRange;
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				result = null;
				ProjectData.ClearProjectError();
			}
			finally
			{
			}
		}
		return result;
	}

	private Microsoft.Office.Interop.Excel.Chart A(Workbook A)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		ChartNameInfo val = (ChartNameInfo)cbxCharts.SelectedItem;
		if (val.IsSelectedChart)
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
					return A.ActiveChart;
				}
			}
		}
		return Charts.FindChartByName(A, val.NameForLink(), val.WsName, val.LinkParentId);
	}

	private void A(Workbook A, Worksheet B)
	{
		Workbooks.EnsureVisible(A);
		if (B != null)
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
			Worksheets.EnsureVisible(B);
		}
		A.Activate();
		if (B == null)
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
			B.Activate();
			return;
		}
	}

	private void A(Microsoft.Office.Interop.Excel.Range A)
	{
		try
		{
			Ranges.ScrollIntoView(A);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		A.Select();
	}

	private void A(ChartObject A)
	{
		if (A == null)
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
			try
			{
				Microsoft.Office.Interop.Excel.ChartArea chartArea = A.Chart.ChartArea;
				try
				{
					this.m_A.Goto(A.TopLeftCell, true);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				chartArea.Select();
				return;
			}
			finally
			{
				Microsoft.Office.Interop.Excel.ChartArea chartArea = null;
			}
		}
	}

	private void A(string A, Workbook B)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		WorkbookStruct value = new WorkbookStruct
		{
			FullName = A,
			Workbook = B
		};
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			cbxFiles.Items.Add(A);
		});
		this.m_A.Add(A, value);
		value = default(WorkbookStruct);
	}

	private void B(string A, Workbook B)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		this.m_A[A] = Base.WorkbookProperties(B);
	}

	private Workbook A()
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		try
		{
			workbook = this.m_A[A()].Workbook;
			if (workbook == null)
			{
				return null;
			}
			_ = workbook.Name;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			workbook = null;
			ProjectData.ClearProjectError();
		}
		return workbook;
	}

	private Workbook A(Workbooks A)
	{
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		string text = this.A();
		try
		{
			workbook = A.Open(text, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			this.m_A[text] = Base.WorkbookProperties(workbook);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(XC.A(12747));
			ProjectData.ClearProjectError();
		}
		return workbook;
	}

	private void B(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private Microsoft.Office.Interop.Excel.Application A()
	{
		return (Microsoft.Office.Interop.Excel.Application)Interaction.GetObject(null, XC.A(12824));
	}

	private Microsoft.Office.Interop.Excel.Application B()
	{
		return (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject(XC.A(12824));
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(XC.A(12859), UriKind.Relative);
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
			cbxFiles = (ComboBox)target;
			return;
		}
		if (connectionId == 2)
		{
			btnApply = (Button)target;
			return;
		}
		if (connectionId == 3)
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
					btnBrowse = (Button)target;
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
					grpRange = (GroupBox)target;
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
					cbxRanges = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnViewRange = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					grpChart = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxCharts = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnViewChart = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkGraphic = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkPicture = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			chkTable = (RadioButton)target;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkEmbedded = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkChart = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkText = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
					return;
				}
			}
		}
		this.m_C = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
