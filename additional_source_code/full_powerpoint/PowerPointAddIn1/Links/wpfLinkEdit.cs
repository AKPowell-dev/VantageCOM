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
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

[DesignerGenerated]
public sealed class wpfLinkEdit : System.Windows.Window, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class XE
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

	private Shapes.EditedShapes m_A;

	private bool m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

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

	[CompilerGenerated]
	[AccessedThroughProperty("btnApply")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBrowse")]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("grpRange")]
	private GroupBox m_A;

	[AccessedThroughProperty("cbxRanges")]
	[CompilerGenerated]
	private ComboBox m_B;

	[AccessedThroughProperty("btnViewRange")]
	[CompilerGenerated]
	private Button m_C;

	[AccessedThroughProperty("grpChart")]
	[CompilerGenerated]
	private GroupBox m_B;

	[AccessedThroughProperty("cbxCharts")]
	[CompilerGenerated]
	private ComboBox m_C;

	[AccessedThroughProperty("btnViewChart")]
	[CompilerGenerated]
	private Button D;

	[AccessedThroughProperty("chkGraphic")]
	[CompilerGenerated]
	private RadioButton m_A;

	[AccessedThroughProperty("chkPicture")]
	[CompilerGenerated]
	private RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTable")]
	private RadioButton m_C;

	[AccessedThroughProperty("chkEmbedded")]
	[CompilerGenerated]
	private RadioButton D;

	[AccessedThroughProperty("chkChart")]
	[CompilerGenerated]
	private RadioButton E;

	[AccessedThroughProperty("chkText")]
	[CompilerGenerated]
	private RadioButton F;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button F;

	private bool m_C;

	public Shapes.EditedShapes ReturnValue
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
			A(AH.A(92193));
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
			A(AH.A(67991));
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
				switch (6)
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
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
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
			if (button != null)
			{
				button.Click += value2;
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
				button.Click -= value2;
			}
			this.m_C = value;
			button = this.m_C;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
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
			if (button == null)
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
				button.Click += value2;
				return;
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
				switch (3)
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
				return;
			}
		}
	}

	public wpfLinkEdit(List<object> listShapes)
	{
		//IL_027e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0283: Unknown result type (might be due to invalid IL or missing references)
		//IL_0285: Unknown result type (might be due to invalid IL or missing references)
		//IL_0288: Unknown result type (might be due to invalid IL or missing references)
		//IL_0295: Unknown result type (might be due to invalid IL or missing references)
		//IL_0297: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0170: Unknown result type (might be due to invalid IL or missing references)
		//IL_0173: Unknown result type (might be due to invalid IL or missing references)
		//IL_0180: Unknown result type (might be due to invalid IL or missing references)
		//IL_0182: Unknown result type (might be due to invalid IL or missing references)
		//IL_018c: Unknown result type (might be due to invalid IL or missing references)
		//IL_018e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0193: Unknown result type (might be due to invalid IL or missing references)
		//IL_0195: Unknown result type (might be due to invalid IL or missing references)
		//IL_0198: Unknown result type (might be due to invalid IL or missing references)
		//IL_019a: Invalid comparison between Unknown and I4
		//IL_06ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_06d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_06da: Unknown result type (might be due to invalid IL or missing references)
		//IL_06fc: Expected I4, but got Unknown
		//IL_0238: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_05da: Unknown result type (might be due to invalid IL or missing references)
		//IL_05dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_05df: Unknown result type (might be due to invalid IL or missing references)
		//IL_0611: Expected I4, but got Unknown
		//IL_023d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0240: Unknown result type (might be due to invalid IL or missing references)
		//IL_024d: Unknown result type (might be due to invalid IL or missing references)
		//IL_024f: Unknown result type (might be due to invalid IL or missing references)
		//IL_022e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0234: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01aa: Invalid comparison between Unknown and I4
		//IL_0488: Unknown result type (might be due to invalid IL or missing references)
		//IL_048d: Unknown result type (might be due to invalid IL or missing references)
		//IL_048f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0492: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c8: Expected I4, but got Unknown
		base.Loaded += wpfLinkEdit_Loaded;
		base.Closing += wpfLinkEdit_Closing;
		this.m_A = true;
		this.m_B = false;
		this.m_A = AH.A(93449);
		this.m_B = AH.A(93478);
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
		Ranges = new ObservableCollection<string>();
		Charts = new ObservableCollection<ChartNameInfo>();
		this.m_A = listShapes;
		this.m_A = new Dictionary<string, WorkbookStruct>();
		this.m_A = NG.A.Application.ActivePresentation;
		this.m_A = new List<Workbook>();
		btnApply.Visibility = Visibility.Collapsed;
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
		foreach (object item in this.m_A)
		{
			object objectValue = RuntimeHelpers.GetObjectValue(item);
			if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
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
				if (!Shapes.IsLinked((Microsoft.Office.Interop.PowerPoint.Shape)objectValue))
				{
					continue;
				}
				Link val = Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)objectValue);
				list.Add(val.Source);
				list2.Add(val.Type);
				ImportType type = val.Type;
				if (type - 6 > 2)
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
					if ((int)type != 12)
					{
						flag = true;
						continue;
					}
				}
				flag2 = true;
			}
			else if (objectValue is TextLink)
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
				Type typeFromHandle = typeof(Text);
				string memberName = AH.A(93278);
				object[] obj = new object[1] { objectValue };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
				if (array2[0])
				{
					objectValue = RuntimeHelpers.GetObjectValue(array[0]);
				}
				_003F val2;
				if (obj3 == null)
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
					val2 = default(Link);
				}
				else
				{
					val2 = (Link)obj3;
				}
				Link val = (Link)val2;
				list.Add(val.Source);
				list2.Add(val.Type);
				flag = true;
			}
			else
			{
				if (!Hyperlinks.IsLinked((Hyperlink)objectValue))
				{
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
				Link val = Hyperlinks.LinkDetails((Hyperlink)objectValue);
				list.Add(val.Source);
				list2.Add(val.Type);
				flag = true;
			}
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
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_02fa;
				}
				continue;
				end_IL_02fa:
				break;
			}
		}
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
			Workbooks workbooks = this.m_A.Workbooks;
			try
			{
				IEnumerator enumerator3 = workbooks.GetEnumerator();
				try
				{
					while (enumerator3.MoveNext())
					{
						Workbook workbook = (Workbook)enumerator3.Current;
						if (this.m_A.ContainsKey(workbook.FullName))
						{
							B(workbook.FullName, workbook);
						}
						else
						{
							A(workbook.FullName, workbook);
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0398;
						}
						continue;
						end_IL_0398:
						break;
					}
				}
				finally
				{
					IDisposable disposable = enumerator3 as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			JG.A(workbooks);
			workbooks = null;
		}
		if (this.m_A.Count == 1)
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
					switch (4)
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
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			chkGraphic.IsEnabled = false;
		}
		list2 = list2.Distinct().ToList();
		int count = list2.Count;
		checked
		{
			if (flag)
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
				if (flag2)
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
							}
							if (num2 == count)
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
								chkPicture.IsChecked = true;
							}
							else if (num == count)
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
								chkGraphic.IsChecked = true;
							}
							else if (num3 == count)
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
								chkEmbedded.IsChecked = true;
							}
							else
							{
								if (num4 != count)
								{
									continue;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									break;
								}
								chkChart.IsChecked = true;
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0586;
							}
							continue;
							end_IL_0586:
							break;
						}
					}
					goto IL_074e;
				}
			}
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
			A(A: false, B: true, C: true, D: true, E: true, F: true);
			if (count == 1)
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
			if (count == 1)
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
		goto IL_074e;
		IL_074e:
		this.m_A = false;
		if (list.Count == 1)
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
			cbxFiles.SelectedIndex = 0;
		}
		if (this.m_A != null)
		{
			new ComAwareEventInfo(typeof(AppEvents_Event), AH.A(92319)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
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
			switch (3)
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
		chkGraphic.IsEnabled = A;
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
		//IL_012b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0130: Unknown result type (might be due to invalid IL or missing references)
		//IL_0132: Unknown result type (might be due to invalid IL or missing references)
		//IL_0135: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Unknown result type (might be due to invalid IL or missing references)
		//IL_06dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_017f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0271: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0297: Unknown result type (might be due to invalid IL or missing references)
		//IL_0207: Unknown result type (might be due to invalid IL or missing references)
		//IL_024b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0369: Unknown result type (might be due to invalid IL or missing references)
		//IL_0376: Unknown result type (might be due to invalid IL or missing references)
		//IL_030b: Unknown result type (might be due to invalid IL or missing references)
		//IL_038a: Unknown result type (might be due to invalid IL or missing references)
		//IL_05fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0611: Unknown result type (might be due to invalid IL or missing references)
		//IL_063c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0535: Unknown result type (might be due to invalid IL or missing references)
		//IL_053b: Unknown result type (might be due to invalid IL or missing references)
		//IL_054c: Unknown result type (might be due to invalid IL or missing references)
		//IL_054d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0552: Unknown result type (might be due to invalid IL or missing references)
		//IL_057a: Unknown result type (might be due to invalid IL or missing references)
		//IL_058e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0593: Unknown result type (might be due to invalid IL or missing references)
		//IL_0599: Unknown result type (might be due to invalid IL or missing references)
		//IL_05a5: Expected O, but got Unknown
		//IL_05a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_05bd: Expected O, but got Unknown
		//IL_05bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_05dc: Expected O, but got Unknown
		//IL_05de: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_05f4: Expected O, but got Unknown
		//IL_0497: Unknown result type (might be due to invalid IL or missing references)
		//IL_049e: Expected O, but got Unknown
		//IL_0471: Unknown result type (might be due to invalid IL or missing references)
		List<object> list;
		List<bool> list2;
		List<string> list3;
		Link val;
		Link val2;
		if (base.DialogResult == true)
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
			base.Cursor = Cursors.Wait;
			if (this.m_A == null)
			{
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
						C(AH.A(92206));
						e.Cancel = true;
						ProjectData.ClearProjectError();
						goto IL_0712;
					}
					ProjectData.ClearProjectError();
				}
			}
			Microsoft.Office.Interop.Excel.Application a = this.m_A;
			if (!a.Visible)
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
				a.Visible = true;
			}
			a.ScreenUpdating = false;
			a.EnableEvents = false;
			a = null;
			list = new List<object>();
			list2 = new List<bool>();
			list3 = new List<string>();
			using (List<object>.Enumerator enumerator = this.m_A.GetEnumerator())
			{
				while (true)
				{
					if (!enumerator.MoveNext())
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0693;
							}
							continue;
							end_IL_0693:
							break;
						}
						break;
					}
					object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
					val = default(Link);
					try
					{
						val = A(RuntimeHelpers.GetObjectValue(objectValue));
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
							int num;
							if (!flag)
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
								num = 6;
							}
							else
							{
								num = 1;
							}
							val2.Type = (ImportType)num;
						}
						else if (chkGraphic.IsChecked == true)
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
							int num2;
							if (!flag)
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
								num2 = 12;
							}
							else
							{
								num2 = 11;
							}
							val2.Type = (ImportType)num2;
						}
						else if (chkEmbedded.IsChecked == true)
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
								num3 = 8;
							}
							else
							{
								num3 = 3;
							}
							val2.Type = (ImportType)num3;
						}
						else if (chkChart.IsChecked == true)
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
							int num4;
							if (!flag)
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
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
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
								num6 = (cbxRanges.SelectedItem.ToString().StartsWith(AH.A(92251)) ? 1 : 0);
							}
						}
						else
						{
							num6 = 0;
						}
						bool flag2 = (byte)num6 != 0;
						if (!num5)
						{
							goto IL_0367;
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
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								break;
							}
							goto IL_0367;
						}
						bool B = false;
						Workbook workbook = A(val2.Source, ref B);
						if (B)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								e.Cancel = true;
								workbook = null;
								break;
							}
						}
						else
						{
							if (flag)
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
								if (Operators.CompareString(cbxRanges.SelectedItem.ToString(), this.m_A, TextCompare: false) == 0)
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
									Range rangeSelection = workbook.Windows[1].RangeSelection;
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
								goto IL_0518;
							}
							ChartNameInfo val3 = (ChartNameInfo)cbxCharts.SelectedItem;
							Microsoft.Office.Interop.Excel.Chart chart = A(workbook);
							if (chart != null)
							{
								val2.Name = (val3.IsSelectedChart ? Edit.ChartToLinkName(chart) : val3.NameForLink());
								val2.ParentId = Edit.GetChartParentId(chart);
								chart = null;
								val3 = null;
								goto IL_0518;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								MessageBox.Show(AH.A(92254), AH.A(5874), MessageBoxButton.OK);
								e.Cancel = true;
								break;
							}
						}
						goto end_IL_0121;
						IL_0518:
						workbook = null;
						goto IL_051a;
						IL_0367:
						val2.Name = val.Name;
						val2.ParentId = val.ParentId;
						goto IL_051a;
						IL_051a:
						if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
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
							A((Microsoft.Office.Interop.PowerPoint.Shape)objectValue, val2.Source, val2.Name, this.m_A.UserName, val2.Type, val2.ParentId);
						}
						else if (objectValue is TextLink)
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
							Text.TextRangeParentShape(((TextLink)objectValue).TextRange);
							Text.UpdateSource((TextLink)objectValue, val.RangeId, val2.Source, blnUpdateLastModified: true);
							Text.UpdateName((TextLink)objectValue, val.RangeId, val2.Name);
							Text.UpdateUser((TextLink)objectValue, val.RangeId, this.m_A.UserName);
							Text.UpdateParentId((TextLink)objectValue, val.RangeId, val2.ParentId);
						}
						else
						{
							Hyperlinks.UpdateSource((Hyperlink)objectValue, null, val2.Source, blnUpdateLastModified: true);
							Hyperlinks.UpdateName((Hyperlink)objectValue, val2.Name);
							Hyperlinks.UpdateUser((Hyperlink)objectValue, this.m_A.UserName);
							Hyperlinks.UpdateParentId((Hyperlink)objectValue, val2.ParentId);
						}
						list2.Add(item: false);
						goto IL_0678;
						end_IL_0121:;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						list2.Add(item: true);
						list3.Add(ex6.Message);
						ProjectData.ClearProjectError();
						goto IL_0678;
					}
					goto end_IL_00f5;
					IL_0678:
					list.Add(RuntimeHelpers.GetObjectValue(objectValue));
				}
				goto IL_06ad;
				end_IL_00f5:;
			}
			goto IL_06d3;
		}
		goto IL_071f;
		IL_06ad:
		this.m_A.Objects = list;
		this.m_A.IsError = list2;
		this.m_A.Errors = list3;
		goto IL_06d3;
		IL_06d3:
		list = null;
		list2 = null;
		list3 = null;
		val = default(Link);
		val2 = default(Link);
		if (!this.m_B)
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
			Microsoft.Office.Interop.Excel.Application a2 = this.m_A;
			a2.ScreenUpdating = true;
			a2.EnableEvents = true;
			_ = null;
		}
		goto IL_0712;
		IL_071f:
		if (e.Cancel)
		{
			return;
		}
		this.m_A.Windows[1].Activate();
		if (this.m_A != null)
		{
			try
			{
				this.m_A.DisplayAlerts = false;
				Workbook workbook;
				using (List<Workbook>.Enumerator enumerator2 = this.m_A.GetEnumerator())
				{
					while (enumerator2.MoveNext())
					{
						workbook = enumerator2.Current;
						try
						{
							workbook.Saved = true;
							workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_07b9;
						}
						continue;
						end_IL_07b9:
						break;
					}
				}
				workbook = null;
				this.m_A.DisplayAlerts = true;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				ProjectData.ClearProjectError();
			}
			new ComAwareEventInfo(typeof(AppEvents_Event), AH.A(92319)).RemoveEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
			if (this.m_B)
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
				this.m_A.Quit();
			}
			JG.A(this.m_A);
			this.m_A = null;
		}
		Properties.EditLinksWidth = base.Width;
		Ranges = null;
		Charts = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		return;
		IL_0712:
		base.Cursor = Cursors.Arrow;
		goto IL_071f;
	}

	private Workbook A(string A, ref bool B)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		B = false;
		Workbook result;
		try
		{
			Workbook workbook = this.m_A[Path.GetFileName(A)].Workbook;
			if (workbook != null)
			{
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
					result = workbook;
					break;
				}
				goto IL_033a;
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
					switch (7)
					{
					case 0:
						continue;
					}
					throw new Exception();
				}
			}
			if (object.Equals(workbook.FullName, A))
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					result = workbook;
					break;
				}
			}
			else if (!workbook.Saved)
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
				MessageBoxResult messageBoxResult = MessageBox.Show(AH.A(92344) + workbook.Name + AH.A(92423), AH.A(5874), MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);
				if (messageBoxResult != MessageBoxResult.Cancel)
				{
					if (messageBoxResult == MessageBoxResult.Yes)
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
						workbook.Save();
					}
					goto IL_0169;
				}
				B = true;
				result = workbook;
			}
			else
			{
				if (MessageBox.Show(AH.A(92344) + workbook.Name + AH.A(92673), AH.A(5874), MessageBoxButton.OKCancel, MessageBoxImage.Exclamation) != MessageBoxResult.Cancel)
				{
					goto IL_0169;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					B = true;
					result = workbook;
					break;
				}
			}
			goto end_IL_0058;
			IL_0169:
			this.m_A.DisplayAlerts = false;
			workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			this.m_A.DisplayAlerts = true;
			result = workbooks.Open(A, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			end_IL_0058:;
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
				throw new Exception(AH.A(92817));
			}
		}
		finally
		{
			JG.A(workbooks);
			workbooks = null;
			Workbook workbook = null;
		}
		goto IL_033a;
		IL_033a:
		return result;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, string B, string C, string D, ImportType E, string F)
	{
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		Shapes.ConvertLegacyLink(A);
		Common.UpdateSource(A.Tags, null, B, blnUpdateLastModified: true);
		Common.UpdateName(A.Tags, C);
		Common.UpdateUser(A.Tags, D);
		Common.UpdateType(A.Tags, E);
		Common.UpdateParentId(A.Tags, F);
	}

	private void A(Workbook A)
	{
		string fullName = A.FullName;
		if (this.m_A.ContainsKey(fullName))
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
					B(fullName, A);
					return;
				}
			}
		}
		this.A(fullName, A);
	}

	private void cbxFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxFiles.SelectedIndex > -1)
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
					this.m_A = cbxFiles.SelectedIndex;
					btnApply.Visibility = Visibility.Collapsed;
					B();
					return;
				}
			}
		}
		Ranges.Clear();
		Charts.Clear();
		btnApply.Visibility = Visibility.Visible;
	}

	private void cbxFiles_PreviewKeyDown(object sender, KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Return)
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
					if (key != Key.Escape)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
					cbxFiles.SelectedIndex = this.m_A;
					e.Handled = true;
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
		//IL_00f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0113: Unknown result type (might be due to invalid IL or missing references)
		FileDialog fileDialog = ((Microsoft.Office.Interop.PowerPoint._Application)NG.A.Application).get_FileDialog(MsoFileDialogType.msoFileDialogOpen);
		fileDialog.Title = AH.A(92870);
		fileDialog.Filters.Clear();
		fileDialog.Filters.Add(AH.A(92901), AH.A(92930), RuntimeHelpers.GetObjectValue(Missing.Value));
		fileDialog.AllowMultiSelect = false;
		fileDialog.ButtonName = AH.A(92969);
		fileDialog.Show();
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		_ = null;
		if (selectedItems.Count == 1)
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
			string text = Conversions.ToString(selectedItems.Cast<object>().ElementAtOrDefault(0));
			cbxFiles.Text = text;
			cbxFiles.Focus();
			if (!this.m_A.ContainsKey(text))
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
		//IL_051c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0081: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Unknown result type (might be due to invalid IL or missing references)
		//IL_0088: Unknown result type (might be due to invalid IL or missing references)
		//IL_0089: Unknown result type (might be due to invalid IL or missing references)
		//IL_0436: Unknown result type (might be due to invalid IL or missing references)
		//IL_0296: Unknown result type (might be due to invalid IL or missing references)
		//IL_0310: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0406: Unknown result type (might be due to invalid IL or missing references)
		if (this.m_A)
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
			Workbook workbook = null;
			Name name = null;
			bool flag = true;
			this.m_A = true;
			Ranges.Clear();
			Charts.Clear();
			base.Cursor = Cursors.Wait;
			base.IsEnabled = false;
			Link val;
			if (this.m_A.Count == 1)
			{
				val = A(RuntimeHelpers.GetObjectValue(this.m_A[0]));
				bool flag2 = Base.SourceIsRange(val);
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
							switch (3)
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
									switch (7)
									{
									case 0:
										continue;
									}
									throw new Exception();
								}
							}
							if (Operators.CompareString(workbook.FullName, text, TextCompare: false) != 0)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									workbook = null;
									break;
								}
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
									workbook = workbooks.Open(text, 0, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									workbook.Windows[1].Visible = false;
									this.m_A.Add(workbook);
									Interaction.AppActivate(NG.A.Application.Caption);
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
						JG.A(workbooks);
						workbooks = null;
					}
					if (workbook != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							if (flag2)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									List<string> list = new List<string>();
									if (val.Name.StartsWith(Base.LINK_PREFIX))
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
										try
										{
											name = workbook.Names.Item(val.Name, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
											if (name == null)
											{
												throw new Exception();
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
										list.Add(val.Name);
									}
									if (workbook.Windows[1].Visible)
									{
										list.Add(this.m_A);
									}
									list.AddRange((IEnumerable<string>)Edit.GetWorkbookNames(workbook, KG.A.ShowAllNames));
									if (!list.Any())
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
										using (List<string>.Enumerator enumerator = list.GetEnumerator())
										{
											while (enumerator.MoveNext())
											{
												string current = enumerator.Current;
												Ranges.Add(current);
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_03ae;
												}
												continue;
												end_IL_03ae:
												break;
											}
										}
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
											if (Operators.CompareString(Ranges[0].ToString(), this.m_A, TextCompare: false) != 0 || Operators.CompareString(workbook.FullName, val.Source, TextCompare: false) == 0)
											{
												break;
											}
											while (true)
											{
												switch (2)
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
								break;
							}
							List<ChartNameInfo> availableCharts = Edit.GetAvailableCharts(val, workbook);
							if (availableCharts.Count <= 0)
							{
								break;
							}
							while (true)
							{
								switch (3)
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
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_0485;
										}
										continue;
										end_IL_0485:
										break;
									}
								}
								cbxCharts.SelectedIndex = 0;
								break;
							}
							break;
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
			val = default(Link);
			workbook = null;
			return;
		}
	}

	private void btnViewRange_Click(object sender, RoutedEventArgs e)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		Range range = null;
		Link val = default(Link);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			val = A(RuntimeHelpers.GetObjectValue(this.m_A[0]));
			if (this.m_A == null)
			{
				this.m_A = B();
				new ComAwareEventInfo(typeof(AppEvents_Event), AH.A(92319)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
				this.m_B = true;
			}
			workbook = A();
			if (workbook == null)
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
				workbook = A(this.m_A.Workbooks);
			}
			if (workbook == null)
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
					C(AH.A(92982));
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
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		if (this.m_A.Count != 1)
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
			A(RuntimeHelpers.GetObjectValue(this.m_A[0]));
			if (this.m_A == null)
			{
				this.m_A = B();
				new ComAwareEventInfo(typeof(AppEvents_Event), AH.A(92319)).AddEventHandler(this.m_A, new AppEvents_WorkbookOpenEventHandler(A));
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
				switch (2)
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
							switch (6)
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
						C(AH.A(93053));
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					C(string.Format(AH.A(93122), ex2.Message));
					ProjectData.ClearProjectError();
				}
				finally
				{
					JG.A(chart);
					chart = null;
				}
				workbook = null;
				return;
			}
		}
	}

	private Range A(Workbook A, Link B)
	{
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		Range result;
		if (object.Equals(cbxRanges.SelectedItem.ToString(), this.m_A))
		{
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
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ChartNameInfo val = (ChartNameInfo)cbxCharts.SelectedItem;
		if (val.IsSelectedChart)
		{
			return A.ActiveChart;
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
			Worksheets.EnsureVisible(B);
		}
		A.Activate();
		if (B == null)
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
			B.Activate();
			return;
		}
	}

	private void A(Range A)
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
		}
		finally
		{
			Microsoft.Office.Interop.Excel.ChartArea chartArea = null;
		}
	}

	private void A(string A, Workbook B)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
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
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		Workbook workbook = null;
		try
		{
			workbook = this.m_A[A()].Workbook;
			if (workbook == null)
			{
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
					return null;
				}
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
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
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
			C(AH.A(93201));
			ProjectData.ClearProjectError();
		}
		return workbook;
	}

	private Link A(object A)
	{
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00af: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		if (A is Microsoft.Office.Interop.PowerPoint.Shape)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)A);
				}
			}
		}
		if (A is TextLink)
		{
			Type typeFromHandle = typeof(Text);
			string memberName = AH.A(93278);
			object[] obj = new object[1] { A };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
			if (array2[0])
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
				A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			return (obj3 != null) ? ((Link)obj3) : default(Link);
		}
		return Hyperlinks.LinkDetails((Hyperlink)A);
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
		return (Microsoft.Office.Interop.Excel.Application)Interaction.GetObject(null, AH.A(93301));
	}

	private Microsoft.Office.Interop.Excel.Application B()
	{
		return (Microsoft.Office.Interop.Excel.Application)Interaction.CreateObject(AH.A(93301));
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!this.m_C)
		{
			this.m_C = true;
			Uri resourceLocator = new Uri(AH.A(93336), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					cbxFiles = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnApply = (Button)target;
					return;
				}
			}
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
					btnBrowse = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (4)
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
			cbxRanges = (ComboBox)target;
			return;
		}
		if (connectionId == 6)
		{
			btnViewRange = (Button)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (7)
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
				switch (4)
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
				switch (7)
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
				switch (5)
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
				switch (2)
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
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkTable = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (5)
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
		switch (connectionId)
		{
		case 15:
			chkText = (RadioButton)target;
			break;
		case 16:
			btnOk = (Button)target;
			break;
		case 17:
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				btnCancel = (Button)target;
				return;
			}
		default:
			this.m_C = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
