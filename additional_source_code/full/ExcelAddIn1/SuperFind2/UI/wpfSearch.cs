using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using ExcelAddIn1.SuperFind2.Callbacks;
using ExcelAddIn1.SuperFind2.Queries;
using ExcelAddIn1.SuperFind2.Results;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.UI;

[DesignerGenerated]
public sealed class wpfSearch : UserControl, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ObservableCollection<BaseItem> m_A;

	[CompilerGenerated]
	private wpfPane m_A;

	[AccessedThroughProperty("chkText")]
	[CompilerGenerated]
	private CheckBox m_A;

	[AccessedThroughProperty("cbxTextFilter")]
	[CompilerGenerated]
	private ComboBox m_A;

	[AccessedThroughProperty("grdTextArgs")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("txtText1")]
	[CompilerGenerated]
	private TextBox m_A;

	[AccessedThroughProperty("txtText2")]
	[CompilerGenerated]
	private TextBox m_B;

	[AccessedThroughProperty("chkMatchCase")]
	[CompilerGenerated]
	private CheckBox m_B;

	[AccessedThroughProperty("chkLookInValues")]
	[CompilerGenerated]
	private CheckBox m_C;

	[AccessedThroughProperty("chkLookInFormulas")]
	[CompilerGenerated]
	private CheckBox m_D;

	[AccessedThroughProperty("chkLookInComments")]
	[CompilerGenerated]
	private CheckBox E;

	[AccessedThroughProperty("chkLookInCharts")]
	[CompilerGenerated]
	private CheckBox F;

	[AccessedThroughProperty("chkLookInHyperlinks")]
	[CompilerGenerated]
	private CheckBox G;

	[AccessedThroughProperty("chkValues")]
	[CompilerGenerated]
	private CheckBox H;

	[AccessedThroughProperty("cbxValuesFilter")]
	[CompilerGenerated]
	private ComboBox m_B;

	[AccessedThroughProperty("grdValuesArgs")]
	[CompilerGenerated]
	private UniformGrid m_A;

	[AccessedThroughProperty("txtValues1")]
	[CompilerGenerated]
	private TextBox m_C;

	[AccessedThroughProperty("txtValues2")]
	[CompilerGenerated]
	private TextBox m_D;

	[AccessedThroughProperty("chkDates")]
	[CompilerGenerated]
	private CheckBox I;

	[AccessedThroughProperty("cbxDatesFilter")]
	[CompilerGenerated]
	private ComboBox m_C;

	[AccessedThroughProperty("grdDatesArgs")]
	[CompilerGenerated]
	private UniformGrid m_B;

	[AccessedThroughProperty("txtDates1")]
	[CompilerGenerated]
	private TextBox E;

	[AccessedThroughProperty("txtDates2")]
	[CompilerGenerated]
	private TextBox F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFormulas")]
	private CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFormulasFilter")]
	private ComboBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("grdFormulasArgs")]
	private UniformGrid m_C;

	[AccessedThroughProperty("txtFormulas1")]
	[CompilerGenerated]
	private TextBox G;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFormulas2")]
	private TextBox H;

	[AccessedThroughProperty("chkFormats")]
	[CompilerGenerated]
	private CheckBox K;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFormatsFilter")]
	private ComboBox E;

	[AccessedThroughProperty("grdFormatsArgs")]
	[CompilerGenerated]
	private UniformGrid m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFormats1")]
	private TextBox I;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFormats2")]
	private TextBox J;

	[AccessedThroughProperty("chkLookInEmptyCells")]
	[CompilerGenerated]
	private CheckBox L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkData")]
	private CheckBox M;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxDataFilter")]
	private ComboBox F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRanges")]
	private CheckBox N;

	[AccessedThroughProperty("cbxRangesFilter")]
	[CompilerGenerated]
	private ComboBox G;

	[AccessedThroughProperty("chkComments")]
	[CompilerGenerated]
	private CheckBox O;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxCommentsFilter")]
	private ComboBox H;

	[CompilerGenerated]
	[AccessedThroughProperty("grdCommentsArgs")]
	private UniformGrid E;

	[AccessedThroughProperty("txtComments1")]
	[CompilerGenerated]
	private TextBox K;

	[CompilerGenerated]
	[AccessedThroughProperty("txtComments2")]
	private TextBox L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkNotes")]
	private CheckBox P;

	[AccessedThroughProperty("chkNumericInputs")]
	[CompilerGenerated]
	private CheckBox Q;

	[AccessedThroughProperty("chkCharts")]
	[CompilerGenerated]
	private CheckBox R;

	[AccessedThroughProperty("chkSparklines")]
	[CompilerGenerated]
	private CheckBox S;

	[AccessedThroughProperty("chkShapes")]
	[CompilerGenerated]
	private CheckBox T;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHyperlinks")]
	private CheckBox U;

	[CompilerGenerated]
	[AccessedThroughProperty("chkExploreMode")]
	private CheckBox V;

	[AccessedThroughProperty("cbxScope")]
	[CompilerGenerated]
	private ComboBox I;

	[AccessedThroughProperty("chkLookInPrintAreas")]
	[CompilerGenerated]
	private CheckBox W;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFind")]
	private Button m_A;

	private bool m_A;

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(36261));
		}
	}

	private ObservableCollection<BaseItem> SearchResults
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	private wpfPane ParentView
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual CheckBox chkText
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

	internal virtual ComboBox cbxTextFilter
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
			SelectionChangedEventHandler value2 = TextFilterChanged;
			ComboBox comboBox = this.m_A;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
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
				return;
			}
		}
	}

	internal virtual Grid grdTextArgs
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

	internal virtual TextBox txtText1
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

	internal virtual TextBox txtText2
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

	internal virtual CheckBox chkMatchCase
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

	internal virtual CheckBox chkLookInValues
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

	internal virtual CheckBox chkLookInFormulas
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual CheckBox chkLookInComments
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

	internal virtual CheckBox chkLookInCharts
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

	internal virtual CheckBox chkLookInHyperlinks
	{
		[CompilerGenerated]
		get
		{
			return this.G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.G = value;
		}
	}

	internal virtual CheckBox chkValues
	{
		[CompilerGenerated]
		get
		{
			return this.H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.H = value;
		}
	}

	internal virtual ComboBox cbxValuesFilter
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
			SelectionChangedEventHandler value2 = ValuesFilterChanged;
			ComboBox comboBox = this.m_B;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
			}
			this.m_B = value;
			comboBox = this.m_B;
			if (comboBox == null)
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
				comboBox.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual UniformGrid grdValuesArgs
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

	internal virtual TextBox txtValues1
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

	internal virtual TextBox txtValues2
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual CheckBox chkDates
	{
		[CompilerGenerated]
		get
		{
			return this.I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.I = value;
		}
	}

	internal virtual ComboBox cbxDatesFilter
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
			SelectionChangedEventHandler value2 = DatesFilterChanged;
			ComboBox comboBox = this.m_C;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
			}
			this.m_C = value;
			comboBox = this.m_C;
			if (comboBox == null)
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
				comboBox.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual UniformGrid grdDatesArgs
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

	internal virtual TextBox txtDates1
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

	internal virtual TextBox txtDates2
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

	internal virtual CheckBox chkFormulas
	{
		[CompilerGenerated]
		get
		{
			return this.J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.J = value;
		}
	}

	internal virtual ComboBox cbxFormulasFilter
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			SelectionChangedEventHandler value2 = FormulasFilterChanged;
			ComboBox comboBox = this.m_D;
			if (comboBox != null)
			{
				comboBox.SelectionChanged -= value2;
			}
			this.m_D = value;
			comboBox = this.m_D;
			if (comboBox == null)
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
				comboBox.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual UniformGrid grdFormulasArgs
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

	internal virtual TextBox txtFormulas1
	{
		[CompilerGenerated]
		get
		{
			return this.G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.G = value;
		}
	}

	internal virtual TextBox txtFormulas2
	{
		[CompilerGenerated]
		get
		{
			return this.H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.H = value;
		}
	}

	internal virtual CheckBox chkFormats
	{
		[CompilerGenerated]
		get
		{
			return this.K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.K = value;
		}
	}

	internal virtual ComboBox cbxFormatsFilter
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
			SelectionChangedEventHandler value2 = FormatsFilterChanged;
			ComboBox comboBox = this.E;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
			}
			this.E = value;
			comboBox = this.E;
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
				return;
			}
		}
	}

	internal virtual UniformGrid grdFormatsArgs
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual TextBox txtFormats1
	{
		[CompilerGenerated]
		get
		{
			return this.I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.I = value;
		}
	}

	internal virtual TextBox txtFormats2
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			J = value;
		}
	}

	internal virtual CheckBox chkLookInEmptyCells
	{
		[CompilerGenerated]
		get
		{
			return this.L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.L = value;
		}
	}

	internal virtual CheckBox chkData
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			M = value;
		}
	}

	internal virtual ComboBox cbxDataFilter
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

	internal virtual CheckBox chkRanges
	{
		[CompilerGenerated]
		get
		{
			return N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			N = value;
		}
	}

	internal virtual ComboBox cbxRangesFilter
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal virtual CheckBox chkComments
	{
		[CompilerGenerated]
		get
		{
			return O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			O = value;
		}
	}

	internal virtual ComboBox cbxCommentsFilter
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			SelectionChangedEventHandler value2 = CommentsFilterChanged;
			ComboBox comboBox = H;
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
			}
			H = value;
			comboBox = H;
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
				return;
			}
		}
	}

	internal virtual UniformGrid grdCommentsArgs
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

	internal virtual TextBox txtComments1
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			K = value;
		}
	}

	internal virtual TextBox txtComments2
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			L = value;
		}
	}

	internal virtual CheckBox chkNotes
	{
		[CompilerGenerated]
		get
		{
			return P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			P = value;
		}
	}

	internal virtual CheckBox chkNumericInputs
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	internal virtual CheckBox chkCharts
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	internal virtual CheckBox chkSparklines
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	internal virtual CheckBox chkShapes
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	internal virtual CheckBox chkHyperlinks
	{
		[CompilerGenerated]
		get
		{
			return U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			U = value;
		}
	}

	internal virtual CheckBox chkExploreMode
	{
		[CompilerGenerated]
		get
		{
			return V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = ExploreModeChecked;
			RoutedEventHandler value3 = ExploreModeUnchecked;
			CheckBox checkBox = V;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			V = value;
			checkBox = V;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual ComboBox cbxScope
	{
		[CompilerGenerated]
		get
		{
			return I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			I = value;
		}
	}

	internal virtual CheckBox chkLookInPrintAreas
	{
		[CompilerGenerated]
		get
		{
			return W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			W = value;
		}
	}

	internal virtual Button btnFind
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
			RoutedEventHandler value2 = btnFind_Click;
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

	public wpfSearch(wpfPane parent)
	{
		base.Loaded += ViewLoaded;
		base.Unloaded += ViewUnloaded;
		base.KeyDown += ViewKeyDown;
		this.m_A = null;
		InitializeComponent();
		ParentView = parent;
		MySettings settings = global::A.K.Settings;
		chkMatchCase.IsChecked = settings.AdvancedFindMatchCase;
		chkLookInValues.IsChecked = settings.AdvancedFindLookInValues;
		chkLookInFormulas.IsChecked = settings.AdvancedFindLookInFormulas;
		chkLookInComments.IsChecked = settings.AdvancedFindLookInComments;
		chkLookInCharts.IsChecked = settings.AdvancedFindLookInCharts;
		chkLookInHyperlinks.IsChecked = settings.AdvancedFindLookInHyperlinks;
		chkLookInEmptyCells.IsChecked = settings.AdvancedFindLookInEmptyCells;
		chkLookInPrintAreas.IsChecked = settings.AdvancedFindLookInPrintAreas;
		cbxScope.SelectedIndex = settings.AdvancedFindScope;
		settings = null;
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

	private void ViewLoaded(object sender, RoutedEventArgs e)
	{
		A();
		chkText.IsChecked = true;
	}

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void ViewKeyDown(object sender, KeyEventArgs e)
	{
		if (e.Key == Key.Return)
		{
			C();
		}
	}

	private void A()
	{
		chkText.Checked += TextSearchChecked;
		chkText.Unchecked += TextSearchUnchecked;
		chkValues.Checked += ValuesSearchChecked;
		chkValues.Unchecked += ValuesSearchUnchecked;
		chkDates.Checked += DatesSearchChecked;
		chkDates.Unchecked += DatesSearchUnchecked;
		chkFormulas.Checked += FormulasSearchChecked;
		chkFormulas.Unchecked += FormulasSearchUnchecked;
		chkFormats.Checked += FormatsSearchChecked;
		chkFormats.Unchecked += FormatsSearchUnchecked;
		chkData.Checked += DataSearchChecked;
		chkData.Unchecked += DataSearchUnchecked;
		chkRanges.Checked += RangesSearchChecked;
		chkRanges.Unchecked += RangesSearchUnchecked;
		chkComments.Checked += CommentsSearchChecked;
		chkComments.Unchecked += CommentsSearchUnchecked;
		chkLookInValues.Checked += LookInTextChanged;
		chkLookInValues.Unchecked += LookInTextChanged;
		chkLookInFormulas.Checked += LookInFormulasChanged;
		chkLookInFormulas.Unchecked += LookInFormulasChanged;
		chkLookInComments.Checked += LookInCommentsChanged;
		chkLookInComments.Unchecked += LookInCommentsChanged;
		chkLookInCharts.Checked += LookInChartsChanged;
		chkLookInCharts.Unchecked += LookInChartsChanged;
		chkLookInHyperlinks.Checked += LookInHyperlinksChanged;
		chkLookInHyperlinks.Unchecked += LookInHyperlinksChanged;
		chkLookInEmptyCells.Checked += LookInEmptyCellsChanged;
		chkLookInEmptyCells.Unchecked += LookInEmptyCellsChanged;
		chkLookInPrintAreas.Checked += LookInPrintAreasChanged;
		chkLookInPrintAreas.Unchecked += LookInPrintAreasChanged;
		cbxScope.SelectionChanged += SearchScopeChanged;
	}

	private void B()
	{
		chkText.Checked -= TextSearchChecked;
		chkText.Unchecked -= TextSearchUnchecked;
		chkValues.Checked -= ValuesSearchChecked;
		chkValues.Unchecked -= ValuesSearchUnchecked;
		chkDates.Checked -= DatesSearchChecked;
		chkDates.Unchecked -= DatesSearchUnchecked;
		chkFormulas.Checked -= FormulasSearchChecked;
		chkFormulas.Unchecked -= FormulasSearchUnchecked;
		chkFormats.Checked -= FormatsSearchChecked;
		chkFormats.Unchecked -= FormatsSearchUnchecked;
		chkData.Checked -= DataSearchChecked;
		chkData.Unchecked -= DataSearchUnchecked;
		chkRanges.Checked -= RangesSearchChecked;
		chkRanges.Unchecked -= RangesSearchUnchecked;
		chkComments.Checked -= CommentsSearchChecked;
		chkComments.Unchecked -= CommentsSearchUnchecked;
		chkLookInValues.Checked -= LookInTextChanged;
		chkLookInValues.Unchecked -= LookInTextChanged;
		chkLookInFormulas.Checked -= LookInFormulasChanged;
		chkLookInFormulas.Unchecked -= LookInFormulasChanged;
		chkLookInComments.Checked -= LookInCommentsChanged;
		chkLookInComments.Unchecked -= LookInCommentsChanged;
		chkLookInCharts.Checked -= LookInChartsChanged;
		chkLookInCharts.Unchecked -= LookInChartsChanged;
		chkLookInHyperlinks.Checked -= LookInHyperlinksChanged;
		chkLookInHyperlinks.Unchecked -= LookInHyperlinksChanged;
		chkLookInEmptyCells.Checked -= LookInEmptyCellsChanged;
		chkLookInEmptyCells.Unchecked -= LookInEmptyCellsChanged;
		chkLookInPrintAreas.Checked -= LookInPrintAreasChanged;
		chkLookInPrintAreas.Unchecked -= LookInPrintAreasChanged;
		cbxScope.SelectionChanged -= SearchScopeChanged;
	}

	private void TextSearchChecked(object sender, RoutedEventArgs e)
	{
		A((CheckBox)sender);
		A(cbxTextFilter, KF.A);
	}

	private void TextSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxTextFilter.SelectedIndex = -1;
		txtText1.Clear();
		txtText2.Clear();
	}

	private void ValuesSearchChecked(object sender, RoutedEventArgs e)
	{
		A((CheckBox)sender);
		A(cbxValuesFilter, KF.B);
	}

	private void ValuesSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxValuesFilter.SelectedIndex = -1;
		txtValues1.Clear();
		txtValues2.Clear();
	}

	private void DatesSearchChecked(object sender, RoutedEventArgs e)
	{
		A((CheckBox)sender);
		A(cbxDatesFilter, KF.C);
	}

	private void DatesSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxDatesFilter.SelectedIndex = -1;
		txtDates1.Clear();
		txtDates2.Clear();
	}

	private void FormulasSearchChecked(object sender, RoutedEventArgs e)
	{
		A((CheckBox)sender);
		A(cbxFormulasFilter, KF.E);
	}

	private void FormulasSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxFormulasFilter.SelectedIndex = -1;
		txtFormulas1.Clear();
		txtFormulas2.Clear();
	}

	private void FormatsSearchChecked(object sender, RoutedEventArgs e)
	{
		A((CheckBox)sender);
		A(cbxFormatsFilter, KF.D);
	}

	private void FormatsSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxFormatsFilter.SelectedIndex = -1;
		txtFormats1.Clear();
		txtFormats2.Clear();
	}

	private void RangesSearchChecked(object sender, RoutedEventArgs e)
	{
		A(cbxRangesFilter, KF.F);
	}

	private void RangesSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxRangesFilter.SelectedIndex = -1;
	}

	private void DataSearchChecked(object sender, RoutedEventArgs e)
	{
		A(cbxDataFilter, KF.G);
	}

	private void DataSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxDataFilter.SelectedIndex = -1;
	}

	private void CommentsSearchChecked(object sender, RoutedEventArgs e)
	{
		A(cbxCommentsFilter, KF.H);
	}

	private void CommentsSearchUnchecked(object sender, RoutedEventArgs e)
	{
		cbxCommentsFilter.SelectedIndex = -1;
		txtComments1.Clear();
		txtComments2.Clear();
	}

	private void A(ComboBox A, List<JF> B)
	{
		A.ItemsSource = B;
		A.SelectedIndex = 0;
	}

	private void A(CheckBox A)
	{
		CheckBox[] array = new CheckBox[5] { chkText, chkValues, chkDates, chkFormulas, chkFormats };
		foreach (CheckBox checkBox in array)
		{
			if (checkBox == A)
			{
				continue;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			checkBox.IsChecked = false;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void ExploreModeChecked(object sender, RoutedEventArgs e)
	{
		CheckBox[] array = A();
		foreach (CheckBox obj in array)
		{
			obj.IsChecked = false;
			obj.IsEnabled = false;
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
			return;
		}
	}

	private void ExploreModeUnchecked(object sender, RoutedEventArgs e)
	{
		CheckBox[] array = A();
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			array[i].IsEnabled = true;
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
			return;
		}
	}

	private CheckBox[] A()
	{
		return new CheckBox[14]
		{
			chkText, chkValues, chkDates, chkFormulas, chkFormats, chkData, chkRanges, chkComments, chkNotes, chkNumericInputs,
			chkCharts, chkSparklines, chkShapes, chkHyperlinks
		};
	}

	private void TextFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxTextFilter.SelectedIndex > -1)
		{
			chkMatchCase.Visibility = ((!A((ComboBox)sender).IsMatchCaseEnabled) ? Visibility.Collapsed : Visibility.Visible);
			A((ComboBox)sender, txtText1, txtText2, grdTextArgs);
		}
	}

	private void ValuesFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		A((ComboBox)sender, txtValues1, txtValues2, grdValuesArgs);
	}

	private void DatesFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		A((ComboBox)sender, txtDates1, txtDates2, grdDatesArgs);
	}

	private void FormulasFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		A((ComboBox)sender, txtFormulas1, txtFormulas2, grdFormulasArgs);
	}

	private void FormatsFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		A((ComboBox)sender, txtFormats1, txtFormats2, grdFormatsArgs);
	}

	private void CommentsFilterChanged(object sender, SelectionChangedEventArgs e)
	{
		A((ComboBox)sender, txtComments1, txtComments2, grdCommentsArgs);
	}

	private JF A(ComboBox A)
	{
		return (JF)A.SelectedItem;
	}

	private void A(ComboBox A, TextBox B, TextBox C, FrameworkElement D)
	{
		if (A.SelectedIndex == -1)
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
			JF jF = this.A(A);
			B.Tag = jF.PlaceholderText;
			B.ToolTip = jF.TextBoxToolTip;
			switch (jF.Arguments)
			{
			case 0:
				D.Visibility = Visibility.Collapsed;
				break;
			case 1:
				C.Visibility = Visibility.Collapsed;
				D.Visibility = Visibility.Visible;
				break;
			case 2:
				C.Visibility = Visibility.Visible;
				D.Visibility = Visibility.Visible;
				break;
			}
			if (D.Visibility == Visibility.Visible)
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
				B.Focus();
			}
			if (A == cbxFormatsFilter)
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
				chkLookInEmptyCells.Visibility = ((!jF.ShowIgnoreEmptyCells) ? Visibility.Collapsed : Visibility.Visible);
			}
			jF = null;
			return;
		}
	}

	private void LookInTextChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInValues = chkLookInValues.IsChecked.Value;
	}

	private void LookInFormulasChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInFormulas = chkLookInFormulas.IsChecked.Value;
	}

	private void LookInCommentsChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInComments = chkLookInComments.IsChecked.Value;
	}

	private void LookInChartsChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInCharts = chkLookInCharts.IsChecked.Value;
	}

	private void LookInHyperlinksChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInHyperlinks = chkLookInHyperlinks.IsChecked.Value;
	}

	private void LookInEmptyCellsChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInEmptyCells = chkLookInEmptyCells.IsChecked.Value;
	}

	private void LookInPrintAreasChanged(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindLookInPrintAreas = chkLookInPrintAreas.IsChecked.Value;
	}

	private void SearchScopeChanged(object sender, SelectionChangedEventArgs e)
	{
		global::A.K.Settings.AdvancedFindScope = cbxScope.SelectedIndex;
		CheckBox checkBox = chkLookInPrintAreas;
		if (cbxScope.SelectedIndex == 0)
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
			checkBox.Visibility = Visibility.Collapsed;
			checkBox.IsChecked = false;
		}
		else
		{
			checkBox.Visibility = Visibility.Visible;
		}
		checkBox = null;
	}

	private void btnFind_Click(object sender, RoutedEventArgs e)
	{
		C();
	}

	private void C()
	{
		List<BaseQuery> list = new List<BaseQuery>();
		bool value = chkExploreMode.IsChecked.Value;
		try
		{
			if (value)
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
				List<BaseQuery> list2 = list;
				list2.Add(new WorksheetQuery(VH.A(113553), Ranges.A));
				list2.Add(new WorksheetQuery(VH.A(113511), Ranges.B));
				list2.Add(new VF(VH.A(111078), ExcelAddIn1.SuperFind2.Callbacks.Formulas.E));
				list2.Add(new VF(VH.A(107985), MiscFormat.A));
				list2.Add(new VF(VH.A(114640), ExcelAddIn1.SuperFind2.Callbacks.Data.A));
				list2.Add(new VF(VH.A(112217), ExcelAddIn1.SuperFind2.Callbacks.Formulas.P));
				list2.Add(new VF(VH.A(112263), ExcelAddIn1.SuperFind2.Callbacks.Formulas.Q));
				list2.Add(new UF(VH.A(114147), Ranges.C));
				list2.Add(new UF(VH.A(114410), Ranges.A));
				list2.Add(new UF(VH.A(114661), Tables.D));
				list2.Add(new UF(VH.A(114711), Tables.A));
				list2.Add(new UF(VH.A(114745), Tables.B));
				list2.Add(new UF(VH.A(114791), Tables.C));
				list2.Add(A());
				list2.Add(B());
				list2.Add(A());
				list2.Add(A());
				list2.Add(B());
				list2.Add(C());
				list2.Add(D());
				_ = null;
			}
			else
			{
				if (chkText.IsChecked == true)
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
					JF jF = A(cbxTextFilter);
					if (!A(jF, txtText1, txtText2))
					{
						return;
					}
					list.Add(A((IF)jF, txtText1, txtText2));
				}
				if (chkValues.IsChecked == true)
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
					JF jF = A(cbxValuesFilter);
					if (!A(jF, txtValues1, txtValues2))
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
						break;
					}
					list.Add(A((IF)jF, txtValues1, txtValues2));
				}
				if (chkDates.IsChecked == true)
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
					JF jF = A(cbxDatesFilter);
					if (!A(jF, txtDates1, txtDates2))
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
						break;
					}
					list.Add(A((IF)jF, txtDates1, txtDates2));
				}
				if (chkFormulas.IsChecked == true)
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
					JF jF = A(cbxFormulasFilter);
					if (jF is IF)
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
						if (!A(jF, txtFormulas1, txtFormulas2))
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
							break;
						}
						list.Add(A((IF)jF, txtFormulas1, txtFormulas2));
					}
				}
				if (chkFormats.IsChecked == true)
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
					JF jF = A(cbxFormatsFilter);
					if (!A(jF, txtFormats1, txtFormats2))
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
						break;
					}
					list.Add(A((IF)jF, txtFormats1, txtFormats2));
				}
				if (chkData.IsChecked == true)
				{
					JF jF = A(cbxDataFilter);
					if (jF is HF)
					{
						list.Add(new UF((HF)jF));
					}
					else if (jF is IF)
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
						list.Add(new VF((IF)jF));
					}
					else if (jF is GF)
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
						list.Add(new TF((GF)jF));
					}
				}
				if (chkRanges.IsChecked == true)
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
					JF jF = A(cbxRangesFilter);
					if (jF is LF)
					{
						list.Add(new WorksheetQuery((LF)jF));
					}
					else if (jF is IF)
					{
						list.Add(new VF((IF)jF));
					}
					else if (jF is HF)
					{
						list.Add(new UF((HF)jF));
					}
				}
				if (chkComments.IsChecked == true)
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
					list.Add(A((HF)A(cbxCommentsFilter), txtComments1, txtComments2));
				}
				if (chkNotes.IsChecked == true)
				{
					list.Add(B());
				}
				if (chkNumericInputs.IsChecked == true)
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
					list.Add(A());
				}
				if (chkCharts.IsChecked == true)
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
					list.Add(A());
				}
				if (chkShapes.IsChecked == true)
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
					list.Add(B());
				}
				if (chkSparklines.IsChecked == true)
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
					list.Add(C());
				}
				if (chkHyperlinks.IsChecked == true)
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
					list.Add(D());
				}
			}
			if (list.Count > 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
					{
						wpfPane parentView = ParentView;
						parentView.B();
						parentView.ResultsView = new wpfResults(ParentView, list, (SearchScope)cbxScope.SelectedIndex, chkLookInPrintAreas.IsChecked.Value, value);
						_ = null;
						using List<BaseQuery>.Enumerator enumerator = list.GetEnumerator();
						while (enumerator.MoveNext())
						{
							BaseQuery current = enumerator.Current;
							clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)14, VH.A(125914) + current.UniqueId);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
					}
				}
			}
			C(VH.A(125925));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			B(ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			JF jF = null;
			list = null;
		}
	}

	private VF A(IF A, TextBox B, TextBox C)
	{
		return new VF(A, B.Text, C.Text, chkMatchCase.IsChecked.Value, chkLookInComments.IsChecked.Value, chkLookInCharts.IsChecked.Value, chkLookInEmptyCells.IsChecked.Value, chkLookInFormulas.IsChecked.Value, chkLookInHyperlinks.IsChecked.Value, chkLookInValues.IsChecked.Value);
	}

	private UF A(HF A, TextBox B, TextBox C)
	{
		return new UF(A, B.Text, C.Text, chkMatchCase.IsChecked.Value, chkLookInComments.IsChecked.Value, chkLookInCharts.IsChecked.Value, chkLookInEmptyCells.IsChecked.Value, chkLookInFormulas.IsChecked.Value, chkLookInHyperlinks.IsChecked.Value, chkLookInValues.IsChecked.Value);
	}

	private UF A()
	{
		return new UF(VH.A(114992), CommentsNotes.A);
	}

	private UF B()
	{
		return new UF(VH.A(125958), CommentsNotes.L);
	}

	private VF A()
	{
		return new VF(VH.A(125977), Values.M);
	}

	private WorksheetQuery A()
	{
		return new WorksheetQuery(VH.A(125998), Worksheet.A);
	}

	private WorksheetQuery B()
	{
		return new WorksheetQuery(VH.A(126013), Worksheet.B);
	}

	private UF C()
	{
		return new UF(VH.A(126028), Other.A);
	}

	private UF D()
	{
		return new UF(VH.A(126049), Other.B);
	}

	private bool A(JF A, TextBox B, TextBox C)
	{
		if (A.Arguments > 0)
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
			if (A.Arguments == 1)
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
				if (B.Text.Length == 0)
				{
					this.A(B);
					return false;
				}
			}
			if (A.Arguments == 2)
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
				if (B.Text.Length == 0)
				{
					this.A(B);
					return false;
				}
			}
			if (A.Arguments == 2)
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
				if (C.Text.Length == 0)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							this.A(C);
							return false;
						}
					}
				}
			}
		}
		return true;
	}

	private void A(TextBox A)
	{
		C(VH.A(126076));
		A.Focus();
	}

	private void B(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(126119), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
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
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					chkText = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			cbxTextFilter = (ComboBox)target;
			return;
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					grdTextArgs = (Grid)target;
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
					txtText1 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtText2 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			chkMatchCase = (CheckBox)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkLookInValues = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkLookInFormulas = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			chkLookInComments = (CheckBox)target;
			return;
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLookInCharts = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkLookInHyperlinks = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkValues = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			cbxValuesFilter = (ComboBox)target;
			return;
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					grdValuesArgs = (UniformGrid)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			txtValues1 = (TextBox)target;
			return;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtValues2 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkDates = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxDatesFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grdDatesArgs = (UniformGrid)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtDates1 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			txtDates2 = (TextBox)target;
			return;
		}
		if (connectionId == 22)
		{
			chkFormulas = (CheckBox)target;
			return;
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxFormulasFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					grdFormulasArgs = (UniformGrid)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			txtFormulas1 = (TextBox)target;
			return;
		}
		if (connectionId == 26)
		{
			txtFormulas2 = (TextBox)target;
			return;
		}
		if (connectionId == 27)
		{
			chkFormats = (CheckBox)target;
			return;
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxFormatsFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					grdFormatsArgs = (UniformGrid)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtFormats1 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtFormats2 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLookInEmptyCells = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkData = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					cbxDataFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkRanges = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxRangesFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkComments = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					cbxCommentsFilter = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			grdCommentsArgs = (UniformGrid)target;
			return;
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtComments1 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtComments2 = (TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			chkNotes = (CheckBox)target;
			return;
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkNumericInputs = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkCharts = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			chkSparklines = (CheckBox)target;
			return;
		}
		if (connectionId == 46)
		{
			chkShapes = (CheckBox)target;
			return;
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkHyperlinks = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			chkExploreMode = (CheckBox)target;
			return;
		}
		if (connectionId == 49)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxScope = (ComboBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 50:
			chkLookInPrintAreas = (CheckBox)target;
			break;
		case 51:
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				btnFind = (Button)target;
				return;
			}
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
}
