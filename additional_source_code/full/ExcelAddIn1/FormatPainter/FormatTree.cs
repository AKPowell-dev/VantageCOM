using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Shapes;
using A;
using ExcelAddIn1.Charts;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FormatPainter;

[DesignerGenerated]
public sealed class FormatTree : UserControl, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<KeyValuePair<int, Properties.ErrorBarsProperties>, bool> A;

		public static Func<KeyValuePair<int, Properties.UpDownBars>, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(KeyValuePair<int, Properties.ErrorBarsProperties> A)
		{
			return A.Value.HasErrorBars;
		}

		[SpecialName]
		internal bool A(KeyValuePair<int, Properties.UpDownBars> A)
		{
			return A.Value.HasUpDownBars;
		}
	}

	private readonly string m_A;

	private readonly string m_B;

	private readonly string m_C;

	private readonly string m_D;

	private readonly string m_E;

	private readonly string m_F;

	private bool m_A;

	[AccessedThroughProperty("btnCopy")]
	[CompilerGenerated]
	private Button m_A;

	[AccessedThroughProperty("btnApply")]
	[CompilerGenerated]
	private Button m_B;

	[AccessedThroughProperty("expSource")]
	[CompilerGenerated]
	private Polygon m_A;

	[AccessedThroughProperty("gridSource")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("imgSource")]
	[CompilerGenerated]
	private System.Windows.Controls.Image m_A;

	[AccessedThroughProperty("expChart")]
	[CompilerGenerated]
	private Polygon m_B;

	[AccessedThroughProperty("chkChart")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("gridChart")]
	private Grid m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkChartFormat")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtChartFormat")]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkChartSize")]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("txtChartSize")]
	private TextBlock m_B;

	[AccessedThroughProperty("chkChartTop")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("txtChartTop")]
	[CompilerGenerated]
	private TextBlock m_C;

	[AccessedThroughProperty("chkChartLeft")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("txtChartLeft")]
	private TextBlock m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("expPlotArea")]
	private Polygon m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPlotArea")]
	private System.Windows.Controls.CheckBox m_F;

	[AccessedThroughProperty("gridPlotArea")]
	[CompilerGenerated]
	private Grid m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPlotFormat")]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("txtPlotFormat")]
	[CompilerGenerated]
	private TextBlock m_E;

	[AccessedThroughProperty("chkPlotSize")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPlotSize")]
	private TextBlock m_F;

	[AccessedThroughProperty("chkPlotLocation")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPlotLocation")]
	private TextBlock m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("expSeries")]
	private Polygon m_D;

	[AccessedThroughProperty("chkSeries")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_J;

	[CompilerGenerated]
	[AccessedThroughProperty("gridSeries")]
	private Grid m_D;

	[AccessedThroughProperty("chkSeriesFormat")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_K;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSeriesFormat")]
	private TextBlock m_H;

	[AccessedThroughProperty("chkSeriesDataLabels")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSeriesDataLabels")]
	private TextBlock m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSeriesGap")]
	private System.Windows.Controls.CheckBox m_M;

	[AccessedThroughProperty("txtSeriesGap")]
	[CompilerGenerated]
	private TextBlock m_J;

	[AccessedThroughProperty("chkSeriesAngle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_N;

	[AccessedThroughProperty("txtSeriesAngle")]
	[CompilerGenerated]
	private TextBlock m_K;

	[AccessedThroughProperty("chkSeriesExplosion")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_O;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSeriesExplosion")]
	private TextBlock m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSeriesErrorBars")]
	private System.Windows.Controls.CheckBox m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSeriesErrorBars")]
	private TextBlock m_M;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSeriesUpDownBars")]
	private System.Windows.Controls.CheckBox m_Q;

	[AccessedThroughProperty("txtSeriesUpDownBars")]
	[CompilerGenerated]
	private TextBlock m_N;

	[AccessedThroughProperty("expAxes")]
	[CompilerGenerated]
	private Polygon m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAxes")]
	private System.Windows.Controls.CheckBox m_R;

	[CompilerGenerated]
	[AccessedThroughProperty("gridAxes")]
	private Grid m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriVal")]
	private System.Windows.Controls.CheckBox m_S;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriVal")]
	private TextBlock m_O;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriValScale")]
	private System.Windows.Controls.CheckBox m_T;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriValScale")]
	private TextBlock m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriValGridlines")]
	private System.Windows.Controls.CheckBox m_U;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriValGridlines")]
	private TextBlock m_Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriValTicks")]
	private System.Windows.Controls.CheckBox m_V;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriValTicks")]
	private TextBlock m_R;

	[AccessedThroughProperty("chkPriValTitle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_W;

	[AccessedThroughProperty("txtPriValTitle")]
	[CompilerGenerated]
	private TextBlock m_S;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriCat")]
	private System.Windows.Controls.CheckBox m_X;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriCat")]
	private TextBlock m_T;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriCatScale")]
	private System.Windows.Controls.CheckBox m_Y;

	[CompilerGenerated]
	[AccessedThroughProperty("txtPriCatScale")]
	private TextBlock m_U;

	[AccessedThroughProperty("chkPriCatGridlines")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox Z;

	[AccessedThroughProperty("txtPriCatGridlines")]
	[CompilerGenerated]
	private TextBlock m_V;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPriCatTicks")]
	private System.Windows.Controls.CheckBox AB;

	[AccessedThroughProperty("txtPriCatTicks")]
	[CompilerGenerated]
	private TextBlock m_W;

	[AccessedThroughProperty("chkPriCatTitle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox BB;

	[AccessedThroughProperty("txtPriCatTitle")]
	[CompilerGenerated]
	private TextBlock m_X;

	[AccessedThroughProperty("chkSecVal")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox CB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecVal")]
	private TextBlock m_Y;

	[AccessedThroughProperty("chkSecValScale")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox DB;

	[AccessedThroughProperty("txtSecValScale")]
	[CompilerGenerated]
	private TextBlock Z;

	[AccessedThroughProperty("chkSecValGridlines")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox EB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecValGridlines")]
	private TextBlock AB;

	[AccessedThroughProperty("chkSecValTicks")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox FB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecValTicks")]
	private TextBlock BB;

	[AccessedThroughProperty("chkSecValTitle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox GB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecValTitle")]
	private TextBlock CB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSecCat")]
	private System.Windows.Controls.CheckBox HB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecCat")]
	private TextBlock DB;

	[AccessedThroughProperty("chkSecCatScale")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox IB;

	[AccessedThroughProperty("txtSecCatScale")]
	[CompilerGenerated]
	private TextBlock EB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSecCatGridlines")]
	private System.Windows.Controls.CheckBox JB;

	[AccessedThroughProperty("txtSecCatGridlines")]
	[CompilerGenerated]
	private TextBlock FB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSecCatTicks")]
	private System.Windows.Controls.CheckBox KB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecCatTicks")]
	private TextBlock GB;

	[AccessedThroughProperty("chkSecCatTitle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox LB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtSecCatTitle")]
	private TextBlock HB;

	[CompilerGenerated]
	[AccessedThroughProperty("expLegend")]
	private Polygon m_F;

	[AccessedThroughProperty("chkLegend")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox MB;

	[AccessedThroughProperty("gridLegend")]
	[CompilerGenerated]
	private Grid m_F;

	[AccessedThroughProperty("chkLegendFormat")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox NB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLegendFormat")]
	private TextBlock IB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLegendPosn")]
	private System.Windows.Controls.CheckBox OB;

	[AccessedThroughProperty("txtLegendPosn")]
	[CompilerGenerated]
	private TextBlock JB;

	[CompilerGenerated]
	[AccessedThroughProperty("expTitle")]
	private Polygon m_G;

	[AccessedThroughProperty("chkTitle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox PB;

	[CompilerGenerated]
	[AccessedThroughProperty("gridTitle")]
	private Grid m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTitleFormat")]
	private System.Windows.Controls.CheckBox QB;

	[AccessedThroughProperty("txtTitleFormat")]
	[CompilerGenerated]
	private TextBlock KB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTitlePosn")]
	private System.Windows.Controls.CheckBox RB;

	[AccessedThroughProperty("txtTitlePosn")]
	[CompilerGenerated]
	private TextBlock LB;

	[AccessedThroughProperty("expDataTable")]
	[CompilerGenerated]
	private Polygon m_H;

	[AccessedThroughProperty("chkDataTable")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox SB;

	[CompilerGenerated]
	[AccessedThroughProperty("gridDataTable")]
	private Grid m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkDataTableFormat")]
	private System.Windows.Controls.CheckBox TB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtDataTableFormat")]
	private TextBlock MB;

	private bool m_B;

	public bool Visible
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A();
		}
	}

	internal virtual Button btnCopy
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
			RoutedEventHandler value2 = btnCopy_Click;
			Button button = this.m_A;
			if (button != null)
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

	internal virtual Button btnApply
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
			RoutedEventHandler value2 = btnApply_Click;
			Button button = this.m_B;
			if (button != null)
			{
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

	internal virtual Polygon expSource
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

	internal virtual Grid gridSource
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

	internal virtual System.Windows.Controls.Image imgSource
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

	internal virtual Polygon expChart
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

	internal virtual System.Windows.Controls.CheckBox chkChart
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
			RoutedEventHandler value2 = ChartCheckedChanged;
			RoutedEventHandler value3 = ChartCheckedChanged;
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
				switch (5)
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

	internal virtual Grid gridChart
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

	internal virtual System.Windows.Controls.CheckBox chkChartFormat
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

	internal virtual TextBlock txtChartFormat
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

	internal virtual System.Windows.Controls.CheckBox chkChartSize
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

	internal virtual TextBlock txtChartSize
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

	internal virtual System.Windows.Controls.CheckBox chkChartTop
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

	internal virtual TextBlock txtChartTop
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

	internal virtual System.Windows.Controls.CheckBox chkChartLeft
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual TextBlock txtChartLeft
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

	internal virtual Polygon expPlotArea
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

	internal virtual System.Windows.Controls.CheckBox chkPlotArea
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = PlotAreaCheckedChanged;
			RoutedEventHandler value3 = PlotAreaCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_F;
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
			this.m_F = value;
			checkBox = this.m_F;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual Grid gridPlotArea
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

	internal virtual System.Windows.Controls.CheckBox chkPlotFormat
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual TextBlock txtPlotFormat
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPlotSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual TextBlock txtPlotSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPlotLocation
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal virtual TextBlock txtPlotLocation
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual Polygon expSeries
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

	internal virtual System.Windows.Controls.CheckBox chkSeries
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = SeriesCheckedChanged;
			RoutedEventHandler value3 = SeriesCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_J;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_J = value;
			checkBox = this.m_J;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual Grid gridSeries
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

	internal virtual System.Windows.Controls.CheckBox chkSeriesFormat
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal virtual TextBlock txtSeriesFormat
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesDataLabels
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_L = value;
		}
	}

	internal virtual TextBlock txtSeriesDataLabels
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesGap
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_M = value;
		}
	}

	internal virtual TextBlock txtSeriesGap
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesAngle
	{
		[CompilerGenerated]
		get
		{
			return this.m_N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_N = value;
		}
	}

	internal virtual TextBlock txtSeriesAngle
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesExplosion
	{
		[CompilerGenerated]
		get
		{
			return this.m_O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_O = value;
		}
	}

	internal virtual TextBlock txtSeriesExplosion
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_L = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesErrorBars
	{
		[CompilerGenerated]
		get
		{
			return this.m_P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_P = value;
		}
	}

	internal virtual TextBlock txtSeriesErrorBars
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_M = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSeriesUpDownBars
	{
		[CompilerGenerated]
		get
		{
			return this.m_Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Q = value;
		}
	}

	internal virtual TextBlock txtSeriesUpDownBars
	{
		[CompilerGenerated]
		get
		{
			return this.m_N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_N = value;
		}
	}

	internal virtual Polygon expAxes
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkAxes
	{
		[CompilerGenerated]
		get
		{
			return this.m_R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_R = value;
		}
	}

	internal virtual Grid gridAxes
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriVal
	{
		[CompilerGenerated]
		get
		{
			return this.m_S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = PriValAxisCheckedChanged;
			RoutedEventHandler value3 = PriValAxisCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_S;
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
			this.m_S = value;
			checkBox = this.m_S;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
			}
		}
	}

	internal virtual TextBlock txtPriVal
	{
		[CompilerGenerated]
		get
		{
			return this.m_O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_O = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriValScale
	{
		[CompilerGenerated]
		get
		{
			return this.m_T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_T = value;
		}
	}

	internal virtual TextBlock txtPriValScale
	{
		[CompilerGenerated]
		get
		{
			return this.m_P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_P = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriValGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.m_U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_U = value;
		}
	}

	internal virtual TextBlock txtPriValGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.m_Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Q = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriValTicks
	{
		[CompilerGenerated]
		get
		{
			return this.m_V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_V = value;
		}
	}

	internal virtual TextBlock txtPriValTicks
	{
		[CompilerGenerated]
		get
		{
			return this.m_R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_R = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriValTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_W = value;
		}
	}

	internal virtual TextBlock txtPriValTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_S = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriCat
	{
		[CompilerGenerated]
		get
		{
			return this.m_X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = PriCatAxisCheckedChanged;
			RoutedEventHandler value3 = PriCatAxisCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.m_X;
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
			this.m_X = value;
			checkBox = this.m_X;
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

	internal virtual TextBlock txtPriCat
	{
		[CompilerGenerated]
		get
		{
			return this.m_T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_T = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriCatScale
	{
		[CompilerGenerated]
		get
		{
			return this.m_Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Y = value;
		}
	}

	internal virtual TextBlock txtPriCatScale
	{
		[CompilerGenerated]
		get
		{
			return this.m_U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_U = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriCatGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.Z = value;
		}
	}

	internal virtual TextBlock txtPriCatGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.m_V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_V = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriCatTicks
	{
		[CompilerGenerated]
		get
		{
			return this.AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.AB = value;
		}
	}

	internal virtual TextBlock txtPriCatTicks
	{
		[CompilerGenerated]
		get
		{
			return this.m_W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_W = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPriCatTitle
	{
		[CompilerGenerated]
		get
		{
			return this.BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.BB = value;
		}
	}

	internal virtual TextBlock txtPriCatTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_X = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecVal
	{
		[CompilerGenerated]
		get
		{
			return this.CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = SecValAxisCheckedChanged;
			RoutedEventHandler value3 = SecValAxisCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.CB;
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
			this.CB = value;
			checkBox = this.CB;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual TextBlock txtSecVal
	{
		[CompilerGenerated]
		get
		{
			return this.m_Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Y = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecValScale
	{
		[CompilerGenerated]
		get
		{
			return this.DB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.DB = value;
		}
	}

	internal virtual TextBlock txtSecValScale
	{
		[CompilerGenerated]
		get
		{
			return Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Z = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecValGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.EB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.EB = value;
		}
	}

	internal virtual TextBlock txtSecValGridlines
	{
		[CompilerGenerated]
		get
		{
			return AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			AB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecValTicks
	{
		[CompilerGenerated]
		get
		{
			return this.FB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.FB = value;
		}
	}

	internal virtual TextBlock txtSecValTicks
	{
		[CompilerGenerated]
		get
		{
			return BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			BB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecValTitle
	{
		[CompilerGenerated]
		get
		{
			return this.GB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.GB = value;
		}
	}

	internal virtual TextBlock txtSecValTitle
	{
		[CompilerGenerated]
		get
		{
			return CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			CB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecCat
	{
		[CompilerGenerated]
		get
		{
			return this.HB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = SecCatAxisCheckedChanged;
			RoutedEventHandler value3 = SecCatAxisCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.HB;
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
			this.HB = value;
			checkBox = this.HB;
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

	internal virtual TextBlock txtSecCat
	{
		[CompilerGenerated]
		get
		{
			return DB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			DB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecCatScale
	{
		[CompilerGenerated]
		get
		{
			return this.IB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.IB = value;
		}
	}

	internal virtual TextBlock txtSecCatScale
	{
		[CompilerGenerated]
		get
		{
			return EB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecCatGridlines
	{
		[CompilerGenerated]
		get
		{
			return this.JB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.JB = value;
		}
	}

	internal virtual TextBlock txtSecCatGridlines
	{
		[CompilerGenerated]
		get
		{
			return FB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			FB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecCatTicks
	{
		[CompilerGenerated]
		get
		{
			return this.KB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.KB = value;
		}
	}

	internal virtual TextBlock txtSecCatTicks
	{
		[CompilerGenerated]
		get
		{
			return GB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			GB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSecCatTitle
	{
		[CompilerGenerated]
		get
		{
			return this.LB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.LB = value;
		}
	}

	internal virtual TextBlock txtSecCatTitle
	{
		[CompilerGenerated]
		get
		{
			return HB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			HB = value;
		}
	}

	internal virtual Polygon expLegend
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLegend
	{
		[CompilerGenerated]
		get
		{
			return this.MB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = LegendCheckedChanged;
			RoutedEventHandler value3 = LegendCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = this.MB;
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
			this.MB = value;
			checkBox = this.MB;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual Grid gridLegend
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLegendFormat
	{
		[CompilerGenerated]
		get
		{
			return NB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			NB = value;
		}
	}

	internal virtual TextBlock txtLegendFormat
	{
		[CompilerGenerated]
		get
		{
			return IB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			IB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLegendPosn
	{
		[CompilerGenerated]
		get
		{
			return OB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			OB = value;
		}
	}

	internal virtual TextBlock txtLegendPosn
	{
		[CompilerGenerated]
		get
		{
			return JB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			JB = value;
		}
	}

	internal virtual Polygon expTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkTitle
	{
		[CompilerGenerated]
		get
		{
			return PB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = TitleCheckedChanged;
			RoutedEventHandler value3 = TitleCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = PB;
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
			PB = value;
			checkBox = PB;
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

	internal virtual Grid gridTitle
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkTitleFormat
	{
		[CompilerGenerated]
		get
		{
			return QB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			QB = value;
		}
	}

	internal virtual TextBlock txtTitleFormat
	{
		[CompilerGenerated]
		get
		{
			return KB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			KB = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkTitlePosn
	{
		[CompilerGenerated]
		get
		{
			return RB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RB = value;
		}
	}

	internal virtual TextBlock txtTitlePosn
	{
		[CompilerGenerated]
		get
		{
			return LB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			LB = value;
		}
	}

	internal virtual Polygon expDataTable
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkDataTable
	{
		[CompilerGenerated]
		get
		{
			return SB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = DataTableCheckedChanged;
			RoutedEventHandler value3 = DataTableCheckedChanged;
			System.Windows.Controls.CheckBox checkBox = SB;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			SB = value;
			checkBox = SB;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
			}
		}
	}

	internal virtual Grid gridDataTable
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkDataTableFormat
	{
		[CompilerGenerated]
		get
		{
			return TB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			TB = value;
		}
	}

	internal virtual TextBlock txtDataTableFormat
	{
		[CompilerGenerated]
		get
		{
			return MB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			MB = value;
		}
	}

	public FormatTree()
	{
		base.Loaded += FormatTree_Loaded;
		this.m_A = VH.A(172798);
		this.m_B = VH.A(172807);
		this.m_C = VH.A(172822);
		this.m_D = VH.A(172837);
		this.m_E = VH.A(172846);
		this.m_F = VH.A(172861);
		InitializeComponent();
	}

	private void FormatTree_Loaded(object sender, RoutedEventArgs e)
	{
		B();
		D();
		F();
		H();
		J();
		L();
		N();
		P();
		R();
		T();
		V();
		X();
	}

	private void A()
	{
	}

	private void A(object A, Range B)
	{
	}

	private void ExpandCollapseSource(object sender, MouseButtonEventArgs e)
	{
		A(expSource, gridSource);
	}

	private void ExpandCollapseChart(object sender, MouseButtonEventArgs e)
	{
		A(expChart, gridChart);
	}

	private void ExpandCollapsePlotArea(object sender, MouseButtonEventArgs e)
	{
		A(expPlotArea, gridPlotArea);
	}

	private void ExpandCollapseSeries(object sender, MouseButtonEventArgs e)
	{
		A(expSeries, gridSeries);
	}

	private void ExpandCollapseAxes(object sender, MouseButtonEventArgs e)
	{
		A(expAxes, gridAxes);
	}

	private void ExpandCollapseLegend(object sender, MouseButtonEventArgs e)
	{
		A(expLegend, gridLegend);
	}

	private void ExpandCollapseTitle(object sender, MouseButtonEventArgs e)
	{
		A(expTitle, gridTitle);
	}

	private void ExpandCollapseDataTable(object sender, MouseButtonEventArgs e)
	{
		A(expDataTable, gridDataTable);
	}

	private void A(Polygon A, Grid B)
	{
		if (B.Visibility == Visibility.Visible)
		{
			C(A, B);
		}
		else
		{
			this.B(A, B);
		}
	}

	private void B(Polygon A, Grid B)
	{
		B.Visibility = Visibility.Visible;
		A.Points = new PointCollection(new System.Windows.Point[3]
		{
			new System.Windows.Point(1.0, 9.0),
			new System.Windows.Point(9.0, 9.0),
			new System.Windows.Point(9.0, 1.0)
		});
	}

	private void C(Polygon A, Grid B)
	{
		B.Visibility = Visibility.Collapsed;
		A.Points = new PointCollection(new System.Windows.Point[3]
		{
			new System.Windows.Point(1.0, 1.0),
			new System.Windows.Point(7.0, 6.0),
			new System.Windows.Point(1.0, 11.0)
		});
	}

	private void btnCopy_Click(object sender, RoutedEventArgs e)
	{
		Chart chart = Helpers.SelectedChart();
		if (chart != null)
		{
			Pane.CopiedProperties = new Properties(chart);
			PopulateProperties();
			chart = null;
		}
		else
		{
			Helpers.NoChartMessage();
		}
	}

	private void btnApply_Click(object sender, RoutedEventArgs e)
	{
		Options options = new Options();
		bool flag = false;
		if (A(chkChart))
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
			Options.myChart chart = options.Chart;
			chart.Size = chkChartSize.IsChecked.Value;
			chart.Top = chkChartTop.IsChecked.Value;
			chart.Left = chkChartLeft.IsChecked.Value;
			chart.Format = chkChartFormat.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkPlotArea))
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
			Options.myPlotArea plotArea = options.PlotArea;
			plotArea.Size = chkPlotSize.IsChecked.Value;
			plotArea.Location = chkPlotLocation.IsChecked.Value;
			plotArea.Format = chkPlotFormat.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkSeries))
		{
			Options.mySeries series = options.Series;
			series.Format = chkSeriesFormat.IsChecked.Value;
			series.GapWidthOverlap = chkSeriesGap.IsChecked.Value;
			series.FirstSliceAngle = chkSeriesAngle.IsChecked.Value;
			series.Explosion = chkSeriesExplosion.IsChecked.Value;
			series.DataLabels = chkSeriesDataLabels.IsChecked.Value;
			series.ErrorBars = chkSeriesErrorBars.IsChecked.Value;
			series.UpDownBars = chkSeriesUpDownBars.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkAxes))
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
			if (A(chkPriVal))
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
				Options.myPrimaryValueAxis primaryValueAxis = options.PrimaryValueAxis;
				primaryValueAxis.Scale = chkPriValScale.IsChecked.Value;
				primaryValueAxis.Gridlines = chkPriValGridlines.IsChecked.Value;
				primaryValueAxis.Ticks = chkPriValTicks.IsChecked.Value;
				primaryValueAxis.Title = chkPriValTitle.IsChecked.Value;
				_ = null;
				flag = true;
			}
			if (A(chkPriCat))
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
				Options.myPrimaryCategoryAxis primaryCategoryAxis = options.PrimaryCategoryAxis;
				primaryCategoryAxis.Scale = chkPriCatScale.IsChecked.Value;
				primaryCategoryAxis.Gridlines = chkPriCatGridlines.IsChecked.Value;
				primaryCategoryAxis.Ticks = chkPriCatTicks.IsChecked.Value;
				primaryCategoryAxis.Title = chkPriCatTitle.IsChecked.Value;
				_ = null;
				flag = true;
			}
			if (A(chkSecVal))
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
				Options.mySecondaryValueAxis secondaryValueAxis = options.SecondaryValueAxis;
				secondaryValueAxis.Scale = chkSecValScale.IsChecked.Value;
				secondaryValueAxis.Gridlines = chkSecValGridlines.IsChecked.Value;
				secondaryValueAxis.Ticks = chkSecValTicks.IsChecked.Value;
				secondaryValueAxis.Title = chkSecValTitle.IsChecked.Value;
				_ = null;
				flag = true;
			}
			if (A(chkSecCat))
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
				Options.mySecondaryCategoryAxis secondaryCategoryAxis = options.SecondaryCategoryAxis;
				secondaryCategoryAxis.Scale = chkSecCatScale.IsChecked.Value;
				secondaryCategoryAxis.Gridlines = chkSecCatGridlines.IsChecked.Value;
				secondaryCategoryAxis.Ticks = chkSecCatTicks.IsChecked.Value;
				secondaryCategoryAxis.Title = chkSecCatTitle.IsChecked.Value;
				_ = null;
			}
			flag = true;
		}
		if (A(chkLegend))
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
			Options.myLegend legend = options.Legend;
			legend.Format = chkLegendFormat.IsChecked.Value;
			legend.Position = chkLegendPosn.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkTitle))
		{
			Options.myTitle title = options.Title;
			title.Format = chkTitleFormat.IsChecked.Value;
			title.Position = chkTitlePosn.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkDataTable))
		{
			options.DataTable.Format = chkDataTableFormat.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (flag)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					Apply.ToSelection(Pane.CopiedProperties, options);
					return;
				}
			}
		}
		MessageBox.Show(VH.A(172184), VH.A(40448), MessageBoxButton.OK, MessageBoxImage.Exclamation);
	}

	private bool A(System.Windows.Controls.CheckBox A)
	{
		bool? isChecked;
		bool? flag = (isChecked = A.IsChecked);
		bool? flag2;
		if (flag.HasValue)
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
			if (isChecked == true)
			{
				flag2 = true;
				goto IL_006d;
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
		if (A.IsChecked.HasValue)
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
			flag2 = isChecked;
		}
		else
		{
			flag2 = true;
		}
		goto IL_006d;
		IL_006d:
		isChecked = flag2;
		return isChecked.Value;
	}

	public void PopulateProperties()
	{
		Properties copiedProperties = Pane.CopiedProperties;
		bool isMetric = RegionInfo.CurrentRegion.IsMetric;
		string value = Conversions.ToString(Operators.CompareString(clsPublish.SystemDecimalSeparator(), VH.A(64021), TextCompare: false) == 0);
		try
		{
			A(copiedProperties, Conversions.ToBoolean(value), isMetric);
			B(copiedProperties, Conversions.ToBoolean(value), isMetric);
			C(copiedProperties);
			E(copiedProperties);
			F(copiedProperties);
			G(copiedProperties);
			H(copiedProperties);
			A(copiedProperties, Conversions.ToBoolean(value));
			B(copiedProperties, Conversions.ToBoolean(value));
			D(copiedProperties);
			if (copiedProperties.Chart.HasPrimaryValueAxis)
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
				D(A: true);
			}
			else
			{
				D(A: false);
			}
			if (copiedProperties.Chart.HasPrimaryCategoryAxis)
			{
				E(A: true);
			}
			else
			{
				E(A: false);
			}
			if (copiedProperties.Chart.HasSecondaryValueAxis)
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
				F(A: true);
			}
			else
			{
				F(A: false);
			}
			if (copiedProperties.Chart.HasSecondaryCategoryAxis)
			{
				G(A: true);
			}
			else
			{
				G(A: false);
			}
			if (copiedProperties.Chart.HasLegend)
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
				A(A: true);
				B(expLegend, gridLegend);
			}
			else
			{
				A(A: false);
				C(expLegend, gridLegend);
			}
			if (copiedProperties.Chart.HasTitle)
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
				B(A: true);
				B(expTitle, gridTitle);
			}
			else
			{
				B(A: false);
				C(expTitle, gridTitle);
			}
			if (copiedProperties.Chart.HasDataTable)
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
				C(A: true);
				B(expDataTable, gridDataTable);
			}
			else
			{
				C(A: false);
				C(expDataTable, gridDataTable);
			}
			if (copiedProperties.SourceImage != null)
			{
				imgSource.Source = copiedProperties.SourceImage;
			}
			UpdateLayout();
			btnApply.IsEnabled = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButton.OK, MessageBoxImage.Hand);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private void A(Properties A, bool B, bool C)
	{
		Properties.ChartObjectProperties chartObject = A.ChartObject;
		if (C)
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
			string text;
			if (!B)
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
				text = this.m_F;
			}
			else
			{
				text = this.m_C;
			}
			string text2 = text;
			txtChartSize.Text = clsPublish.PointsToCentimeters((float)chartObject.Width).ToString(text2) + VH.A(172237) + clsPublish.PointsToCentimeters((float)chartObject.Height).ToString(text2);
			txtChartLeft.Text = clsPublish.PointsToCentimeters((float)chartObject.Left).ToString(text2);
			txtChartTop.Text = clsPublish.PointsToCentimeters((float)chartObject.Top).ToString(text2);
		}
		else
		{
			string text2 = (B ? this.m_B : this.m_E);
			txtChartSize.Text = clsPublish.PointsToInches((float)chartObject.Width).ToString(text2) + VH.A(172237) + clsPublish.PointsToInches((float)chartObject.Height).ToString(text2);
			txtChartLeft.Text = clsPublish.PointsToInches((float)chartObject.Left).ToString(text2);
			txtChartTop.Text = clsPublish.PointsToInches((float)chartObject.Top).ToString(text2);
		}
		txtChartFormat.Text = VH.A(172244);
	}

	private void B(Properties A, bool B, bool C)
	{
		Properties.PlotAreaProperties plotArea = A.PlotArea;
		if (C)
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
			string text = (B ? this.m_C : this.m_F);
			txtPlotSize.Text = clsPublish.PointsToCentimeters((float)plotArea.InsideWidth).ToString(text) + VH.A(172237) + clsPublish.PointsToCentimeters((float)plotArea.InsideHeight).ToString(text);
			txtPlotLocation.Text = clsPublish.PointsToCentimeters((float)plotArea.InsideLeft).ToString(text) + VH.A(10524) + clsPublish.PointsToCentimeters((float)plotArea.InsideTop).ToString(text);
		}
		else
		{
			string text2;
			if (!B)
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
				text2 = this.m_E;
			}
			else
			{
				text2 = this.m_B;
			}
			string text = text2;
			txtPlotSize.Text = clsPublish.PointsToInches((float)plotArea.InsideWidth).ToString(text) + VH.A(172237) + clsPublish.PointsToInches((float)plotArea.InsideHeight).ToString(text);
			txtPlotLocation.Text = clsPublish.PointsToInches((float)plotArea.InsideLeft).ToString(text) + VH.A(10524) + clsPublish.PointsToInches((float)plotArea.InsideTop).ToString(text);
		}
		plotArea = default(Properties.PlotAreaProperties);
		txtPlotFormat.Text = VH.A(172244);
	}

	private void C(Properties A)
	{
		txtSeriesFormat.Text = VH.A(172269) + A.Series.Fill.Count + VH.A(172298);
		List<string> list = new List<string>();
		using (Dictionary<int, int>.Enumerator enumerator = A.Series.GapWidth.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				list.Add(((double)enumerator.Current.Value / 100.0).ToString(VH.A(75697)));
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
				break;
			}
		}
		using (Dictionary<int, int>.Enumerator enumerator2 = A.Series.Overlap.GetEnumerator())
		{
			while (enumerator2.MoveNext())
			{
				list.Add(((double)enumerator2.Current.Value / 100.0).ToString(VH.A(75697)));
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_012a;
				}
				continue;
				end_IL_012a:
				break;
			}
		}
		this.A(list, chkSeriesGap, txtSeriesGap);
		list = new List<string>();
		foreach (KeyValuePair<int, int> item in A.Series.FirstSliceAngle)
		{
			list.Add(item.Value.ToString(VH.A(64024)) + VH.A(172315));
		}
		this.A(list, chkSeriesAngle, txtSeriesAngle);
		list = new List<string>();
		foreach (KeyValuePair<int, int> item2 in A.Series.Explosion)
		{
			list.Add(((double)item2.Value / 100.0).ToString(VH.A(75697)));
		}
		this.A(list, chkSeriesExplosion, txtSeriesExplosion);
		list = new List<string>();
		using (Dictionary<int, Properties.DataLabelProperties>.Enumerator enumerator5 = A.Series.DataLabels.GetEnumerator())
		{
			while (enumerator5.MoveNext())
			{
				KeyValuePair<int, Properties.DataLabelProperties> current = enumerator5.Current;
				if (current.Value.HasDataLabels)
				{
					XlDataLabelPosition position = current.Value.Position;
					if (position <= XlDataLabelPosition.xlLabelPositionLeft)
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
						switch (position)
						{
						case XlDataLabelPosition.xlLabelPositionRight:
							list.Add(VH.A(56591));
							continue;
						case XlDataLabelPosition.xlLabelPositionLeft:
							list.Add(VH.A(56582));
							continue;
						}
					}
					else
					{
						if (position == XlDataLabelPosition.xlLabelPositionCenter)
						{
							list.Add(VH.A(172324));
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
						switch (position)
						{
						case XlDataLabelPosition.xlLabelPositionAbove:
							list.Add(VH.A(56560));
							continue;
						case XlDataLabelPosition.xlLabelPositionBelow:
							list.Add(VH.A(56571));
							continue;
						case XlDataLabelPosition.xlLabelPositionInsideBase:
							list.Add(VH.A(56602));
							continue;
						case XlDataLabelPosition.xlLabelPositionInsideEnd:
							list.Add(VH.A(56625));
							continue;
						case XlDataLabelPosition.xlLabelPositionOutsideEnd:
							list.Add(VH.A(56646));
							continue;
						case XlDataLabelPosition.xlLabelPositionBestFit:
							list.Add(VH.A(172337));
							continue;
						}
					}
					list.Add(VH.A(106865));
				}
				else
				{
					list.Add(VH.A(61174));
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_041d;
				}
				continue;
				end_IL_041d:
				break;
			}
		}
		this.A(list, chkSeriesDataLabels, txtSeriesDataLabels);
		Dictionary<int, Properties.ErrorBarsProperties> errorBars = A.Series.ErrorBars;
		Func<KeyValuePair<int, Properties.ErrorBarsProperties>, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (KeyValuePair<int, Properties.ErrorBarsProperties> keyValuePair) => keyValuePair.Value.HasErrorBars);
		}
		else
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
			predicate = _Closure_0024__.A;
		}
		if (errorBars.Where(predicate).Any())
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
			chkSeriesErrorBars.IsEnabled = true;
			txtSeriesErrorBars.Text = VH.A(172354);
		}
		else
		{
			chkSeriesErrorBars.IsEnabled = false;
			chkSeriesErrorBars.IsChecked = false;
			txtSeriesErrorBars.Text = VH.A(13778);
		}
		Dictionary<int, Properties.UpDownBars> upBars = A.Series.UpBars;
		Func<KeyValuePair<int, Properties.UpDownBars>, bool> predicate2;
		if (_Closure_0024__.A == null)
		{
			predicate2 = (_Closure_0024__.A = [SpecialName] (KeyValuePair<int, Properties.UpDownBars> keyValuePair) => keyValuePair.Value.HasUpDownBars);
		}
		else
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
			predicate2 = _Closure_0024__.A;
		}
		if (upBars.Where(predicate2).Any())
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
			chkSeriesUpDownBars.IsEnabled = true;
			txtSeriesUpDownBars.Text = VH.A(172377);
		}
		else
		{
			chkSeriesUpDownBars.IsEnabled = false;
			chkSeriesUpDownBars.IsChecked = false;
			txtSeriesUpDownBars.Text = VH.A(13778);
		}
		list = null;
	}

	private void A(List<string> A, System.Windows.Controls.CheckBox B, TextBlock C)
	{
		if (A.Any())
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
					B.IsEnabled = true;
					C.Text = string.Join(VH.A(10524), A.ToArray());
					return;
				}
			}
		}
		B.IsEnabled = false;
		B.IsChecked = false;
		C.Text = VH.A(13778);
	}

	private void A(Properties A, bool B)
	{
		if (A.Chart.HasLegend)
		{
			while (true)
			{
				string text;
				Properties.LegendProperties legend;
				string text2;
				switch (1)
				{
				case 0:
					break;
				default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						legend = A.Legend;
						XlLegendPosition position = legend.Position;
						if (position <= XlLegendPosition.xlLegendPositionRight)
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
							if (position != XlLegendPosition.xlLegendPositionTop)
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
								if (position != XlLegendPosition.xlLegendPositionRight)
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
									goto IL_00e5;
								}
								text = VH.A(56591);
							}
							else
							{
								text = VH.A(57409);
							}
						}
						else if (position != XlLegendPosition.xlLegendPositionLeft)
						{
							if (position != XlLegendPosition.xlLegendPositionBottom)
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
								if (position != XlLegendPosition.xlLegendPositionCorner)
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
									goto IL_00e5;
								}
								text = VH.A(172416);
							}
							else
							{
								text = VH.A(57416);
							}
						}
						else
						{
							text = VH.A(56582);
						}
						goto IL_013c;
					}
					IL_013c:
					legend = default(Properties.LegendProperties);
					txtLegendPosn.Text = text;
					txtLegendFormat.Text = VH.A(172459);
					return;
					IL_00e5:
					if (B)
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
						text2 = VH.A(172429);
					}
					else
					{
						text2 = VH.A(172444);
					}
					text = legend.Left.ToString(text2) + VH.A(10524) + legend.Top.ToString(text2);
					goto IL_013c;
				}
			}
		}
		txtLegendPosn.Text = VH.A(13778);
		txtLegendFormat.Text = VH.A(13778);
	}

	private void A(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[3] { chkLegend, chkLegendFormat, chkLegendPosn }));
	}

	private void B(Properties A, bool B)
	{
		if (A.Chart.HasTitle)
		{
			Properties.TitleProperties title = A.Title;
			XlChartElementPosition position = title.Position;
			string text;
			if (position == XlChartElementPosition.xlChartElementPositionAutomatic)
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
				text = VH.A(172496);
			}
			else
			{
				string text2;
				if (B)
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
					text2 = VH.A(172429);
				}
				else
				{
					text2 = VH.A(172444);
				}
				text = title.Left.ToString(text2) + VH.A(10524) + title.Top.ToString(text2);
			}
			title = default(Properties.TitleProperties);
			txtTitlePosn.Text = text;
			txtTitleFormat.Text = VH.A(172459);
		}
		else
		{
			txtTitlePosn.Text = VH.A(13778);
			txtTitleFormat.Text = VH.A(13778);
		}
	}

	private void B(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[3] { chkTitle, chkTitleFormat, chkTitlePosn }));
	}

	private void D(Properties A)
	{
		if (A.Chart.HasDataTable)
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
					txtDataTableFormat.Text = VH.A(172459);
					return;
				}
			}
		}
		txtDataTableFormat.Text = VH.A(13778);
	}

	private void C(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[2] { chkDataTable, chkDataTableFormat }));
	}

	private void E(Properties A)
	{
		this.A(A.PrimaryValueAxis, A.Chart.HasPrimaryValueAxis, txtPriValScale, txtPriValTicks, txtPriValGridlines, txtPriValTitle);
	}

	private void D(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[5] { chkPriVal, chkPriValScale, chkPriValGridlines, chkPriValTicks, chkPriValTitle }));
	}

	private void F(Properties A)
	{
		this.A(A.PrimaryCategoryAxis, A.Chart.HasPrimaryCategoryAxis, txtPriCatScale, txtPriCatTicks, txtPriCatGridlines, txtPriCatTitle);
	}

	private void E(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[5] { chkPriCat, chkPriCatScale, chkPriCatGridlines, chkPriCatTicks, chkPriCatTitle }));
	}

	private void G(Properties A)
	{
		this.A(A.SecondaryValueAxis, A.Chart.HasSecondaryValueAxis, txtSecValScale, txtSecValTicks, txtSecValGridlines, txtSecValTitle);
	}

	private void F(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[5] { chkSecVal, chkSecValScale, chkSecValGridlines, chkSecValTicks, chkSecValTitle }));
	}

	private void H(Properties A)
	{
		this.A(A.SecondaryCategoryAxis, A.Chart.HasSecondaryCategoryAxis, txtSecCatScale, txtSecCatTicks, txtSecCatGridlines, txtSecCatTitle);
	}

	private void G(bool A)
	{
		this.A(A, new List<System.Windows.Controls.CheckBox>(new System.Windows.Controls.CheckBox[5] { chkSecCat, chkSecCatScale, chkSecCatGridlines, chkSecCatTicks, chkSecCatTitle }));
	}

	private void A(bool A, List<System.Windows.Controls.CheckBox> B)
	{
		using (List<System.Windows.Controls.CheckBox>.Enumerator enumerator = B.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				enumerator.Current.IsEnabled = A;
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
				break;
			}
		}
		if (A)
		{
			return;
		}
		foreach (System.Windows.Controls.CheckBox item in B)
		{
			item.IsChecked = false;
		}
	}

	private void A(Properties.AxisProperties A, bool B, TextBlock C, TextBlock D, TextBlock E, TextBlock F)
	{
		if (B)
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
					C.Text = this.A(A);
					D.Text = this.B(A);
					E.Text = this.C(A);
					F.Text = this.D(A);
					return;
				}
			}
		}
		C.Text = VH.A(13778);
		D.Text = VH.A(13778);
		E.Text = VH.A(13778);
		F.Text = VH.A(13778);
	}

	private string A(Properties.AxisProperties A)
	{
		string text;
		if (A.MaximumScaleIsAuto)
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
			text = VH.A(172496);
		}
		else
		{
			text = A.MaximumScale.ToString();
		}
		text += VH.A(172505);
		if (A.MinimumScaleIsAuto)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return text + VH.A(172496);
				}
			}
		}
		return text + A.MinimumScale;
	}

	private string B(Properties.AxisProperties A)
	{
		string text = "";
		text += this.A(A.MajorTickMark);
		text += VH.A(172505);
		text += this.A(A.MinorTickMark);
		text += VH.A(172505);
		XlTickLabelPosition tickLabelPosition = A.TickLabelPosition;
		if (tickLabelPosition != XlTickLabelPosition.xlTickLabelPositionLow)
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
					if (tickLabelPosition != XlTickLabelPosition.xlTickLabelPositionHigh)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								if (tickLabelPosition == XlTickLabelPosition.xlTickLabelPositionNextToAxis)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											return text + VH.A(172512);
										}
									}
								}
								return text + VH.A(61174);
							}
						}
					}
					return text + VH.A(81380);
				}
			}
		}
		return text + VH.A(81373);
	}

	private string A(XlTickMark A)
	{
		if (A != XlTickMark.xlTickMarkNone)
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
					if (A != XlTickMark.xlTickMarkInside)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								if (A != XlTickMark.xlTickMarkOutside)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											return VH.A(146128);
										}
									}
								}
								return VH.A(172529);
							}
						}
					}
					return VH.A(172544);
				}
			}
		}
		return VH.A(61174);
	}

	private string C(Properties.AxisProperties A)
	{
		if (A.HasMajorGridlines)
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
			if (A.HasMinorGridlines)
			{
				return VH.A(172557);
			}
		}
		if (A.HasMajorGridlines)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return VH.A(172582);
				}
			}
		}
		if (A.HasMinorGridlines)
		{
			return VH.A(172582);
		}
		return VH.A(61174);
	}

	private string D(Properties.AxisProperties A)
	{
		if (A.HasTitle)
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
					return VH.A(172593);
				}
			}
		}
		return VH.A(172646);
	}

	private string A(string A)
	{
		return Regex.Replace(A, VH.A(172663), VH.A(172678), RegexOptions.Compiled).Trim();
	}

	private void A(System.Windows.Shapes.Rectangle A, System.Drawing.Color B)
	{
		System.Windows.Shapes.Rectangle rectangle = A;
		rectangle.Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(B.R, B.G, B.B));
		if (this.A(B) < 180)
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
			rectangle.StrokeThickness = 0.0;
		}
		else
		{
			rectangle.StrokeThickness = 1.0;
		}
		rectangle = null;
	}

	private int A(System.Drawing.Color A)
	{
		return checked((int)Math.Round(Math.Sqrt((double)(A.R * A.R) * 0.299 + (double)(A.G * A.G) * 0.587 + (double)(A.B * A.B) * 0.114)));
	}

	private string A(System.Drawing.Color A)
	{
		return clsColors.Color2RGB(A).Replace(VH.A(2378), VH.A(10524));
	}

	private void B()
	{
		chkChartSize.Checked += ChartChildCheckChanged;
		chkChartSize.Unchecked += ChartChildCheckChanged;
		chkChartTop.Checked += ChartChildCheckChanged;
		chkChartTop.Unchecked += ChartChildCheckChanged;
		chkChartLeft.Checked += ChartChildCheckChanged;
		chkChartLeft.Unchecked += ChartChildCheckChanged;
		chkChartFormat.Checked += ChartChildCheckChanged;
		chkChartFormat.Unchecked += ChartChildCheckChanged;
	}

	private void C()
	{
		chkChartSize.Checked -= ChartChildCheckChanged;
		chkChartSize.Unchecked -= ChartChildCheckChanged;
		chkChartTop.Checked -= ChartChildCheckChanged;
		chkChartTop.Unchecked -= ChartChildCheckChanged;
		chkChartLeft.Checked -= ChartChildCheckChanged;
		chkChartLeft.Unchecked -= ChartChildCheckChanged;
		chkChartFormat.Checked -= ChartChildCheckChanged;
		chkChartFormat.Unchecked -= ChartChildCheckChanged;
	}

	private void ChartCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkChart.IsChecked.Value;
		C();
		chkChartSize.IsChecked = value;
		chkChartTop.IsChecked = value;
		chkChartLeft.IsChecked = value;
		chkChartFormat.IsChecked = value;
		B();
	}

	private void ChartChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkChartSize.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			if (isChecked != true)
			{
				obj = false;
				goto IL_0095;
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
		}
		bool? isChecked2;
		flag = (isChecked2 = chkChartTop.IsChecked);
		if (!flag.HasValue)
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0095;
		IL_0270:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		isChecked = (bool?)obj2;
		bool? obj3;
		if (isChecked.HasValue)
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
			if (flag2 != true)
			{
				obj3 = false;
				goto IL_0314;
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
		}
		isChecked = chkChartLeft.IsChecked;
		bool? flag3;
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
			flag3 = isChecked;
		}
		else
		{
			flag3 = isChecked != true;
		}
		bool? flag4 = flag3;
		isChecked = flag3;
		if (!isChecked.HasValue)
		{
			obj3 = null;
		}
		else if (flag4 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag2;
		}
		goto IL_0314;
		IL_03a7:
		chkChart.IsChecked = null;
		return;
		IL_0111:
		bool? obj4;
		bool? flag5 = (bool?)obj4;
		if (flag5.HasValue)
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
			if (flag5 != true)
			{
				goto IL_0186;
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
		}
		if (chkChartFormat.IsChecked == true)
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
			if (flag5.HasValue)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						chkChart.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_0186;
		IL_0095:
		flag4 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
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
			if (flag4 != true)
			{
				obj4 = false;
				goto IL_0111;
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
		isChecked2 = (flag2 = chkChartLeft.IsChecked);
		if (!isChecked2.HasValue)
		{
			obj4 = null;
		}
		else if (flag2 != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = flag4;
		}
		goto IL_0111;
		IL_0186:
		flag = chkChartSize.IsChecked;
		flag = (isChecked2 = (!flag) ?? flag);
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj2 = false;
				goto IL_0270;
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
		}
		flag = chkChartTop.IsChecked;
		bool? flag6;
		if (!flag.HasValue)
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
			flag6 = flag;
		}
		else
		{
			flag6 = flag != true;
		}
		isChecked = flag6;
		flag = flag6;
		if (!flag.HasValue)
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
		else if (isChecked != true)
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
			obj2 = false;
		}
		else
		{
			obj2 = isChecked2;
		}
		goto IL_0270;
		IL_0314:
		flag5 = obj3;
		if (flag5.HasValue)
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
			if (flag5 != true)
			{
				goto IL_03a7;
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
		flag4 = chkChartFormat.IsChecked;
		bool? flag7;
		if (!flag4.HasValue)
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
			flag7 = flag4;
		}
		else
		{
			flag7 = flag4 != true;
		}
		flag4 = flag7;
		if (flag4 == true)
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
			if (flag5.HasValue)
			{
				chkChart.IsChecked = false;
				return;
			}
		}
		goto IL_03a7;
	}

	private void D()
	{
		chkPlotSize.Checked += PlotAreaChildCheckChanged;
		chkPlotSize.Unchecked += PlotAreaChildCheckChanged;
		chkPlotLocation.Checked += PlotAreaChildCheckChanged;
		chkPlotLocation.Unchecked += PlotAreaChildCheckChanged;
		chkPlotFormat.Checked += PlotAreaChildCheckChanged;
		chkPlotFormat.Unchecked += PlotAreaChildCheckChanged;
	}

	private void E()
	{
		chkPlotSize.Checked -= PlotAreaChildCheckChanged;
		chkPlotSize.Unchecked -= PlotAreaChildCheckChanged;
		chkPlotLocation.Checked -= PlotAreaChildCheckChanged;
		chkPlotLocation.Unchecked -= PlotAreaChildCheckChanged;
		chkPlotFormat.Checked -= PlotAreaChildCheckChanged;
		chkPlotFormat.Unchecked -= PlotAreaChildCheckChanged;
	}

	private void PlotAreaCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkPlotArea.IsChecked.Value;
		E();
		chkPlotSize.IsChecked = value;
		chkPlotLocation.IsChecked = value;
		chkPlotFormat.IsChecked = value;
		D();
	}

	private void PlotAreaChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkPlotSize.IsChecked);
		bool? obj;
		if (flag.HasValue)
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_0095;
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
		bool? isChecked2;
		flag = (isChecked2 = chkPlotLocation.IsChecked);
		if (!flag.HasValue)
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0095;
		IL_026c:
		chkPlotArea.IsChecked = null;
		return;
		IL_010c:
		flag = chkPlotSize.IsChecked;
		bool? flag2;
		if (!flag.HasValue)
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
			flag2 = flag;
		}
		else
		{
			flag2 = flag != true;
		}
		isChecked2 = flag2;
		flag = flag2;
		bool? obj2;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj2 = false;
				goto IL_01e1;
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
		}
		flag = chkPlotLocation.IsChecked;
		bool? flag3;
		if (!flag.HasValue)
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
			flag3 = flag;
		}
		else
		{
			flag3 = flag != true;
		}
		isChecked = flag3;
		flag = flag3;
		if (flag.HasValue)
		{
			obj2 = (isChecked == true) & isChecked2;
		}
		else
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
			obj2 = null;
		}
		goto IL_01e1;
		IL_01e1:
		bool? flag4 = obj2;
		if (flag4.HasValue)
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
			if (flag4 != true)
			{
				goto IL_026c;
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
		}
		isChecked = chkPlotFormat.IsChecked;
		bool? flag5;
		if (!isChecked.HasValue)
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
			flag5 = isChecked;
		}
		else
		{
			flag5 = isChecked != true;
		}
		isChecked = flag5;
		if (isChecked == true && flag4.HasValue)
		{
			chkPlotArea.IsChecked = false;
			return;
		}
		goto IL_026c;
		IL_0095:
		flag4 = obj;
		if (flag4.HasValue)
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
			if (flag4 != true)
			{
				goto IL_010c;
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
		}
		if (chkPlotFormat.IsChecked == true)
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
			if (flag4.HasValue)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						chkPlotArea.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_010c;
	}

	private void F()
	{
		chkSeriesFormat.Checked += SeriesChildCheckChanged;
		chkSeriesFormat.Unchecked += SeriesChildCheckChanged;
		chkSeriesGap.Checked += SeriesChildCheckChanged;
		chkSeriesGap.Unchecked += SeriesChildCheckChanged;
		chkSeriesAngle.Checked += SeriesChildCheckChanged;
		chkSeriesAngle.Unchecked += SeriesChildCheckChanged;
		chkSeriesExplosion.Checked += SeriesChildCheckChanged;
		chkSeriesExplosion.Unchecked += SeriesChildCheckChanged;
		chkSeriesDataLabels.Checked += SeriesChildCheckChanged;
		chkSeriesDataLabels.Unchecked += SeriesChildCheckChanged;
		chkSeriesErrorBars.Checked += SeriesChildCheckChanged;
		chkSeriesErrorBars.Unchecked += SeriesChildCheckChanged;
		chkSeriesUpDownBars.Checked += SeriesChildCheckChanged;
		chkSeriesUpDownBars.Unchecked += SeriesChildCheckChanged;
	}

	private void G()
	{
		chkSeriesFormat.Checked -= SeriesChildCheckChanged;
		chkSeriesFormat.Unchecked -= SeriesChildCheckChanged;
		chkSeriesGap.Checked -= SeriesChildCheckChanged;
		chkSeriesGap.Unchecked -= SeriesChildCheckChanged;
		chkSeriesAngle.Checked -= SeriesChildCheckChanged;
		chkSeriesAngle.Unchecked -= SeriesChildCheckChanged;
		chkSeriesExplosion.Checked -= SeriesChildCheckChanged;
		chkSeriesExplosion.Unchecked -= SeriesChildCheckChanged;
		chkSeriesDataLabels.Checked -= SeriesChildCheckChanged;
		chkSeriesDataLabels.Unchecked -= SeriesChildCheckChanged;
		chkSeriesErrorBars.Checked -= SeriesChildCheckChanged;
		chkSeriesErrorBars.Unchecked -= SeriesChildCheckChanged;
		chkSeriesUpDownBars.Checked -= SeriesChildCheckChanged;
		chkSeriesUpDownBars.Unchecked -= SeriesChildCheckChanged;
	}

	private void SeriesCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkSeries.IsChecked.Value;
		G();
		chkSeriesFormat.IsChecked = value;
		chkSeriesGap.IsChecked = value && chkSeriesGap.IsEnabled;
		chkSeriesAngle.IsChecked = value && chkSeriesAngle.IsEnabled;
		System.Windows.Controls.CheckBox checkBox = chkSeriesExplosion;
		int value2;
		if (value)
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
			value2 = (chkSeriesExplosion.IsEnabled ? 1 : 0);
		}
		else
		{
			value2 = 0;
		}
		checkBox.IsChecked = (byte)value2 != 0;
		System.Windows.Controls.CheckBox checkBox2 = chkSeriesDataLabels;
		int value3;
		if (value)
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
			value3 = (chkSeriesDataLabels.IsEnabled ? 1 : 0);
		}
		else
		{
			value3 = 0;
		}
		checkBox2.IsChecked = (byte)value3 != 0;
		chkSeriesErrorBars.IsChecked = value && chkSeriesErrorBars.IsEnabled;
		System.Windows.Controls.CheckBox checkBox3 = chkSeriesUpDownBars;
		int value4;
		if (value)
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
			value4 = (chkSeriesUpDownBars.IsEnabled ? 1 : 0);
		}
		else
		{
			value4 = 0;
		}
		checkBox3.IsChecked = (byte)value4 != 0;
		F();
	}

	private void SeriesChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkSeriesFormat.IsChecked);
		bool? obj;
		if (flag.HasValue)
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_00f0;
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
		bool? flag2 = (flag = chkSeriesGap.IsChecked);
		bool? flag3;
		if (flag2.HasValue)
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
			if (flag == true)
			{
				flag3 = true;
				goto IL_00a3;
			}
		}
		if (chkSeriesGap.IsEnabled)
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
			flag3 = flag;
		}
		else
		{
			flag3 = true;
		}
		goto IL_00a3;
		IL_06a4:
		bool? flag5;
		bool? flag4 = (bool?)flag5;
		isChecked = (bool?)flag5;
		bool? obj2;
		bool? flag6;
		if (isChecked.HasValue)
		{
			obj2 = (flag4 == true) & flag6;
		}
		else
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
			obj2 = null;
		}
		goto IL_06e5;
		IL_036c:
		bool? obj3;
		bool? flag7 = (bool?)obj3;
		bool? flag8 = (bool?)obj3;
		bool? obj4;
		if (flag8.HasValue)
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
			if (flag7 != true)
			{
				obj4 = false;
				goto IL_0436;
			}
		}
		bool? flag9 = (flag8 = chkSeriesDataLabels.IsChecked);
		bool? flag10;
		if (flag9.HasValue)
		{
			if (flag8 == true)
			{
				flag10 = true;
				goto IL_03ed;
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
		if (chkSeriesDataLabels.IsEnabled)
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
			flag10 = flag8;
		}
		else
		{
			flag10 = true;
		}
		goto IL_03ed;
		IL_06e5:
		bool? flag11 = obj2;
		flag4 = obj2;
		bool? obj5;
		if (flag4.HasValue)
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
			if (flag11 != true)
			{
				obj5 = false;
				goto IL_07f7;
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
		flag6 = chkSeriesExplosion.IsChecked;
		bool? flag12;
		if (!flag6.HasValue)
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
			flag12 = flag6;
		}
		else
		{
			flag12 = flag6 != true;
		}
		flag4 = flag12;
		flag6 = flag12;
		bool? flag13;
		if (flag6.HasValue)
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
			if (flag4 == true)
			{
				flag13 = true;
				goto IL_07aa;
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
		if (chkSeriesExplosion.IsEnabled)
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
			flag13 = flag4;
		}
		else
		{
			flag13 = true;
		}
		goto IL_07aa;
		IL_0a93:
		chkSeries.IsChecked = null;
		return;
		IL_04bc:
		flag = chkSeriesFormat.IsChecked;
		bool? flag14;
		flag = (flag14 = (!flag) ?? flag);
		bool? obj6;
		if (flag.HasValue)
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
			if (flag14 != true)
			{
				obj6 = false;
				goto IL_05e9;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		flag2 = chkSeriesGap.IsChecked;
		bool? flag15;
		if (!flag2.HasValue)
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
			flag15 = flag2;
		}
		else
		{
			flag15 = flag2 != true;
		}
		flag = flag15;
		flag2 = flag15;
		bool? obj7;
		if (flag2.HasValue)
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
			if (flag == true)
			{
				obj7 = true;
				goto IL_05a6;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		obj7 = ((!chkSeriesGap.IsEnabled) ? new bool?(true) : flag);
		goto IL_05a6;
		IL_05e9:
		flag6 = obj6;
		isChecked = obj6;
		if (isChecked.HasValue)
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
			if (flag6 != true)
			{
				obj2 = false;
				goto IL_06e5;
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
		}
		flag14 = chkSeriesAngle.IsChecked;
		flag14 = (isChecked = (!flag14) ?? flag14);
		if (flag14.HasValue)
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
			if (isChecked == true)
			{
				flag5 = true;
				goto IL_06a4;
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
		}
		if (chkSeriesAngle.IsEnabled)
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
			flag5 = isChecked;
		}
		else
		{
			flag5 = true;
		}
		goto IL_06a4;
		IL_0999:
		bool? obj8;
		flag7 = (bool?)obj8;
		flag9 = (bool?)obj8;
		bool? obj9;
		bool? flag16;
		if (!flag9.HasValue)
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
			obj9 = null;
		}
		else if (flag7 != true)
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
			obj9 = false;
		}
		else
		{
			obj9 = flag16;
		}
		goto IL_09e0;
		IL_09e0:
		bool? flag17 = obj9;
		if (flag17.HasValue)
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
			if (flag17 != true)
			{
				goto IL_0a93;
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
		}
		flag7 = chkSeriesUpDownBars.IsChecked;
		bool? flag18;
		if (!flag7.HasValue)
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
			flag18 = flag7;
		}
		else
		{
			flag18 = flag7 != true;
		}
		flag7 = flag18;
		if (flag7 != true)
		{
			if (chkSeriesUpDownBars.IsEnabled)
			{
				goto IL_0a93;
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
		if (flag17.HasValue)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkSeries.IsChecked = false;
					return;
				}
			}
		}
		goto IL_0a93;
		IL_07aa:
		bool? flag19 = flag13;
		flag4 = flag13;
		if (!flag4.HasValue)
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
			obj5 = null;
		}
		else if (flag19 != true)
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
			obj5 = false;
		}
		else
		{
			obj5 = flag11;
		}
		goto IL_07f7;
		IL_05a6:
		isChecked = obj7;
		flag = obj7;
		if (flag.HasValue)
		{
			obj6 = (isChecked == true) & flag14;
		}
		else
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
			obj6 = null;
		}
		goto IL_05e9;
		IL_08f4:
		bool? obj10;
		flag16 = (bool?)obj10;
		flag9 = (bool?)obj10;
		if (flag9.HasValue)
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
			if (flag16 != true)
			{
				obj9 = false;
				goto IL_09e0;
			}
		}
		flag8 = chkSeriesDataLabels.IsChecked;
		bool? flag20;
		if (!flag8.HasValue)
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
			flag20 = flag8;
		}
		else
		{
			flag20 = flag8 != true;
		}
		flag9 = flag20;
		flag8 = flag20;
		if (flag8.HasValue)
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
			if (flag9 == true)
			{
				obj8 = true;
				goto IL_0999;
			}
		}
		obj8 = ((!chkSeriesDataLabels.IsEnabled) ? new bool?(true) : flag9);
		goto IL_0999;
		IL_0436:
		flag17 = obj4;
		if (flag17.HasValue)
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
			if (flag17 != true)
			{
				goto IL_04bc;
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
		if (chkSeriesUpDownBars.IsChecked != true)
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
			if (chkSeriesUpDownBars.IsEnabled)
			{
				goto IL_04bc;
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
		if (flag17.HasValue)
		{
			chkSeries.IsChecked = true;
			return;
		}
		goto IL_04bc;
		IL_0171:
		bool? flag21;
		flag6 = (bool?)flag21;
		flag14 = (bool?)flag21;
		bool? obj11;
		if (!flag14.HasValue)
		{
			obj11 = null;
		}
		else if (flag6 != true)
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
			obj11 = false;
		}
		else
		{
			obj11 = flag4;
		}
		goto IL_01b4;
		IL_08b2:
		bool? obj12;
		flag9 = (bool?)obj12;
		flag19 = (bool?)obj12;
		if (!flag19.HasValue)
		{
			obj10 = null;
		}
		else if (flag9 != true)
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
			obj10 = false;
		}
		else
		{
			obj10 = flag8;
		}
		goto IL_08f4;
		IL_00a3:
		flag14 = flag3;
		flag = flag3;
		if (!flag.HasValue)
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
			obj = null;
		}
		else if (flag14 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_00f0;
		IL_03ed:
		flag16 = flag10;
		flag8 = flag10;
		if (!flag8.HasValue)
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
			obj4 = null;
		}
		else if (flag16 != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = flag7;
		}
		goto IL_0436;
		IL_0324:
		bool? flag22;
		flag8 = (bool?)flag22;
		flag11 = (bool?)flag22;
		if (!flag11.HasValue)
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
			obj3 = null;
		}
		else if (flag8 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag9;
		}
		goto IL_036c;
		IL_01b4:
		flag19 = obj11;
		flag6 = obj11;
		bool? obj13;
		if (flag6.HasValue)
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
			if (flag19 != true)
			{
				obj13 = false;
				goto IL_0298;
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
		flag4 = (flag6 = chkSeriesExplosion.IsChecked);
		bool? flag23;
		if (flag4.HasValue)
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
			if (flag6 == true)
			{
				flag23 = true;
				goto IL_024d;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (chkSeriesExplosion.IsEnabled)
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
			flag23 = flag6;
		}
		else
		{
			flag23 = true;
		}
		goto IL_024d;
		IL_0298:
		flag9 = obj13;
		flag11 = obj13;
		if (flag11.HasValue)
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
			if (flag9 != true)
			{
				obj3 = false;
				goto IL_036c;
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
		flag19 = (flag11 = chkSeriesErrorBars.IsChecked);
		if (flag19.HasValue)
		{
			if (flag11 == true)
			{
				flag22 = true;
				goto IL_0324;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (chkSeriesErrorBars.IsEnabled)
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
			flag22 = flag11;
		}
		else
		{
			flag22 = true;
		}
		goto IL_0324;
		IL_07f7:
		flag8 = obj5;
		flag19 = obj5;
		if (flag19.HasValue)
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
			if (flag8 != true)
			{
				obj10 = false;
				goto IL_08f4;
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
		}
		flag11 = chkSeriesErrorBars.IsChecked;
		bool? flag24;
		if (!flag11.HasValue)
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
			flag24 = flag11;
		}
		else
		{
			flag24 = flag11 != true;
		}
		flag19 = flag24;
		flag11 = flag24;
		if (flag11.HasValue)
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
			if (flag19 == true)
			{
				obj12 = true;
				goto IL_08b2;
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
		}
		obj12 = ((!chkSeriesErrorBars.IsEnabled) ? new bool?(true) : flag19);
		goto IL_08b2;
		IL_024d:
		flag11 = flag23;
		flag6 = flag23;
		if (!flag6.HasValue)
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
			obj13 = null;
		}
		else if (flag11 != true)
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
			obj13 = false;
		}
		else
		{
			obj13 = flag19;
		}
		goto IL_0298;
		IL_00f0:
		flag4 = obj;
		flag14 = obj;
		if (flag14.HasValue)
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
			if (flag4 != true)
			{
				obj11 = false;
				goto IL_01b4;
			}
		}
		isChecked = (flag14 = chkSeriesAngle.IsChecked);
		if (isChecked.HasValue)
		{
			if (flag14 == true)
			{
				flag21 = true;
				goto IL_0171;
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
		if (chkSeriesAngle.IsEnabled)
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
			flag21 = flag14;
		}
		else
		{
			flag21 = true;
		}
		goto IL_0171;
	}

	private void H()
	{
		chkPriVal.Checked += AxesChildCheckChanged;
		chkPriVal.Unchecked += AxesChildCheckChanged;
		chkPriVal.Indeterminate += AxesChildCheckChanged;
		chkPriCat.Checked += AxesChildCheckChanged;
		chkPriCat.Unchecked += AxesChildCheckChanged;
		chkPriCat.Indeterminate += AxesChildCheckChanged;
		chkSecVal.Checked += AxesChildCheckChanged;
		chkSecVal.Unchecked += AxesChildCheckChanged;
		chkSecVal.Indeterminate += AxesChildCheckChanged;
		chkSecCat.Checked += AxesChildCheckChanged;
		chkSecCat.Unchecked += AxesChildCheckChanged;
		chkSecCat.Indeterminate += AxesChildCheckChanged;
	}

	private void I()
	{
		chkPriVal.Checked -= AxesChildCheckChanged;
		chkPriVal.Unchecked -= AxesChildCheckChanged;
		chkPriVal.Indeterminate -= AxesChildCheckChanged;
		chkPriCat.Checked -= AxesChildCheckChanged;
		chkPriCat.Unchecked -= AxesChildCheckChanged;
		chkPriCat.Indeterminate -= AxesChildCheckChanged;
		chkSecVal.Checked -= AxesChildCheckChanged;
		chkSecVal.Unchecked -= AxesChildCheckChanged;
		chkSecVal.Indeterminate -= AxesChildCheckChanged;
		chkSecCat.Checked -= AxesChildCheckChanged;
		chkSecCat.Unchecked -= AxesChildCheckChanged;
		chkSecCat.Indeterminate -= AxesChildCheckChanged;
	}

	private void J()
	{
		chkAxes.Checked += AxesCheckedChanged;
		chkAxes.Unchecked += AxesCheckedChanged;
	}

	private void K()
	{
		chkAxes.Checked -= AxesCheckedChanged;
		chkAxes.Unchecked -= AxesCheckedChanged;
	}

	private void AxesCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkAxes.IsChecked.Value;
		I();
		System.Windows.Controls.CheckBox checkBox = chkPriVal;
		int value2;
		if (value)
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
			value2 = (chkPriVal.IsEnabled ? 1 : 0);
		}
		else
		{
			value2 = 0;
		}
		checkBox.IsChecked = (byte)value2 != 0;
		chkPriCat.IsChecked = value && chkPriCat.IsEnabled;
		chkSecVal.IsChecked = value && chkSecVal.IsEnabled;
		System.Windows.Controls.CheckBox checkBox2 = chkSecCat;
		int value3;
		if (value)
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
			value3 = (chkSecCat.IsEnabled ? 1 : 0);
		}
		else
		{
			value3 = 0;
		}
		checkBox2.IsChecked = (byte)value3 != 0;
		H();
	}

	private void AxesChildCheckChanged(object sender, RoutedEventArgs e)
	{
		K();
		bool? isChecked;
		bool? flag = (isChecked = chkPriVal.IsChecked);
		bool? flag2;
		if (!flag.HasValue || isChecked != true)
		{
			if (chkPriVal.IsEnabled)
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
				flag2 = isChecked;
			}
			else
			{
				flag2 = true;
			}
		}
		else
		{
			flag2 = true;
		}
		bool? flag3 = flag2;
		isChecked = flag2;
		bool? obj;
		if (isChecked.HasValue)
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
			if (flag3 != true)
			{
				obj = false;
				goto IL_013c;
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
		}
		flag = (isChecked = chkPriCat.IsChecked);
		bool? flag4;
		if (flag.HasValue)
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
			if (isChecked == true)
			{
				flag4 = true;
				goto IL_00fe;
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
		}
		if (chkPriCat.IsEnabled)
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
			flag4 = isChecked;
		}
		else
		{
			flag4 = true;
		}
		goto IL_00fe;
		IL_00fe:
		bool? flag5 = flag4;
		isChecked = flag4;
		if (isChecked.HasValue)
		{
			obj = (flag5 == true) & flag3;
		}
		else
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
			obj = null;
		}
		goto IL_013c;
		IL_0274:
		flag = chkPriVal.IsChecked;
		bool? flag6;
		if (!flag.HasValue)
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
			flag6 = flag;
		}
		else
		{
			flag6 = flag != true;
		}
		isChecked = flag6;
		flag = flag6;
		bool? flag7;
		if (flag.HasValue)
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
			if (isChecked == true)
			{
				flag7 = true;
				goto IL_0305;
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
		if (chkPriVal.IsEnabled)
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
			flag7 = isChecked;
		}
		else
		{
			flag7 = true;
		}
		goto IL_0305;
		IL_0410:
		bool? obj2;
		bool? flag8 = (bool?)obj2;
		flag3 = (bool?)obj2;
		bool? flag9;
		if (!flag3.HasValue || flag8 == true)
		{
			flag5 = chkSecVal.IsChecked;
			flag5 = (flag3 = (!flag5) ?? flag5);
			if (flag5.HasValue)
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
				if (flag3 == true)
				{
					flag9 = true;
					goto IL_04ab;
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
			}
			if (chkSecVal.IsEnabled)
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
				flag9 = flag3;
			}
			else
			{
				flag9 = true;
			}
			goto IL_04ab;
		}
		bool? obj3 = false;
		goto IL_04e8;
		IL_03c4:
		bool? flag10;
		flag3 = (bool?)flag10;
		isChecked = (bool?)flag10;
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
		else if (flag3 != true)
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
			obj2 = false;
		}
		else
		{
			obj2 = flag5;
		}
		goto IL_0410;
		IL_0305:
		flag5 = flag7;
		isChecked = flag7;
		if (isChecked.HasValue)
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
			if (flag5 != true)
			{
				obj2 = false;
				goto IL_0410;
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
		flag = chkPriCat.IsChecked;
		bool? flag11;
		if (!flag.HasValue)
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
			flag11 = flag;
		}
		else
		{
			flag11 = flag != true;
		}
		isChecked = flag11;
		flag = flag11;
		if (flag.HasValue)
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
			if (isChecked == true)
			{
				flag10 = true;
				goto IL_03c4;
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
		if (chkPriCat.IsEnabled)
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
			flag10 = isChecked;
		}
		else
		{
			flag10 = true;
		}
		goto IL_03c4;
		IL_013c:
		bool? flag12 = obj;
		flag5 = obj;
		bool? obj4;
		if (!flag5.HasValue || flag12 == true)
		{
			flag3 = (flag5 = chkSecVal.IsChecked);
			bool? flag13;
			if (!flag3.HasValue || flag5 != true)
			{
				if (chkSecVal.IsEnabled)
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
					flag13 = flag5;
				}
				else
				{
					flag13 = true;
				}
			}
			else
			{
				flag13 = true;
			}
			flag8 = flag13;
			flag5 = flag13;
			if (flag5.HasValue)
			{
				obj4 = (flag8 == true) & flag12;
			}
			else
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
				obj4 = null;
			}
		}
		else
		{
			obj4 = false;
		}
		bool? flag14 = obj4;
		if (flag14.HasValue)
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
			if (flag14 != true)
			{
				goto IL_0274;
			}
		}
		if (chkSecCat.IsChecked != true)
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
			if (chkSecCat.IsEnabled)
			{
				goto IL_0274;
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
		if (!flag14.HasValue)
		{
			goto IL_0274;
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
		chkAxes.IsChecked = true;
		goto IL_05ae;
		IL_04ab:
		flag12 = flag9;
		flag3 = flag9;
		if (!flag3.HasValue)
		{
			obj3 = null;
		}
		else if (flag12 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag8;
		}
		goto IL_04e8;
		IL_05ae:
		J();
		return;
		IL_04e8:
		flag14 = obj3;
		if (flag14.HasValue)
		{
			if (flag14 != true)
			{
				goto IL_059a;
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
		flag12 = chkSecCat.IsChecked;
		bool? flag15;
		if (!flag12.HasValue)
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
			flag15 = flag12;
		}
		else
		{
			flag15 = flag12 != true;
		}
		flag12 = flag15;
		if (flag12 != true)
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
			if (chkSecCat.IsEnabled)
			{
				goto IL_059a;
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
		if (!flag14.HasValue)
		{
			goto IL_059a;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			break;
		}
		chkAxes.IsChecked = false;
		goto IL_05ae;
		IL_059a:
		chkAxes.IsChecked = null;
		goto IL_05ae;
	}

	private void L()
	{
		chkPriValScale.Checked += PriValAxisChildCheckChanged;
		chkPriValScale.Unchecked += PriValAxisChildCheckChanged;
		chkPriValGridlines.Checked += PriValAxisChildCheckChanged;
		chkPriValGridlines.Unchecked += PriValAxisChildCheckChanged;
		chkPriValTicks.Checked += PriValAxisChildCheckChanged;
		chkPriValTicks.Unchecked += PriValAxisChildCheckChanged;
		chkPriValTitle.Checked += PriValAxisChildCheckChanged;
		chkPriValTitle.Unchecked += PriValAxisChildCheckChanged;
	}

	private void M()
	{
		chkPriValScale.Checked -= PriValAxisChildCheckChanged;
		chkPriValScale.Unchecked -= PriValAxisChildCheckChanged;
		chkPriValGridlines.Checked -= PriValAxisChildCheckChanged;
		chkPriValGridlines.Unchecked -= PriValAxisChildCheckChanged;
		chkPriValTicks.Checked -= PriValAxisChildCheckChanged;
		chkPriValTicks.Unchecked -= PriValAxisChildCheckChanged;
		chkPriValTitle.Checked -= PriValAxisChildCheckChanged;
		chkPriValTitle.Unchecked -= PriValAxisChildCheckChanged;
	}

	private void PriValAxisCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkPriVal.IsChecked.Value;
		M();
		chkPriValScale.IsChecked = value;
		chkPriValGridlines.IsChecked = value;
		chkPriValTicks.IsChecked = value;
		chkPriValTitle.IsChecked = value;
		L();
	}

	private void PriValAxisChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkPriValScale.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			if (isChecked != true)
			{
				obj = false;
				goto IL_0089;
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
		}
		bool? isChecked2;
		flag = (isChecked2 = chkPriValGridlines.IsChecked);
		if (!flag.HasValue)
		{
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0089;
		IL_0228:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		isChecked = (bool?)obj2;
		bool? obj3;
		if (isChecked.HasValue)
		{
			if (flag2 != true)
			{
				obj3 = false;
				goto IL_02c6;
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
		}
		isChecked = chkPriValTicks.IsChecked;
		bool? flag3;
		if (!isChecked.HasValue)
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
			flag3 = isChecked;
		}
		else
		{
			flag3 = isChecked != true;
		}
		bool? flag4 = flag3;
		isChecked = flag3;
		if (!isChecked.HasValue)
		{
			obj3 = null;
		}
		else if (flag4 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag2;
		}
		goto IL_02c6;
		IL_0354:
		chkPriVal.IsChecked = null;
		return;
		IL_02c6:
		bool? flag5 = obj3;
		if (flag5.HasValue)
		{
			if (flag5 != true)
			{
				goto IL_0354;
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
		}
		flag4 = chkPriValTitle.IsChecked;
		if (((!flag4) ?? flag4) == true)
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
			if (flag5.HasValue)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						chkPriVal.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_0354;
		IL_0103:
		bool? obj4;
		flag5 = (bool?)obj4;
		if ((flag5 ?? true) && chkPriValTitle.IsChecked == true)
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
			if (flag5.HasValue)
			{
				chkPriVal.IsChecked = true;
				return;
			}
		}
		flag = chkPriValScale.IsChecked;
		flag = (isChecked2 = (!flag) ?? flag);
		if (flag.HasValue)
		{
			if (isChecked2 != true)
			{
				obj2 = false;
				goto IL_0228;
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
		}
		flag = chkPriValGridlines.IsChecked;
		bool? flag6;
		if (!flag.HasValue)
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
			flag6 = flag;
		}
		else
		{
			flag6 = flag != true;
		}
		isChecked = flag6;
		flag = flag6;
		if (flag.HasValue)
		{
			obj2 = (isChecked == true) & isChecked2;
		}
		else
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
			obj2 = null;
		}
		goto IL_0228;
		IL_0089:
		flag4 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
		{
			if (flag4 != true)
			{
				obj4 = false;
				goto IL_0103;
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
		isChecked2 = (flag2 = chkPriValTicks.IsChecked);
		if (!isChecked2.HasValue)
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
			obj4 = null;
		}
		else if (flag2 != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = flag4;
		}
		goto IL_0103;
	}

	private void N()
	{
		chkPriCatScale.Checked += PriCatAxisChildCheckChanged;
		chkPriCatScale.Unchecked += PriCatAxisChildCheckChanged;
		chkPriCatGridlines.Checked += PriCatAxisChildCheckChanged;
		chkPriCatGridlines.Unchecked += PriCatAxisChildCheckChanged;
		chkPriCatTicks.Checked += PriCatAxisChildCheckChanged;
		chkPriCatTicks.Unchecked += PriCatAxisChildCheckChanged;
		chkPriCatTitle.Checked += PriCatAxisChildCheckChanged;
		chkPriCatTitle.Unchecked += PriCatAxisChildCheckChanged;
	}

	private void O()
	{
		chkPriCatScale.Checked -= PriCatAxisChildCheckChanged;
		chkPriCatScale.Unchecked -= PriCatAxisChildCheckChanged;
		chkPriCatGridlines.Checked -= PriCatAxisChildCheckChanged;
		chkPriCatGridlines.Unchecked -= PriCatAxisChildCheckChanged;
		chkPriCatTicks.Checked -= PriCatAxisChildCheckChanged;
		chkPriCatTicks.Unchecked -= PriCatAxisChildCheckChanged;
		chkPriCatTitle.Checked -= PriCatAxisChildCheckChanged;
		chkPriCatTitle.Unchecked -= PriCatAxisChildCheckChanged;
	}

	private void PriCatAxisCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkPriCat.IsChecked.Value;
		O();
		chkPriCatScale.IsChecked = value;
		chkPriCatGridlines.IsChecked = value;
		chkPriCatTicks.IsChecked = value;
		chkPriCatTitle.IsChecked = value;
		N();
	}

	private void PriCatAxisChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkPriCatScale.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			if (isChecked != true)
			{
				obj = false;
				goto IL_0089;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		bool? isChecked2;
		flag = (isChecked2 = chkPriCatGridlines.IsChecked);
		if (!flag.HasValue)
		{
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0089;
		IL_0105:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (flag2.HasValue)
		{
			if (flag2 != true)
			{
				goto IL_016e;
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
		}
		if (chkPriCatTitle.IsChecked == true)
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
			if (flag2.HasValue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						chkPriCat.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_016e;
		IL_016e:
		flag = chkPriCatScale.IsChecked;
		flag = (isChecked2 = (!flag) ?? flag);
		bool? obj3;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj3 = false;
				goto IL_0240;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		flag = chkPriCatGridlines.IsChecked;
		bool? flag3;
		if (!flag.HasValue)
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
			flag3 = flag;
		}
		else
		{
			flag3 = flag != true;
		}
		isChecked = flag3;
		flag = flag3;
		if (!flag.HasValue)
		{
			obj3 = null;
		}
		else if (isChecked != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = isChecked2;
		}
		goto IL_0240;
		IL_02e4:
		bool? obj4;
		flag2 = (bool?)obj4;
		if (flag2.HasValue)
		{
			if (flag2 != true)
			{
				goto IL_035f;
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
		}
		bool? isChecked3 = chkPriCatTitle.IsChecked;
		if (((!isChecked3) ?? isChecked3) == true)
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
			if (flag2.HasValue)
			{
				chkPriCat.IsChecked = false;
				return;
			}
		}
		goto IL_035f;
		IL_0240:
		bool? flag4 = obj3;
		isChecked = obj3;
		if (isChecked.HasValue)
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
			if (flag4 != true)
			{
				obj4 = false;
				goto IL_02e4;
			}
		}
		isChecked = chkPriCatTicks.IsChecked;
		bool? flag5;
		if (!isChecked.HasValue)
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
			flag5 = isChecked;
		}
		else
		{
			flag5 = isChecked != true;
		}
		isChecked3 = flag5;
		isChecked = flag5;
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
			obj4 = null;
		}
		else if (isChecked3 != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = flag4;
		}
		goto IL_02e4;
		IL_0089:
		isChecked3 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
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
			if (isChecked3 != true)
			{
				obj2 = false;
				goto IL_0105;
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
		isChecked2 = (flag4 = chkPriCatTicks.IsChecked);
		if (isChecked2.HasValue)
		{
			obj2 = (flag4 == true) & isChecked3;
		}
		else
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
		goto IL_0105;
		IL_035f:
		chkPriCat.IsChecked = null;
	}

	private void P()
	{
		chkSecValScale.Checked += SecValAxisChildCheckChanged;
		chkSecValScale.Unchecked += SecValAxisChildCheckChanged;
		chkSecValGridlines.Checked += SecValAxisChildCheckChanged;
		chkSecValGridlines.Unchecked += SecValAxisChildCheckChanged;
		chkSecValTicks.Checked += SecValAxisChildCheckChanged;
		chkSecValTicks.Unchecked += SecValAxisChildCheckChanged;
		chkSecValTitle.Checked += SecValAxisChildCheckChanged;
		chkSecValTitle.Unchecked += SecValAxisChildCheckChanged;
	}

	private void Q()
	{
		chkSecValScale.Checked -= SecValAxisChildCheckChanged;
		chkSecValScale.Unchecked -= SecValAxisChildCheckChanged;
		chkSecValGridlines.Checked -= SecValAxisChildCheckChanged;
		chkSecValGridlines.Unchecked -= SecValAxisChildCheckChanged;
		chkSecValTicks.Checked -= SecValAxisChildCheckChanged;
		chkSecValTicks.Unchecked -= SecValAxisChildCheckChanged;
		chkSecValTitle.Checked -= SecValAxisChildCheckChanged;
		chkSecValTitle.Unchecked -= SecValAxisChildCheckChanged;
	}

	private void SecValAxisCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkSecVal.IsChecked.Value;
		Q();
		chkSecValScale.IsChecked = value;
		chkSecValGridlines.IsChecked = value;
		chkSecValTicks.IsChecked = value;
		chkSecValTitle.IsChecked = value;
		P();
	}

	private void SecValAxisChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkSecValScale.IsChecked);
		bool? obj;
		if (flag.HasValue)
		{
			if (isChecked != true)
			{
				obj = false;
				goto IL_0095;
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
		}
		bool? isChecked2;
		flag = (isChecked2 = chkSecValGridlines.IsChecked);
		if (!flag.HasValue)
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0095;
		IL_0111:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (flag2.HasValue)
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
			if (flag2 != true)
			{
				goto IL_0174;
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
		}
		if (chkSecValTitle.IsChecked == true && flag2.HasValue)
		{
			chkSecVal.IsChecked = true;
			return;
		}
		goto IL_0174;
		IL_0174:
		flag = chkSecValScale.IsChecked;
		bool? flag3;
		if (!flag.HasValue)
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
			flag3 = flag;
		}
		else
		{
			flag3 = flag != true;
		}
		isChecked2 = flag3;
		flag = flag3;
		bool? obj3;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj3 = false;
				goto IL_024c;
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
		flag = chkSecValGridlines.IsChecked;
		flag = (isChecked = (!flag) ?? flag);
		if (!flag.HasValue)
		{
			obj3 = null;
		}
		else if (isChecked != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = isChecked2;
		}
		goto IL_024c;
		IL_024c:
		bool? flag4 = obj3;
		isChecked = obj3;
		bool? obj4;
		if (isChecked.HasValue)
		{
			if (flag4 != true)
			{
				obj4 = false;
				goto IL_02e4;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		isChecked = chkSecValTicks.IsChecked;
		bool? flag5;
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
			flag5 = isChecked;
		}
		else
		{
			flag5 = isChecked != true;
		}
		bool? flag6 = flag5;
		isChecked = flag5;
		if (!isChecked.HasValue)
		{
			obj4 = null;
		}
		else if (flag6 != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = flag4;
		}
		goto IL_02e4;
		IL_0095:
		flag6 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
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
			if (flag6 != true)
			{
				obj2 = false;
				goto IL_0111;
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
		}
		isChecked2 = (flag4 = chkSecValTicks.IsChecked);
		if (isChecked2.HasValue)
		{
			obj2 = (flag4 == true) & flag6;
		}
		else
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
			obj2 = null;
		}
		goto IL_0111;
		IL_037a:
		chkSecVal.IsChecked = null;
		return;
		IL_02e4:
		flag2 = obj4;
		if (flag2.HasValue)
		{
			if (flag2 != true)
			{
				goto IL_037a;
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
		flag6 = chkSecValTitle.IsChecked;
		bool? flag7;
		if (!flag6.HasValue)
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
			flag7 = flag6;
		}
		else
		{
			flag7 = flag6 != true;
		}
		flag6 = flag7;
		if (flag6 == true)
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
			if (flag2.HasValue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						chkSecVal.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_037a;
	}

	private void R()
	{
		chkSecCatScale.Checked += SecCatAxisChildCheckChanged;
		chkSecCatScale.Unchecked += SecCatAxisChildCheckChanged;
		chkSecCatGridlines.Checked += SecCatAxisChildCheckChanged;
		chkSecCatGridlines.Unchecked += SecCatAxisChildCheckChanged;
		chkSecCatTicks.Checked += SecCatAxisChildCheckChanged;
		chkSecCatTicks.Unchecked += SecCatAxisChildCheckChanged;
		chkSecCatTitle.Checked += SecCatAxisChildCheckChanged;
		chkSecCatTitle.Unchecked += SecCatAxisChildCheckChanged;
	}

	private void S()
	{
		chkSecCatScale.Checked -= SecCatAxisChildCheckChanged;
		chkSecCatScale.Unchecked -= SecCatAxisChildCheckChanged;
		chkSecCatGridlines.Checked -= SecCatAxisChildCheckChanged;
		chkSecCatGridlines.Unchecked -= SecCatAxisChildCheckChanged;
		chkSecCatTicks.Checked -= SecCatAxisChildCheckChanged;
		chkSecCatTicks.Unchecked -= SecCatAxisChildCheckChanged;
		chkSecCatTitle.Checked -= SecCatAxisChildCheckChanged;
		chkSecCatTitle.Unchecked -= SecCatAxisChildCheckChanged;
	}

	private void SecCatAxisCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkSecCat.IsChecked.Value;
		S();
		chkSecCatScale.IsChecked = value;
		chkSecCatGridlines.IsChecked = value;
		chkSecCatTicks.IsChecked = value;
		chkSecCatTitle.IsChecked = value;
		R();
	}

	private void SecCatAxisChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkSecCatScale.IsChecked);
		bool? obj;
		if (flag.HasValue)
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_0087;
			}
		}
		bool? isChecked2;
		flag = (isChecked2 = chkSecCatGridlines.IsChecked);
		if (flag.HasValue)
		{
			obj = (isChecked2 == true) & isChecked;
		}
		else
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
			obj = null;
		}
		goto IL_0087;
		IL_010d:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (flag2.HasValue)
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
			if (flag2 != true)
			{
				goto IL_0170;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		if (chkSecCatTitle.IsChecked == true && flag2.HasValue)
		{
			chkSecCat.IsChecked = true;
			return;
		}
		goto IL_0170;
		IL_02e4:
		bool? obj3;
		flag2 = (bool?)obj3;
		if (flag2.HasValue)
		{
			if (flag2 != true)
			{
				goto IL_0367;
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
		bool? isChecked3 = chkSecCatTitle.IsChecked;
		bool? flag3;
		if (!isChecked3.HasValue)
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
			flag3 = isChecked3;
		}
		else
		{
			flag3 = isChecked3 != true;
		}
		isChecked3 = flag3;
		if (isChecked3 == true && flag2.HasValue)
		{
			chkSecCat.IsChecked = false;
			return;
		}
		goto IL_0367;
		IL_0170:
		flag = chkSecCatScale.IsChecked;
		flag = (isChecked2 = (!flag) ?? flag);
		bool? obj4;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj4 = false;
				goto IL_024a;
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
		}
		flag = chkSecCatGridlines.IsChecked;
		bool? flag4;
		if (!flag.HasValue)
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
			flag4 = flag;
		}
		else
		{
			flag4 = flag != true;
		}
		isChecked = flag4;
		flag = flag4;
		if (!flag.HasValue)
		{
			obj4 = null;
		}
		else if (isChecked != true)
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
			obj4 = false;
		}
		else
		{
			obj4 = isChecked2;
		}
		goto IL_024a;
		IL_024a:
		bool? flag5 = obj4;
		isChecked = obj4;
		if (isChecked.HasValue)
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
			if (flag5 != true)
			{
				obj3 = false;
				goto IL_02e4;
			}
		}
		isChecked = chkSecCatTicks.IsChecked;
		bool? flag6;
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
			flag6 = isChecked;
		}
		else
		{
			flag6 = isChecked != true;
		}
		isChecked3 = flag6;
		isChecked = flag6;
		if (isChecked.HasValue)
		{
			obj3 = (isChecked3 == true) & flag5;
		}
		else
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
			obj3 = null;
		}
		goto IL_02e4;
		IL_0087:
		isChecked3 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
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
			if (isChecked3 != true)
			{
				obj2 = false;
				goto IL_010d;
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
		}
		isChecked2 = (flag5 = chkSecCatTicks.IsChecked);
		if (!isChecked2.HasValue)
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
			obj2 = null;
		}
		else if (flag5 != true)
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
			obj2 = false;
		}
		else
		{
			obj2 = isChecked3;
		}
		goto IL_010d;
		IL_0367:
		chkSecCat.IsChecked = null;
	}

	private void T()
	{
		chkLegendFormat.Checked += LegendChildCheckChanged;
		chkLegendFormat.Unchecked += LegendChildCheckChanged;
		chkLegendPosn.Checked += LegendChildCheckChanged;
		chkLegendPosn.Unchecked += LegendChildCheckChanged;
	}

	private void U()
	{
		chkLegendFormat.Checked -= LegendChildCheckChanged;
		chkLegendFormat.Unchecked -= LegendChildCheckChanged;
		chkLegendPosn.Checked -= LegendChildCheckChanged;
		chkLegendPosn.Unchecked -= LegendChildCheckChanged;
	}

	private void LegendCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkLegend.IsChecked.Value;
		U();
		chkLegendFormat.IsChecked = value;
		chkLegendPosn.IsChecked = value;
		T();
	}

	private void LegendChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkLegendFormat.IsChecked;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_007b;
			}
		}
		if (chkLegendPosn.IsChecked == true)
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
			if (isChecked.HasValue)
			{
				chkLegend.IsChecked = true;
				return;
			}
		}
		goto IL_007b;
		IL_007b:
		bool? isChecked2 = chkLegendFormat.IsChecked;
		bool? flag;
		if (!isChecked2.HasValue)
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
			flag = isChecked2;
		}
		else
		{
			flag = isChecked2 != true;
		}
		isChecked = flag;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_0144;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		isChecked2 = chkLegendPosn.IsChecked;
		bool? flag2;
		if (!isChecked2.HasValue)
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
			flag2 = isChecked2;
		}
		else
		{
			flag2 = isChecked2 != true;
		}
		isChecked2 = flag2;
		if (isChecked2 == true)
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
			if (isChecked.HasValue)
			{
				chkLegend.IsChecked = false;
				return;
			}
		}
		goto IL_0144;
		IL_0144:
		chkLegend.IsChecked = null;
	}

	private void V()
	{
		chkTitleFormat.Checked += TitleChildCheckChanged;
		chkTitleFormat.Unchecked += TitleChildCheckChanged;
		chkTitlePosn.Checked += TitleChildCheckChanged;
		chkTitlePosn.Unchecked += TitleChildCheckChanged;
	}

	private void W()
	{
		chkTitleFormat.Checked -= TitleChildCheckChanged;
		chkTitleFormat.Unchecked -= TitleChildCheckChanged;
		chkTitlePosn.Checked -= TitleChildCheckChanged;
		chkTitlePosn.Unchecked -= TitleChildCheckChanged;
	}

	private void TitleCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkTitle.IsChecked.Value;
		W();
		chkTitleFormat.IsChecked = value;
		chkTitlePosn.IsChecked = value;
		V();
	}

	private void TitleChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkTitleFormat.IsChecked;
		if (isChecked.HasValue)
		{
			if (isChecked != true)
			{
				goto IL_0075;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		if (chkTitlePosn.IsChecked == true)
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
			if (isChecked.HasValue)
			{
				chkTitle.IsChecked = true;
				return;
			}
		}
		goto IL_0075;
		IL_0137:
		chkTitle.IsChecked = null;
		return;
		IL_0075:
		bool? isChecked2 = chkTitleFormat.IsChecked;
		bool? flag;
		if (!isChecked2.HasValue)
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
			flag = isChecked2;
		}
		else
		{
			flag = isChecked2 != true;
		}
		isChecked = flag;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_0137;
			}
		}
		isChecked2 = chkTitlePosn.IsChecked;
		bool? flag2;
		if (!isChecked2.HasValue)
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
			flag2 = isChecked2;
		}
		else
		{
			flag2 = isChecked2 != true;
		}
		isChecked2 = flag2;
		if (isChecked2 == true && isChecked.HasValue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkTitle.IsChecked = false;
					return;
				}
			}
		}
		goto IL_0137;
	}

	private void X()
	{
		chkDataTableFormat.Checked += DataTableChildCheckChanged;
		chkDataTableFormat.Unchecked += DataTableChildCheckChanged;
	}

	private void Y()
	{
		chkDataTableFormat.Checked -= DataTableChildCheckChanged;
		chkDataTableFormat.Unchecked -= DataTableChildCheckChanged;
	}

	private void DataTableCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkDataTable.IsChecked.Value;
		Y();
		chkDataTableFormat.IsChecked = value;
		X();
	}

	private void DataTableChildCheckChanged(object sender, RoutedEventArgs e)
	{
		if (chkDataTableFormat.IsChecked == true)
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
					chkDataTable.IsChecked = true;
					return;
				}
			}
		}
		bool? isChecked = chkDataTableFormat.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		if (isChecked == true)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkDataTable.IsChecked = false;
					return;
				}
			}
		}
		chkDataTable.IsChecked = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(VH.A(172685), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
			btnCopy = (Button)target;
			return;
		}
		if (connectionId == 2)
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
					btnApply = (Button)target;
					return;
				}
			}
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
					((StackPanel)target).MouseDown += ExpandCollapseSource;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			expSource = (Polygon)target;
			return;
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
					gridSource = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					imgSource = (System.Windows.Controls.Image)target;
					return;
				}
			}
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
					((StackPanel)target).MouseDown += ExpandCollapseChart;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					expChart = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkChart = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					gridChart = (Grid)target;
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
					chkChartFormat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtChartFormat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkChartSize = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtChartSize = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkChartTop = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtChartTop = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkChartLeft = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtChartLeft = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapsePlotArea;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					expPlotArea = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkPlotArea = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			gridPlotArea = (Grid)target;
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
					chkPlotFormat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			txtPlotFormat = (TextBlock)target;
			return;
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkPlotSize = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtPlotSize = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPlotLocation = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtPlotLocation = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseSeries;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					expSeries = (Polygon)target;
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
					chkSeries = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					gridSeries = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkSeriesFormat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtSeriesFormat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			chkSeriesDataLabels = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtSeriesDataLabels = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkSeriesGap = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtSeriesGap = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			chkSeriesAngle = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtSeriesAngle = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			chkSeriesExplosion = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtSeriesExplosion = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkSeriesErrorBars = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtSeriesErrorBars = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkSeriesUpDownBars = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 46)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtSeriesUpDownBars = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 47)
		{
			((StackPanel)target).MouseDown += ExpandCollapseAxes;
			return;
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					expAxes = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 49)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkAxes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 50)
		{
			gridAxes = (Grid)target;
			return;
		}
		if (connectionId == 51)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPriVal = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 52)
		{
			txtPriVal = (TextBlock)target;
			return;
		}
		if (connectionId == 53)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPriValScale = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 54)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtPriValScale = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 55)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkPriValGridlines = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 56)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtPriValGridlines = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkPriValTicks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 58)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtPriValTicks = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 59)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkPriValTitle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 60)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtPriValTitle = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 61)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPriCat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 62)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtPriCat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 63)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkPriCatScale = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 64)
		{
			txtPriCatScale = (TextBlock)target;
			return;
		}
		if (connectionId == 65)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkPriCatGridlines = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 66)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtPriCatGridlines = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 67)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPriCatTicks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 68)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtPriCatTicks = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 69)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPriCatTitle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 70)
		{
			txtPriCatTitle = (TextBlock)target;
			return;
		}
		if (connectionId == 71)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkSecVal = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 72)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtSecVal = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 73)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSecValScale = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 74)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtSecValScale = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 75)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkSecValGridlines = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 76)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtSecValGridlines = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 77)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSecValTicks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 78)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtSecValTicks = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 79)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSecValTitle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 80)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtSecValTitle = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 81)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkSecCat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 82)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtSecCat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 83)
		{
			chkSecCatScale = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 84)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtSecCatScale = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 85)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSecCatGridlines = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 86)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtSecCatGridlines = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 87)
		{
			chkSecCatTicks = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 88)
		{
			txtSecCatTicks = (TextBlock)target;
			return;
		}
		if (connectionId == 89)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkSecCatTitle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 90)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtSecCatTitle = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 91)
		{
			((StackPanel)target).MouseDown += ExpandCollapseLegend;
			return;
		}
		if (connectionId == 92)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					expLegend = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 93)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLegend = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 94)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					gridLegend = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 95)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkLegendFormat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 96)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtLegendFormat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 97)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkLegendPosn = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 98)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtLegendPosn = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 99)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseTitle;
					return;
				}
			}
		}
		if (connectionId == 100)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					expTitle = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 101)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkTitle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 102)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					gridTitle = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 103)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkTitleFormat = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 104)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtTitleFormat = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 105)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkTitlePosn = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 106)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtTitlePosn = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 107)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseDataTable;
					return;
				}
			}
		}
		if (connectionId == 108)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					expDataTable = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 109)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkDataTable = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 110:
			gridDataTable = (Grid)target;
			break;
		case 111:
			chkDataTableFormat = (System.Windows.Controls.CheckBox)target;
			break;
		case 112:
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				txtDataTableFormat = (TextBlock)target;
				return;
			}
		default:
			this.m_B = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
