using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Windows.Shapes;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.UI;

[DesignerGenerated]
public sealed class wpfCallout : Window, INotifyPropertyChanged, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<IMsoDataLabel, double> A;

		public static Func<IMsoDataLabel, double> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal double A(IMsoDataLabel A)
		{
			return A.Top;
		}

		[SpecialName]
		internal double B(IMsoDataLabel A)
		{
			return A.Left;
		}
	}

	[CompilerGenerated]
	internal sealed class UC
	{
		public TextRange2 A;

		public VC A;

		public UC(UC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetTextRangeRectangle(this.A, this.A.A, this.A.B);
		}
	}

	[CompilerGenerated]
	internal sealed class VC
	{
		public float A;

		public float B;

		public BaseError A;

		public VC(VC A)
		{
			if (A == null)
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
				this.A = A.A;
				this.B = A.B;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.TextRangesTopLeft(((BaseError)this.A).TextRanges, this.A, this.B);
		}

		[SpecialName]
		internal Rect B()
		{
			return MarchingAnts.GetPlotAreaRectangle(this.A.PlotArea, this.A, this.B);
		}

		[SpecialName]
		internal Rect C()
		{
			return MarchingAnts.GetPlotAreaRectangle(this.A.Shape.Chart.PlotArea, this.A, this.B);
		}

		[SpecialName]
		internal Rect D()
		{
			return MarchingAnts.GetLegendEntryRectangle(((BaseError)this.A).LegendEntry, this.A, this.B);
		}

		[SpecialName]
		internal Rect E()
		{
			return MarchingAnts.GetLegendKeyRectangle(((BaseError)this.A).LegendKey, this.A, this.B);
		}

		[SpecialName]
		internal Rect F()
		{
			return MarchingAnts.GetLegendRectangle(this.A.Legend, this.A, this.B);
		}

		[SpecialName]
		internal Rect G()
		{
			return MarchingAnts.GetChartTitleRectangle(this.A.ChartTitle, this.A, this.B);
		}

		[SpecialName]
		internal Rect H()
		{
			return MarchingAnts.GetAxisRectangle(this.A.Axis, this.A, this.B);
		}

		[SpecialName]
		internal Rect I()
		{
			return MarchingAnts.GetAxisTitleRectangle(this.A.AxisTitle, this.A, this.B);
		}

		[SpecialName]
		internal Rect J()
		{
			return MarchingAnts.GetPlotAreaOuterRectangle(this.A.Shape);
		}

		[SpecialName]
		internal Rect K()
		{
			return MarchingAnts.GetChartPointRectangle(this.A.Shape, ((BaseError)this.A).ChartPoint);
		}

		[SpecialName]
		internal Rect L()
		{
			return MarchingAnts.GetShapeRectangle(this.A.Shape);
		}
	}

	[CompilerGenerated]
	internal sealed class WC
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public WC(WC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetTextFrameRectangle(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class XC
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public XC(XC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetShapeRectangle(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class YC
	{
		public Microsoft.Office.Core.Shape A;

		public YC(YC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetShapeRectangle(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class ZC
	{
		public IMsoDataLabel A;

		public AD A;

		public ZC(ZC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetLabelRectangle(this.A, this.A.A.A, this.A.A.B);
		}
	}

	[CompilerGenerated]
	internal sealed class AD
	{
		public List<IMsoDataLabel> A;

		public VC A;

		public AD(AD A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return MarchingAnts.GetObjectRectangle((float)(this.A[0].Left + (double)this.A.A), (float)(this.A[0].Top + (double)this.A.B), 1f, 1f);
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private BaseError m_A;

	[CompilerGenerated]
	private double m_A;

	private bool m_A;

	[AccessedThroughProperty("gridMain")]
	[CompilerGenerated]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("pathIcon")]
	private Path m_A;

	[AccessedThroughProperty("tbTitle")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("btnPrevious")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnNext")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFix")]
	private System.Windows.Controls.Button C;

	[AccessedThroughProperty("btnPopup")]
	[CompilerGenerated]
	private ToggleButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private System.Windows.Controls.Button D;

	private bool m_B;

	public BaseError ActiveError
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(57355));
		}
	}

	public double XOffset
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

	internal virtual Grid gridMain
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

	internal virtual Path pathIcon
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

	internal virtual TextBlock tbTitle
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

	internal virtual System.Windows.Controls.Button btnPrevious
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
			RoutedEventHandler value2 = btnPrevious_Click;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
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
				switch (4)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnNext
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
			RoutedEventHandler value2 = btnNext_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
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
				switch (2)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnFix
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
			RoutedEventHandler value2 = FixButtonClicked;
			System.Windows.Controls.Button button = C;
			if (button != null)
			{
				button.Click -= value2;
			}
			C = value;
			button = C;
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

	internal virtual ToggleButton btnPopup
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
			RoutedEventHandler value2 = ShowFixOptions;
			ToggleButton toggleButton = this.m_A;
			if (toggleButton != null)
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
				toggleButton.Checked -= value2;
			}
			this.m_A = value;
			toggleButton = this.m_A;
			if (toggleButton == null)
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
				toggleButton.Checked += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = D;
			if (button != null)
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
				button.Click -= value2;
			}
			D = value;
			button = D;
			if (button == null)
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
				button.Click += value2;
				return;
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
	}

	public wpfCallout()
	{
		base.Loaded += wpfCallout_Loaded;
		base.Closing += wpfCallout_FormClosing;
		InitializeComponent();
		Callout.Dialog = this;
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfCallout_Loaded(object sender, RoutedEventArgs e)
	{
		if (Callout.MarchingAnts == null)
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
			A();
			return;
		}
	}

	internal void A(BaseError A, ref int B)
	{
		//IL_0a23: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a2d: Expected O, but got Unknown
		VC a = default(VC);
		VC CS_0024_003C_003E8__locals118 = new VC(a);
		CS_0024_003C_003E8__locals118.A = A;
		double? D = null;
		double? C = null;
		CS_0024_003C_003E8__locals118.A = 0f;
		CS_0024_003C_003E8__locals118.B = 0f;
		Callout.A();
		if (CS_0024_003C_003E8__locals118.A.Slide == null)
		{
			base.Top = -10000.0;
			return;
		}
		base.Deactivated -= CalloutDeactivated;
		NG.A.Application.ActiveWindow.View.GotoSlide(CS_0024_003C_003E8__locals118.A.Slide.SlideIndex);
		System.Windows.Forms.Application.DoEvents();
		if (CS_0024_003C_003E8__locals118.A.Shape != null && CS_0024_003C_003E8__locals118.A.Shape.HasChart == MsoTriState.msoTrue)
		{
			ErrorType type = CS_0024_003C_003E8__locals118.A.Type;
			if (type == ErrorType.ColorPaletteChartSeries)
			{
				Callout.DoNotClose = true;
				Chart chart = CS_0024_003C_003E8__locals118.A.Shape.Chart;
				try
				{
					CS_0024_003C_003E8__locals118.A.Shape.Select();
					chart.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
					ErrorType type2 = CS_0024_003C_003E8__locals118.A.Type;
					if (type2 == ErrorType.ColorPaletteChartSeries)
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
							((BaseError)CS_0024_003C_003E8__locals118.A).Series.Select();
							break;
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				chart = null;
				Callout.DoNotClose = false;
				Pane.F();
			}
		}
		ActiveError = CS_0024_003C_003E8__locals118.A;
		XOffset = 0.0;
		List<Rect> A2 = new List<Rect>();
		if (((BaseError)CS_0024_003C_003E8__locals118.A).TextRanges != null)
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
			if (MarchingAnts.UseRelativePosition(CS_0024_003C_003E8__locals118.A.Shape))
			{
				CS_0024_003C_003E8__locals118.A = CS_0024_003C_003E8__locals118.A.Shape.Left;
				CS_0024_003C_003E8__locals118.B = CS_0024_003C_003E8__locals118.A.Shape.Top;
			}
			IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
			try
			{
				enumerator = ((BaseError)CS_0024_003C_003E8__locals118.A).TextRanges.GetEnumerator();
				UC uC = default(UC);
				while (enumerator.MoveNext())
				{
					uC = new UC(uC);
					uC.A = CS_0024_003C_003E8__locals118;
					uC.A = enumerator.Current;
					try
					{
						this.A(ref A2, uC.A, ref C, ref D, ref B);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						clsReporting.LogException(ex4);
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0282;
					}
					continue;
					end_IL_0282:
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			Func<Rect> a2 = [SpecialName] () => MarchingAnts.TextRangesTopLeft(((BaseError)CS_0024_003C_003E8__locals118.A).TextRanges, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B);
			double B2 = C.Value;
			double C2 = D.Value;
			this.A(a2, ref B2, ref C2);
			D = C2;
			C = B2;
		}
		else if (CS_0024_003C_003E8__locals118.A.Shapes != null)
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
			if (CS_0024_003C_003E8__locals118.A.Type != ErrorType.TableCellMargins)
			{
				using List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator2 = CS_0024_003C_003E8__locals118.A.Shapes.GetEnumerator();
				XC xC = default(XC);
				while (enumerator2.MoveNext())
				{
					xC = new XC(xC);
					xC.A = enumerator2.Current;
					this.A(ref A2, xC.A, ref C, ref D, ref B);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_03dc;
					}
					continue;
					end_IL_03dc:
					break;
				}
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
				using List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator3 = CS_0024_003C_003E8__locals118.A.Shapes.GetEnumerator();
				WC wC = default(WC);
				while (enumerator3.MoveNext())
				{
					wC = new WC(wC);
					wC.A = enumerator3.Current;
					this.A(ref A2, wC.A, ref C, ref D, ref B);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_036c;
					}
					continue;
					end_IL_036c:
					break;
				}
			}
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).OfficeShapes != null)
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
			using IEnumerator<Microsoft.Office.Core.Shape> enumerator4 = ((BaseError)CS_0024_003C_003E8__locals118.A).OfficeShapes.GetEnumerator();
			YC yC = default(YC);
			while (enumerator4.MoveNext())
			{
				yC = new YC(yC);
				yC.A = enumerator4.Current;
				this.A(ref A2, yC.A, ref C, ref D, ref B);
			}
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).DataLabels != null)
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
			AD a3 = default(AD);
			AD CS_0024_003C_003E8__locals90 = new AD(a3);
			CS_0024_003C_003E8__locals90.A = CS_0024_003C_003E8__locals118;
			CS_0024_003C_003E8__locals90.A.A = (float)this.A(CS_0024_003C_003E8__locals90.A.A);
			CS_0024_003C_003E8__locals90.A.B = (float)this.B(CS_0024_003C_003E8__locals90.A.A);
			using (IEnumerator<IMsoDataLabel> enumerator5 = ((BaseError)CS_0024_003C_003E8__locals90.A.A).DataLabels.GetEnumerator())
			{
				ZC zC = default(ZC);
				while (enumerator5.MoveNext())
				{
					zC = new ZC(zC);
					zC.A = CS_0024_003C_003E8__locals90;
					zC.A = enumerator5.Current;
					this.A(ref A2, zC.A, ref C, ref D, ref B);
				}
			}
			AD aD = CS_0024_003C_003E8__locals90;
			IList<IMsoDataLabel> dataLabels = ((BaseError)CS_0024_003C_003E8__locals90.A.A).DataLabels;
			Func<IMsoDataLabel, double> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (IMsoDataLabel msoDataLabel) => msoDataLabel.Top);
			}
			else
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
				keySelector = _Closure_0024__.A;
			}
			IOrderedEnumerable<IMsoDataLabel> source = dataLabels.OrderBy(keySelector);
			Func<IMsoDataLabel, double> keySelector2;
			if (_Closure_0024__.B == null)
			{
				keySelector2 = (_Closure_0024__.B = [SpecialName] (IMsoDataLabel msoDataLabel) => msoDataLabel.Left);
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
				keySelector2 = _Closure_0024__.B;
			}
			aD.A = source.ThenBy(keySelector2).ToList();
			Func<Rect> a4 = [SpecialName] () => MarchingAnts.GetObjectRectangle((float)(CS_0024_003C_003E8__locals90.A[0].Left + (double)CS_0024_003C_003E8__locals90.A.A), (float)(CS_0024_003C_003E8__locals90.A[0].Top + (double)CS_0024_003C_003E8__locals90.A.B), 1f, 1f);
			double C2 = C.Value;
			double B2 = D.Value;
			this.A(a4, ref C2, ref B2);
			D = B2;
			C = C2;
			CS_0024_003C_003E8__locals90.A = null;
		}
		else if (CS_0024_003C_003E8__locals118.A.PlotArea != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetPlotAreaRectangle(CS_0024_003C_003E8__locals118.A.PlotArea, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).Series != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetPlotAreaRectangle(CS_0024_003C_003E8__locals118.A.Shape.Chart.PlotArea, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).LegendEntry != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetLegendEntryRectangle(((BaseError)CS_0024_003C_003E8__locals118.A).LegendEntry, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).LegendKey != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetLegendKeyRectangle(((BaseError)CS_0024_003C_003E8__locals118.A).LegendKey, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (CS_0024_003C_003E8__locals118.A.Legend != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetLegendRectangle(CS_0024_003C_003E8__locals118.A.Legend, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (CS_0024_003C_003E8__locals118.A.ChartTitle != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetChartTitleRectangle(CS_0024_003C_003E8__locals118.A.ChartTitle, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (CS_0024_003C_003E8__locals118.A.Axis != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetAxisRectangle(CS_0024_003C_003E8__locals118.A.Axis, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (CS_0024_003C_003E8__locals118.A.AxisTitle != null)
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
			CS_0024_003C_003E8__locals118.A = (float)this.A(CS_0024_003C_003E8__locals118.A);
			CS_0024_003C_003E8__locals118.B = (float)this.B(CS_0024_003C_003E8__locals118.A);
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetAxisTitleRectangle(CS_0024_003C_003E8__locals118.A.AxisTitle, CS_0024_003C_003E8__locals118.A, CS_0024_003C_003E8__locals118.B), ref C, ref D, ref B);
		}
		else if (CS_0024_003C_003E8__locals118.A.ChartGroup != null)
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
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetPlotAreaOuterRectangle(CS_0024_003C_003E8__locals118.A.Shape), ref C, ref D, ref B);
		}
		else if (((BaseError)CS_0024_003C_003E8__locals118.A).ChartPoint != null)
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
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetChartPointRectangle(CS_0024_003C_003E8__locals118.A.Shape, ((BaseError)CS_0024_003C_003E8__locals118.A).ChartPoint), ref C, ref D, ref B);
		}
		else
		{
			if (CS_0024_003C_003E8__locals118.A.Shape == null)
			{
				base.Top = -10000.0;
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
			this.A(ref A2, [SpecialName] () => MarchingAnts.GetShapeRectangle(CS_0024_003C_003E8__locals118.A.Shape), ref C, ref D, ref B);
		}
		Callout.DashBoxes = A2;
		UpdateLayout();
		Callout.A(this, C ?? (-10000.0), D ?? (-10000.0));
		if (A2.Count > 0)
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
			Callout.MarchingAnts = new wpfMarchingAnts(A2);
			if (base.IsLoaded)
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
				this.A();
			}
		}
		base.Deactivated += CalloutDeactivated;
		A2 = null;
	}

	private void A(ref List<Rect> A, Func<Rect> B, ref double? C, ref double? D, ref int E)
	{
		checked
		{
			try
			{
				Rect item = B();
				A.Add(item);
				if (C.HasValue)
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
					if (!(item.X < C.Value))
					{
						goto IL_005f;
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
				C = item.X;
				goto IL_005f;
				IL_005f:
				if (D.HasValue)
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
					if (!(item.Y < D.Value))
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
						break;
					}
				}
				D = item.Y;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				E++;
				ProjectData.ClearProjectError();
			}
		}
	}

	private void A(Func<Rect> A, ref double B, ref double C)
	{
		try
		{
			Rect rect = A();
			XOffset = rect.X - B;
			B = rect.X;
			C = rect.Y;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void A()
	{
		wpfMarchingAnts marchingAnts = Callout.MarchingAnts;
		try
		{
			((Window)(object)marchingAnts).Owner = Window.GetWindow(this);
		}
		catch (InvalidOperationException ex)
		{
			ProjectData.SetProjectError(ex);
			InvalidOperationException ex2 = ex;
			ProjectData.ClearProjectError();
			return;
		}
		((Window)(object)marchingAnts).Top = base.Top + base.ActualHeight;
		((Window)(object)marchingAnts).Left = base.Left + (double)Callout.POINTER_X_OFFSET - XOffset;
		((Window)(object)marchingAnts).ShowActivated = false;
		((Window)(object)marchingAnts).Show();
		marchingAnts = null;
	}

	private double A(BaseError A)
	{
		return MarchingAnts.ChartLeftOffset(A.Shape);
	}

	private double B(BaseError A)
	{
		return MarchingAnts.ChartTopOffset(A.Shape);
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void B()
	{
		if (!this.m_A)
		{
			try
			{
				this.m_A = true;
				Close();
			}
			finally
			{
				this.m_A = false;
			}
		}
	}

	private void wpfCallout_FormClosing(object sender, CancelEventArgs e)
	{
		bool a = this.m_A;
		try
		{
			this.m_A = true;
			Callout.MarchingAnts = null;
			Callout.Dialog = null;
		}
		finally
		{
			this.m_A = a;
		}
	}

	private void CalloutDeactivated(object sender, EventArgs e)
	{
		if (Callout.DoNotClose)
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
			if (!this.m_A)
			{
				try
				{
					B();
					return;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
					return;
				}
			}
			return;
		}
	}

	private void FixButtonClicked(object sender, RoutedEventArgs e)
	{
		Fixes.DefaultFixButtonClicked(ActiveError, suppressMsgs: true);
	}

	private void ShowFixOptions(object sender, RoutedEventArgs e)
	{
		Fixes.ShowOptions(ActiveError, (ToggleButton)sender, blnRefocusPane: false);
	}

	private void btnPrevious_Click(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ListBox lbxResults = Pane.TaskPane.WarningsView.lbxResults;
		checked
		{
			if (lbxResults.SelectedIndex > 0)
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
				lbxResults.SelectedIndex--;
				lbxResults.ScrollIntoView(RuntimeHelpers.GetObjectValue(lbxResults.SelectedItem));
			}
			lbxResults = null;
		}
	}

	private void btnNext_Click(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ListBox lbxResults = Pane.TaskPane.WarningsView.lbxResults;
		checked
		{
			if (lbxResults.SelectedIndex < lbxResults.Items.Count - 1)
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
				lbxResults.SelectedIndex++;
				lbxResults.ScrollIntoView(RuntimeHelpers.GetObjectValue(lbxResults.SelectedItem));
			}
			lbxResults = null;
		}
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(57378), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					gridMain = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					pathIcon = (Path)target;
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
					tbTitle = (TextBlock)target;
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
					btnPrevious = (System.Windows.Controls.Button)target;
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
					btnNext = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnFix = (System.Windows.Controls.Button)target;
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
					btnPopup = (ToggleButton)target;
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
					btnClose = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		this.m_B = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
