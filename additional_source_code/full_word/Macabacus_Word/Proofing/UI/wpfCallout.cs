using System;
using System.CodeDom.Compiler;
using System.Collections;
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
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.UI;

[DesignerGenerated]
public sealed class wpfCallout : System.Windows.Window, INotifyPropertyChanged, IComponentConnector
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
	private PropertyChangedEventHandler m_A;

	private BaseError m_A;

	private double m_A;

	private bool m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("gridMain")]
	private Grid m_A;

	[AccessedThroughProperty("pathIcon")]
	[CompilerGenerated]
	private Path m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbTitle")]
	private TextBlock m_A;

	[AccessedThroughProperty("btnPrevious")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnNext")]
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
			A(XC.A(38609));
		}
	}

	public double XOffset
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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
				toggleButton.Checked -= value2;
			}
			this.m_A = value;
			toggleButton = this.m_A;
			if (toggleButton != null)
			{
				toggleButton.Checked += value2;
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
			D = value;
			button = D;
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
				switch (1)
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

	public wpfCallout()
	{
		base.Loaded += wpfCallout_Loaded;
		base.Closing += wpfCallout_FormClosing;
		this.m_A = false;
		InitializeComponent();
		Callout.Dialog = this;
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
			switch (1)
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

	private void wpfCallout_Loaded(object sender, RoutedEventArgs e)
	{
		if (Callout.MarchingAnts == null)
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
			A();
			return;
		}
	}

	public void GoToError(BaseError err)
	{
		//IL_0996: Unknown result type (might be due to invalid IL or missing references)
		//IL_09a0: Expected O, but got Unknown
		double num = 10000.0;
		double num2 = 10000.0;
		float sngLeftOffset = 0f;
		float sngTopOffset = 0f;
		Callout.RemoveMarchingAnts();
		if (err.Range == null)
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
				base.Top = -10000.0;
				return;
			}
		}
		base.Deactivated -= CalloutDeactivated;
		Microsoft.Office.Interop.Word.Window activeWindow = PC.A.Application.ActiveWindow;
		Range range = err.Range;
		object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
		activeWindow.ScrollIntoView(range, ref Start);
		System.Windows.Forms.Application.DoEvents();
		if (err.Shape != null)
		{
			if (err.Shape.HasChart == MsoTriState.msoTrue)
			{
				goto IL_00d9;
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
		if (err.InlineShape != null && err.InlineShape.HasChart == MsoTriState.msoTrue)
		{
			goto IL_00d9;
		}
		goto IL_0187;
		IL_0187:
		ActiveError = err;
		XOffset = 0.0;
		List<Rect> list = new List<Rect>();
		this.m_A = true;
		Rect textRangeRectangle;
		if (((BaseError)err).TextRanges != null)
		{
			if (Callout.UseRelativePosition(err))
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
				sngLeftOffset = err.Shape.Left;
				sngTopOffset = err.Shape.Top;
			}
			IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
			try
			{
				enumerator = ((BaseError)err).TextRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					TextRange2 current = enumerator.Current;
					try
					{
						textRangeRectangle = Callout.GetTextRangeRectangle(current, sngLeftOffset, sngTopOffset);
						list.Add(textRangeRectangle);
						if (textRangeRectangle.X < num2)
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
							num2 = textRangeRectangle.X;
						}
						if (!(textRangeRectangle.Y < num))
						{
							continue;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							num = textRangeRectangle.Y;
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						clsReporting.LogException(ex2);
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
						goto end_IL_027b;
					}
					continue;
					end_IL_027b:
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
		}
		else if (err.Shapes != null)
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
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = err.Shapes.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					textRangeRectangle = Callout.GetObjectRectangle((Microsoft.Office.Interop.Word.Shape)enumerator2.Current);
					list.Add(textRangeRectangle);
					if (textRangeRectangle.X < num2)
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
						num2 = textRangeRectangle.X;
					}
					if (!(textRangeRectangle.Y < num))
					{
						continue;
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
					num = textRangeRectangle.Y;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_032d;
					}
					continue;
					end_IL_032d:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		else if (((BaseError)err).OfficeShapes != null)
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
			IEnumerator<Microsoft.Office.Core.Shape> enumerator3 = default(IEnumerator<Microsoft.Office.Core.Shape>);
			try
			{
				enumerator3 = ((BaseError)err).OfficeShapes.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					textRangeRectangle = Callout.GetShapeRectangle(enumerator3.Current);
					list.Add(textRangeRectangle);
					if (textRangeRectangle.X < num2)
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
						num2 = textRangeRectangle.X;
					}
					if (!(textRangeRectangle.Y < num))
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
					num = textRangeRectangle.Y;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_03e8;
					}
					continue;
					end_IL_03e8:
					break;
				}
			}
			finally
			{
				if (enumerator3 != null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						enumerator3.Dispose();
						break;
					}
				}
			}
		}
		else if (((BaseError)err).DataLabels != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			IEnumerator<IMsoDataLabel> enumerator4 = default(IEnumerator<IMsoDataLabel>);
			try
			{
				enumerator4 = ((BaseError)err).DataLabels.GetEnumerator();
				while (enumerator4.MoveNext())
				{
					textRangeRectangle = Callout.GetLabelRectangle(enumerator4.Current, sngLeftOffset, sngTopOffset);
					list.Add(textRangeRectangle);
					if (textRangeRectangle.X < num2)
					{
						num2 = textRangeRectangle.X;
					}
					if (!(textRangeRectangle.Y < num))
					{
						continue;
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
					num = textRangeRectangle.Y;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_04a7;
					}
					continue;
					end_IL_04a7:
					break;
				}
			}
			finally
			{
				if (enumerator4 != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						enumerator4.Dispose();
						break;
					}
				}
			}
			IList<IMsoDataLabel> dataLabels = ((BaseError)err).DataLabels;
			Func<IMsoDataLabel, double> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (IMsoDataLabel A) => A.Top);
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
				keySelector = _Closure_0024__.A;
			}
			IOrderedEnumerable<IMsoDataLabel> source = dataLabels.OrderBy(keySelector);
			Func<IMsoDataLabel, double> keySelector2;
			if (_Closure_0024__.B == null)
			{
				keySelector2 = (_Closure_0024__.B = [SpecialName] (IMsoDataLabel A) => A.Left);
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
				keySelector2 = _Closure_0024__.B;
			}
			List<IMsoDataLabel> list2 = source.ThenBy(keySelector2).ToList();
			textRangeRectangle = Callout.GetObjectRectangle((float)(list2[0].Left + (double)sngLeftOffset), (float)(list2[0].Top + (double)sngTopOffset), 1f, 1f);
			XOffset = textRangeRectangle.X - num2;
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
			list2 = null;
		}
		else if (err.PlotArea != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = Callout.GetPlotAreaRectangle(err.PlotArea, sngLeftOffset, sngTopOffset);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (((BaseError)err).Series != null)
		{
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = ((err.Shape == null) ? Callout.GetPlotAreaRectangle(err.InlineShape.Chart.PlotArea, sngLeftOffset, sngTopOffset) : Callout.GetPlotAreaRectangle(err.Shape.Chart.PlotArea, sngLeftOffset, sngTopOffset));
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.Legend != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = Callout.GetLegendRectangle(err.Legend, sngLeftOffset, sngTopOffset);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.ChartTitle != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = Callout.GetChartTitleRectangle(err.ChartTitle, sngLeftOffset, sngTopOffset);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.Axis != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = Callout.GetAxisRectangle(err.Axis, sngLeftOffset, sngTopOffset);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.AxisTitle != null)
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
			sngLeftOffset = (float)A(err);
			sngTopOffset = (float)B(err);
			textRangeRectangle = Callout.GetAxisTitleRectangle(err.AxisTitle, sngLeftOffset, sngTopOffset);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.InlineShape != null)
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
			textRangeRectangle = Callout.GetObjectRectangle(err.InlineShape);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.Shape != null)
		{
			textRangeRectangle = Callout.GetObjectRectangle(err.Shape);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.Table != null)
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
			textRangeRectangle = Callout.GetObjectRectangle(err.Table.Range);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else if (err.Range != null)
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
			textRangeRectangle = Callout.GetObjectRectangle(err.Range);
			list.Add(textRangeRectangle);
			num2 = textRangeRectangle.X;
			num = textRangeRectangle.Y;
		}
		else
		{
			Microsoft.Office.Interop.Word.Application application = PC.A.Application;
			Start = false;
			num2 = application.PointsToPixels(5f, ref Start);
			float pageHeight = application.ActiveDocument.PageSetup.PageHeight;
			Start = true;
			num = application.PointsToPixels(pageHeight, ref Start);
			_ = null;
		}
		this.m_A = false;
		Callout.DashBoxes = list;
		UpdateLayout();
		Callout.Reposition(this, num2, num);
		if (list.Count > 0)
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
			Callout.MarchingAnts = new wpfMarchingAnts(list);
			if (base.IsLoaded)
			{
				A();
			}
		}
		base.Deactivated += CalloutDeactivated;
		list = null;
		textRangeRectangle = default(Rect);
		return;
		IL_00d9:
		ErrorType type = err.Type;
		if (type == ErrorType.ColorPaletteChartSeries)
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
			Callout.DoNotClose = true;
			Chart chart = err.Shape.Chart;
			try
			{
				Microsoft.Office.Interop.Word.Shape shape = err.Shape;
				Start = RuntimeHelpers.GetObjectValue(Missing.Value);
				shape.Select(ref Start);
				chart.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
				ErrorType type2 = err.Type;
				if (type2 == ErrorType.ColorPaletteChartSeries)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						((BaseError)err).Series.Select();
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
			chart = null;
			Callout.DoNotClose = false;
			Pane.RefocusActiveListBoxItem();
		}
		goto IL_0187;
	}

	private void A()
	{
		wpfMarchingAnts marchingAnts = Callout.MarchingAnts;
		((System.Windows.Window)(object)marchingAnts).Owner = System.Windows.Window.GetWindow(this);
		((System.Windows.Window)(object)marchingAnts).Top = base.Top + base.ActualHeight;
		((System.Windows.Window)(object)marchingAnts).Left = base.Left + (double)Callout.POINTER_X_OFFSET - XOffset;
		((System.Windows.Window)(object)marchingAnts).ShowActivated = false;
		((System.Windows.Window)(object)marchingAnts).Show();
		_ = null;
	}

	private double A(BaseError A)
	{
		if (A.Shape != null)
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
					return (double)A.Shape.Left + A.Chart.ChartArea.Left;
				}
			}
		}
		_ = A.InlineShape;
		double result = default(double);
		return result;
	}

	private double B(BaseError A)
	{
		if (A.Shape != null)
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
					return (double)A.Shape.Top + A.Chart.ChartArea.Top;
				}
			}
		}
		double result = default(double);
		return result;
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void wpfCallout_FormClosing(object sender, CancelEventArgs e)
	{
		if (this.m_A)
		{
			e.Cancel = true;
			return;
		}
		Callout.MarchingAnts = null;
		Callout.Dialog = null;
	}

	private void CalloutDeactivated(object sender, EventArgs e)
	{
		if (Callout.DoNotClose)
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
			Close();
			return;
		}
	}

	private void FixButtonClicked(object sender, RoutedEventArgs e)
	{
		Fixes.DefaultFixButtonClicked(ActiveError);
	}

	private void ShowFixOptions(object sender, RoutedEventArgs e)
	{
		Fixes.ShowOptions(ActiveError, (ToggleButton)sender);
	}

	private void btnPrevious_Click(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ListBox lbxResults = Pane.TaskPane.lbxResults;
		checked
		{
			if (lbxResults.SelectedIndex > 0)
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
				lbxResults.SelectedIndex--;
				lbxResults.ScrollIntoView(RuntimeHelpers.GetObjectValue(lbxResults.SelectedItem));
			}
			lbxResults = null;
		}
	}

	private void btnNext_Click(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ListBox lbxResults = Pane.TaskPane.lbxResults;
		checked
		{
			if (lbxResults.SelectedIndex < lbxResults.Items.Count - 1)
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
				lbxResults.SelectedIndex++;
				lbxResults.ScrollIntoView(RuntimeHelpers.GetObjectValue(lbxResults.SelectedItem));
			}
			lbxResults = null;
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_B)
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
			this.m_B = true;
			Uri resourceLocator = new Uri(XC.A(38632), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			gridMain = (Grid)target;
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
				switch (5)
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
				switch (3)
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
				switch (5)
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
				switch (4)
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
				switch (4)
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
