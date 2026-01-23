using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts.GrowthArrow;

public sealed class Arrow : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly string m_A;

	private readonly int m_A;

	[CompilerGenerated]
	private Series m_A;

	[CompilerGenerated]
	private DataLabel m_A;

	private int m_B;

	private int C;

	[CompilerGenerated]
	private int D;

	private double m_A;

	private string m_B;

	[CompilerGenerated]
	private Range m_A;

	private bool m_A;

	private bool m_B;

	internal Series Series
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

	internal DataLabel Label
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

	public int StartPeriod
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(VH.A(52754));
		}
	}

	public int EndPeriod
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			A(VH.A(52777));
		}
	}

	private int MaxPeriod
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	public double GrowthRate
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(52796));
			GrowthRateText = Strings.Format(value, this.m_A);
		}
	}

	public string GrowthRateText
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(VH.A(52817));
		}
	}

	internal Range GrowthRange
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

	public bool ManualEntryIsEnabled
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(52846));
		}
	}

	public bool GrowthRateIsLinked
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(VH.A(52887));
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
	}

	public Arrow(Series ser)
	{
		this.m_A = VH.A(52949);
		this.m_A = ColorTranslator.ToOle(Color.White);
		this.m_A = 0.0;
		this.m_A = true;
		Series = ser;
	}

	public Arrow(Series ser, int intStart, int intEnd)
	{
		this.m_A = VH.A(52949);
		this.m_A = ColorTranslator.ToOle(Color.White);
		this.m_A = 0.0;
		this.m_A = true;
		Series = ser;
		StartPeriod = intStart;
		EndPeriod = intEnd;
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
			switch (7)
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

	internal void A(double A, string B)
	{
		GrowthRate = A;
		this.B(B);
	}

	internal void A(List<double> A, int B, double C, double D, ArrowOptions E)
	{
		if (StartPeriod >= EndPeriod)
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
			this.A(A, C, D, E);
			if (E.Rotate)
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
				this.A();
			}
			if (GrowthRange != null)
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
				this.A(A, B);
				this.B(E.Format);
				return;
			}
		}
	}

	internal void A(List<double> A, int B)
	{
		GrowthRate = this.A(A, B);
	}

	internal double A(List<double> A, int B)
	{
		return checked(Math.Pow(A[EndPeriod - 1] / A[StartPeriod - 1], (double)B / (double)(EndPeriod - StartPeriod))) - 1.0;
	}

	internal void A(List<double> A, double B, double C, ArrowOptions D)
	{
		int startPeriod = StartPeriod;
		int endPeriod = EndPeriod;
		checked
		{
			double num = A[startPeriod - 1];
			double num2 = A[endPeriod - 1];
			double num3 = 0.0;
			int num4 = startPeriod - 1;
			int num5 = endPeriod - 1;
			for (int i = num4; i <= num5; i++)
			{
				if (A[i] > num3)
				{
					num3 = A[i];
				}
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
				float num6 = (float)((double)StartPeriod + (double)(EndPeriod - StartPeriod) / 2.0);
				Series series = Series;
				if (D.LineType == ArrowType.Elbow)
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
					double num7 = num3 + B + C;
					series.Values = new double[5]
					{
						num + B,
						num7,
						num7,
						num7,
						num2 + B
					};
					series.XValues = new double[5] { startPeriod, startPeriod, num6, endPeriod, endPeriod };
				}
				else
				{
					if (D.LineType == ArrowType.Angled)
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
						series.Values = new double[3]
						{
							num + B,
							(num + B + num2 + B) / 2.0,
							num2 + B
						};
					}
					else
					{
						series.Values = new double[3]
						{
							num3 + B,
							num3 + B,
							num3 + B
						};
					}
					series.XValues = new double[3] { startPeriod, num6, endPeriod };
				}
				series = null;
				return;
			}
		}
	}

	internal void B(string A)
	{
		if (Label != null)
		{
			try
			{
				Label.Text = Strings.Format(GrowthRate, A);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	internal void A(int A)
	{
		try
		{
			if (Conversions.ToBoolean(NewLateBinding.LateGet(Series.Points(A), null, VH.A(52924), new object[0], null, null, null)))
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
						Label = (DataLabel)Series.DataLabels(A);
						return;
					}
				}
			}
			Label = null;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			Label = null;
			ProjectData.ClearProjectError();
		}
	}

	internal void A(ArrowOptions A, int B)
	{
		Series series = Series;
		if (!series.HasDataLabels)
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
			series.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			series.HasLeaderLines = false;
			Microsoft.Office.Interop.Excel.DataLabels dataLabels = (Microsoft.Office.Interop.Excel.DataLabels)series.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value));
			dataLabels.Format.TextFrame2.WordWrap = MsoTriState.msoFalse;
			dataLabels.Format.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
			if (A.LabelPosition == CagrLabelPosition.Inline)
			{
				dataLabels.Position = XlDataLabelPosition.xlLabelPositionCenter;
			}
			else if (A.LabelPosition == CagrLabelPosition.AboveLine)
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
				dataLabels.Position = XlDataLabelPosition.xlLabelPositionAbove;
			}
			else
			{
				dataLabels.Position = XlDataLabelPosition.xlLabelPositionBelow;
			}
			for (int i = dataLabels.Count; i >= 1; i = checked(i + -1))
			{
				if (i == B)
				{
					Label = dataLabels.Item(i);
					this.B(A.Format);
				}
				else
				{
					dataLabels.Item(i).Delete();
				}
			}
			dataLabels = null;
		}
		series = null;
	}

	internal void A()
	{
		Microsoft.Office.Interop.Excel.Point obj = (Microsoft.Office.Interop.Excel.Point)Series.Points(1);
		float num = (float)obj.Left;
		float num2 = (float)obj.Top;
		_ = null;
		Microsoft.Office.Interop.Excel.Point obj2 = (Microsoft.Office.Interop.Excel.Point)Series.Points(((Points)Series.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).Count);
		float num3 = (float)obj2.Left;
		float num4 = (float)obj2.Top;
		_ = null;
		float num5 = (float)(Math.Atan2(num4 - num2, num3 - num) * 57.2957795);
		Label.Orientation = checked(-(int)Math.Round(num5));
	}

	internal void B(int A)
	{
		Label.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = A;
	}

	internal void A(bool A)
	{
		Label.Format.TextFrame2.TextRange.Font.Bold = (MsoTriState)(0 - (A ? 1 : 0));
	}
}
