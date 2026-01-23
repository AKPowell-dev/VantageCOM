using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes.SelectMatch;

[DesignerGenerated]
public sealed class wpfSelectMatch : UserControl, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class YD
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public wpfSelectMatch A;

		public YD(YD A)
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
		internal bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			int result;
			if (A != this.A)
			{
				MySettings settings = PB.Settings;
				if (!settings.SelectMatchWidth || A.Width == this.A.Width)
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
					if (!settings.SelectMatchHeight)
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
					}
					else if (A.Height != this.A.Height)
					{
						goto IL_0327;
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
					if (!settings.SelectMatchTop)
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
					}
					else if (A.Top != this.A.Top)
					{
						goto IL_0327;
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
					if (!settings.SelectMatchLeft)
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
					}
					else if (A.Left != this.A.Left)
					{
						goto IL_0327;
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
					if ((!settings.SelectMatchBottom || this.A.A(this.A, A)) && (!settings.SelectMatchRight || this.A.B(this.A, A)))
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
						if (!settings.SelectMatchRotation)
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
						}
						else if (A.Rotation != this.A.Rotation)
						{
							goto IL_0327;
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
						if (!settings.SelectMatchShapeType)
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
						}
						else if (!this.A.F(this.A, A))
						{
							goto IL_0327;
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
						if (!settings.SelectMatchAdjustments)
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
						}
						else if (!this.A.G(this.A, A))
						{
							goto IL_0327;
						}
						if (!settings.SelectMatchFreeformPoints)
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
						}
						else if (!this.A.H(this.A, A))
						{
							goto IL_0327;
						}
						if (!settings.SelectMatchFill)
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
						}
						else if (!this.A.D(this.A, A))
						{
							goto IL_0327;
						}
						if (!settings.SelectMatchFont)
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
						}
						else if (!this.A.C(this.A, A))
						{
							goto IL_0327;
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
						if (!settings.SelectMatchBorder)
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
						}
						else if (!this.A.E(this.A, A))
						{
							goto IL_0327;
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
						if (!settings.SelectMatchZOrderAbove)
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
						}
						else if (A.ZOrderPosition <= this.A.ZOrderPosition)
						{
							goto IL_0327;
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
						result = ((!settings.SelectMatchZOrderBelow || A.ZOrderPosition < this.A.ZOrderPosition) ? 1 : 0);
						goto IL_0328;
					}
				}
				goto IL_0327;
			}
			return true;
			IL_0328:
			return (byte)result != 0;
			IL_0327:
			result = 0;
			goto IL_0328;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly int m_A;

	private int m_B;

	private int m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_A;

	[AccessedThroughProperty("chkHeight")]
	[CompilerGenerated]
	private CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWidth")]
	private CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTop")]
	private CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLeft")]
	private CheckBox m_D;

	[AccessedThroughProperty("chkBottom")]
	[CompilerGenerated]
	private CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRight")]
	private CheckBox m_F;

	[AccessedThroughProperty("chkFont")]
	[CompilerGenerated]
	private CheckBox m_G;

	[AccessedThroughProperty("chkFill")]
	[CompilerGenerated]
	private CheckBox m_H;

	[AccessedThroughProperty("chkBorder")]
	[CompilerGenerated]
	private CheckBox I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShapeType")]
	private CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAdjustments")]
	private CheckBox K;

	[AccessedThroughProperty("chkPoints")]
	[CompilerGenerated]
	private CheckBox L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAbove")]
	private CheckBox M;

	[AccessedThroughProperty("chkBelow")]
	[CompilerGenerated]
	private CheckBox N;

	[AccessedThroughProperty("chkRotation")]
	[CompilerGenerated]
	private CheckBox O;

	[AccessedThroughProperty("tbPlural1")]
	[CompilerGenerated]
	private Run m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbPlural2")]
	private Run m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSelect")]
	private Button m_A;

	private bool m_A;

	public int NumReferenceShapes
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(75133));
			Button button = btnSelect;
			int isEnabled;
			if (value > 0)
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
				isEnabled = ((NumProperties > 0) ? 1 : 0);
			}
			else
			{
				isEnabled = 0;
			}
			button.IsEnabled = (byte)isEnabled != 0;
		}
	}

	public int NumProperties
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(AH.A(75170));
			btnSelect.IsEnabled = value > 0 && NumReferenceShapes > 0;
		}
	}

	internal virtual ScrollViewer scroller
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

	internal virtual CheckBox chkHeight
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

	internal virtual CheckBox chkWidth
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

	internal virtual CheckBox chkTop
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

	internal virtual CheckBox chkLeft
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

	internal virtual CheckBox chkBottom
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

	internal virtual CheckBox chkRight
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

	internal virtual CheckBox chkFont
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

	internal virtual CheckBox chkFill
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

	internal virtual CheckBox chkBorder
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

	internal virtual CheckBox chkShapeType
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

	internal virtual CheckBox chkAdjustments
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

	internal virtual CheckBox chkPoints
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

	internal virtual CheckBox chkAbove
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

	internal virtual CheckBox chkBelow
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

	internal virtual CheckBox chkRotation
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

	internal virtual Run tbPlural1
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

	internal virtual Run tbPlural2
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

	internal virtual Button btnSelect
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
			RoutedEventHandler value2 = SelectShapes;
			Button button = this.m_A;
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
				switch (5)
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
		}
	}

	public wpfSelectMatch()
	{
		base.Unloaded += wpfSelectMatch_Unloaded;
		this.m_A = 2;
		this.m_B = 0;
		this.m_C = 0;
		InitializeComponent();
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
			switch (2)
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

	private void wpfSelectMatch_Unloaded(object sender, RoutedEventArgs e)
	{
		C();
		E();
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	public void ShowPane()
	{
		A();
		B();
		D();
		base.SizeChanged += PaneSizeChanged;
		int a;
		try
		{
			a = A(NG.A.Application.ActiveWindow.Selection);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			a = 0;
			ProjectData.ClearProjectError();
		}
		A(a);
	}

	public void HidePane()
	{
		C();
		E();
		base.SizeChanged -= PaneSizeChanged;
	}

	private void A()
	{
		MySettings settings = PB.Settings;
		chkHeight.IsChecked = settings.SelectMatchHeight;
		chkWidth.IsChecked = settings.SelectMatchWidth;
		chkTop.IsChecked = settings.SelectMatchTop;
		chkLeft.IsChecked = settings.SelectMatchLeft;
		chkBottom.IsChecked = settings.SelectMatchBottom;
		chkRight.IsChecked = settings.SelectMatchRight;
		chkRotation.IsChecked = settings.SelectMatchRotation;
		chkFont.IsChecked = settings.SelectMatchFont;
		chkFill.IsChecked = settings.SelectMatchFill;
		chkBorder.IsChecked = settings.SelectMatchBorder;
		chkAbove.IsChecked = settings.SelectMatchZOrderAbove;
		chkBelow.IsChecked = settings.SelectMatchZOrderBelow;
		chkShapeType.IsChecked = settings.SelectMatchShapeType;
		chkAdjustments.IsChecked = settings.SelectMatchAdjustments;
		chkPoints.IsChecked = settings.SelectMatchFreeformPoints;
		settings = null;
	}

	private void HeightCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchHeight = chkHeight.IsChecked.Value;
		A(chkHeight.IsChecked.Value);
	}

	private void WidthCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchWidth = chkWidth.IsChecked.Value;
		A(chkWidth.IsChecked.Value);
	}

	private void TopCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchTop = chkTop.IsChecked.Value;
		A(chkTop.IsChecked.Value);
	}

	private void LeftCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchLeft = chkLeft.IsChecked.Value;
		A(chkLeft.IsChecked.Value);
	}

	private void BottomCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchBottom = chkBottom.IsChecked.Value;
		A(chkBottom.IsChecked.Value);
	}

	private void RightCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchRight = chkRight.IsChecked.Value;
		A(chkRight.IsChecked.Value);
	}

	private void RotationCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchRotation = chkRotation.IsChecked.Value;
		A(chkRotation.IsChecked.Value);
	}

	private void FontCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchFont = chkFont.IsChecked.Value;
		A(chkFont.IsChecked.Value);
	}

	private void FillCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchFill = chkFill.IsChecked.Value;
		A(chkFill.IsChecked.Value);
	}

	private void BorderCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchBorder = chkBorder.IsChecked.Value;
		A(chkBorder.IsChecked.Value);
	}

	private void AboveCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchZOrderAbove = chkAbove.IsChecked.Value;
		A(chkAbove.IsChecked.Value);
	}

	private void BelowCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchZOrderBelow = chkBelow.IsChecked.Value;
		A(chkBelow.IsChecked.Value);
	}

	private void ShapeTypeCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchShapeType = chkShapeType.IsChecked.Value;
		A(chkShapeType.IsChecked.Value);
	}

	private void AdjustmentsCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchAdjustments = chkAdjustments.IsChecked.Value;
		A(chkAdjustments.IsChecked.Value);
	}

	private void PointsCheckedChanged(object sender, RoutedEventArgs e)
	{
		PB.Settings.SelectMatchFreeformPoints = chkPoints.IsChecked.Value;
		A(chkPoints.IsChecked.Value);
	}

	private void B()
	{
		chkHeight.Checked += HeightCheckedChanged;
		chkHeight.Unchecked += HeightCheckedChanged;
		chkWidth.Checked += WidthCheckedChanged;
		chkWidth.Unchecked += WidthCheckedChanged;
		chkTop.Checked += TopCheckedChanged;
		chkTop.Unchecked += TopCheckedChanged;
		chkLeft.Checked += LeftCheckedChanged;
		chkLeft.Unchecked += LeftCheckedChanged;
		chkBottom.Checked += BottomCheckedChanged;
		chkBottom.Unchecked += BottomCheckedChanged;
		chkRight.Checked += RightCheckedChanged;
		chkRight.Unchecked += RightCheckedChanged;
		chkRotation.Checked += RotationCheckedChanged;
		chkRotation.Unchecked += RotationCheckedChanged;
		chkFont.Checked += FontCheckedChanged;
		chkFont.Unchecked += FontCheckedChanged;
		chkFill.Checked += FillCheckedChanged;
		chkFill.Unchecked += FillCheckedChanged;
		chkBorder.Checked += BorderCheckedChanged;
		chkBorder.Unchecked += BorderCheckedChanged;
		chkAbove.Checked += AboveCheckedChanged;
		chkAbove.Unchecked += AboveCheckedChanged;
		chkBelow.Checked += BelowCheckedChanged;
		chkBelow.Unchecked += BelowCheckedChanged;
		chkShapeType.Checked += ShapeTypeCheckedChanged;
		chkShapeType.Unchecked += ShapeTypeCheckedChanged;
		chkAdjustments.Checked += AdjustmentsCheckedChanged;
		chkAdjustments.Unchecked += AdjustmentsCheckedChanged;
		chkPoints.Checked += PointsCheckedChanged;
		chkPoints.Unchecked += PointsCheckedChanged;
	}

	private void C()
	{
		chkHeight.Checked -= HeightCheckedChanged;
		chkHeight.Unchecked -= HeightCheckedChanged;
		chkWidth.Checked -= WidthCheckedChanged;
		chkWidth.Unchecked -= WidthCheckedChanged;
		chkTop.Checked -= TopCheckedChanged;
		chkTop.Unchecked -= TopCheckedChanged;
		chkLeft.Checked -= LeftCheckedChanged;
		chkLeft.Unchecked -= LeftCheckedChanged;
		chkBottom.Checked -= BottomCheckedChanged;
		chkBottom.Unchecked -= BottomCheckedChanged;
		chkRight.Checked -= RightCheckedChanged;
		chkRight.Unchecked -= RightCheckedChanged;
		chkRotation.Checked -= RotationCheckedChanged;
		chkRotation.Unchecked -= RotationCheckedChanged;
		chkFont.Checked -= FontCheckedChanged;
		chkFont.Unchecked -= FontCheckedChanged;
		chkFill.Checked -= FillCheckedChanged;
		chkFill.Unchecked -= FillCheckedChanged;
		chkBorder.Checked -= BorderCheckedChanged;
		chkBorder.Unchecked -= BorderCheckedChanged;
		chkAbove.Checked -= AboveCheckedChanged;
		chkAbove.Unchecked -= AboveCheckedChanged;
		chkBelow.Checked -= BelowCheckedChanged;
		chkBelow.Unchecked -= BelowCheckedChanged;
		chkShapeType.Checked -= ShapeTypeCheckedChanged;
		chkShapeType.Unchecked -= ShapeTypeCheckedChanged;
		chkAdjustments.Checked -= AdjustmentsCheckedChanged;
		chkAdjustments.Unchecked -= AdjustmentsCheckedChanged;
		chkPoints.Checked -= PointsCheckedChanged;
		chkPoints.Unchecked -= PointsCheckedChanged;
	}

	private void A(bool A)
	{
		checked
		{
			if (A)
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
				NumProperties++;
			}
			else
			{
				NumProperties--;
			}
			Run run = tbPlural1;
			string text;
			if (NumProperties != 1)
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
				text = AH.A(75197);
			}
			else
			{
				text = AH.A(8172);
			}
			run.Text = text;
		}
	}

	private void D()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void E()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void A(Selection A)
	{
		this.A(this.A(A));
	}

	private int A(Selection A)
	{
		try
		{
			if (A.Type == PpSelectionType.ppSelectionShapes)
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
						return A.ShapeRange.Count;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return 0;
	}

	private void A(int A)
	{
		NumReferenceShapes = A;
		Run run = tbPlural2;
		object text;
		if (A != 1)
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
			text = AH.A(8154);
		}
		else
		{
			text = "";
		}
		run.Text = (string)text;
	}

	private void SelectShapes(object sender, RoutedEventArgs e)
	{
		if (!Licensing.AllowAdvancedShapeOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		YD yD = default(YD);
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			if (application.Windows.Count > 0)
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
				Selection selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
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
					if (!selection.HasChildShapeRange)
					{
						Slide slide = selection.SlideRange[1];
						List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
						List<Microsoft.Office.Interop.PowerPoint.Shape> list2 = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
						try
						{
							enumerator = slide.Shapes.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape item = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								list.Add(item);
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
						foreach (Microsoft.Office.Interop.PowerPoint.Shape item2 in selection.ShapeRange)
						{
							yD = new YD(yD);
							yD.A = this;
							yD.A = item2;
							list2.AddRange(list.Where(yD.A).ToList());
							yD.A = null;
						}
						list2 = list2.Distinct().ToList();
						if (list2.Count > 0)
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
							try
							{
								using List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator3 = list2.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									enumerator3.Current.Select(MsoTriState.msoFalse);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_01c7;
									}
									continue;
									end_IL_01c7:
									break;
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								B(ex2.Message);
								ProjectData.ClearProjectError();
							}
						}
						else
						{
							D(AH.A(75204));
						}
						slide = null;
						list = null;
						list2 = null;
					}
					else
					{
						C(AH.A(75255));
					}
				}
				selection = null;
			}
			application = null;
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(75085));
			return;
		}
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		return Math.Round(A.Top + A.Height - B.Top - B.Height, this.m_A) == 0.0;
	}

	private bool B(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		return Math.Round(A.Left + A.Width - B.Left - B.Width, this.m_A) == 0.0;
	}

	private bool C(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		if (A.TextFrame.TextRange.Font.Color.RGB == B.TextFrame.TextRange.Font.Color.RGB)
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
			if (Operators.CompareString(A.TextFrame2.TextRange.Font.Name, B.TextFrame2.TextRange.Font.Name, TextCompare: false) == 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return A.TextFrame2.TextRange.Font.Size == B.TextFrame2.TextRange.Font.Size;
					}
				}
			}
		}
		return false;
	}

	private bool D(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		Microsoft.Office.Interop.PowerPoint.FillFormat fill = A.Fill;
		Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = B.Fill;
		try
		{
			int result;
			if (fill.ForeColor.RGB == fill2.ForeColor.RGB)
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
				if (fill.BackColor.RGB == fill2.BackColor.RGB)
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
					if (fill.Pattern == fill2.Pattern)
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
						result = ((fill.Transparency == fill2.Transparency) ? 1 : 0);
						goto IL_0094;
					}
				}
			}
			result = 0;
			goto IL_0094;
			IL_0094:
			return (byte)result != 0;
		}
		finally
		{
			fill = null;
			fill2 = null;
		}
	}

	private bool E(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
		Microsoft.Office.Interop.PowerPoint.LineFormat line2 = B.Line;
		try
		{
			int result;
			if (line.Visible == line2.Visible && line.ForeColor.RGB == line2.ForeColor.RGB)
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
				if (line.BackColor.RGB == line2.BackColor.RGB && line.Weight == line2.Weight)
				{
					result = ((line.Style == line2.Style) ? 1 : 0);
					goto IL_009b;
				}
			}
			result = 0;
			goto IL_009b;
			IL_009b:
			return (byte)result != 0;
		}
		finally
		{
			line = null;
			line2 = null;
		}
	}

	private bool F(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		if (A.Type == B.Type)
		{
			if (A.Type == MsoShapeType.msoAutoShape)
			{
				return A.AutoShapeType == B.AutoShapeType;
			}
			return true;
		}
		return false;
	}

	private bool G(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		try
		{
			if (A.Type == MsoShapeType.msoAutoShape)
			{
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
					if (B.Type != MsoShapeType.msoAutoShape)
					{
						break;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						if (A.AutoShapeType != B.AutoShapeType)
						{
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							if (A.Adjustments.Count <= 0 || A.Adjustments.Count != B.Adjustments.Count)
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
								int count = A.Adjustments.Count;
								for (int i = 1; i <= count; i = checked(i + 1))
								{
									if (A.Adjustments[i] != B.Adjustments[i])
									{
										return false;
									}
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									return true;
								}
							}
						}
						break;
					}
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
		return false;
	}

	private bool H(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		try
		{
			if (A.Type == MsoShapeType.msoFreeform)
			{
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
					if (B.Type != MsoShapeType.msoFreeform)
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
						if (A.Nodes.Count <= 0)
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
							if (A.Nodes.Count != B.Nodes.Count)
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
								int count = A.Nodes.Count;
								for (int i = 1; i <= count; i = checked(i + 1))
								{
									float num = Conversions.ToSingle(NewLateBinding.LateGet(null, typeof(Math), AH.A(75352), new object[2]
									{
										Operators.SubtractObject(NewLateBinding.LateIndexGet(A.Nodes[i].Points, new object[2] { 1, 1 }, null), A.Left),
										this.m_A
									}, null, null, null));
									float num2 = Conversions.ToSingle(NewLateBinding.LateGet(null, typeof(Math), AH.A(75352), new object[2]
									{
										Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Nodes[i].Points, new object[2] { 1, 1 }, null), B.Left),
										this.m_A
									}, null, null, null));
									if (num != num2)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												return false;
											}
										}
									}
									float num3 = Conversions.ToSingle(NewLateBinding.LateGet(null, typeof(Math), AH.A(75352), new object[2]
									{
										Operators.SubtractObject(NewLateBinding.LateIndexGet(A.Nodes[i].Points, new object[2] { 1, 2 }, null), A.Top),
										this.m_A
									}, null, null, null));
									float num4 = Conversions.ToSingle(NewLateBinding.LateGet(null, typeof(Math), AH.A(75352), new object[2]
									{
										Operators.SubtractObject(NewLateBinding.LateIndexGet(B.Nodes[i].Points, new object[2] { 1, 2 }, null), B.Top),
										this.m_A
									}, null, null, null));
									if (num3 == num4)
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
										return false;
									}
								}
								return true;
							}
						}
						break;
					}
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
		return false;
	}

	private void B(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_A)
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
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(75363), UriKind.Relative);
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
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					scroller = (ScrollViewer)target;
					return;
				}
			}
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
					chkHeight = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkWidth = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkTop = (CheckBox)target;
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
					chkLeft = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			chkBottom = (CheckBox)target;
			return;
		}
		if (connectionId == 7)
		{
			chkRight = (CheckBox)target;
			return;
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
					chkFont = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			chkFill = (CheckBox)target;
			return;
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkBorder = (CheckBox)target;
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
					chkShapeType = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkAdjustments = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkPoints = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			chkAbove = (CheckBox)target;
			return;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkBelow = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkRotation = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			tbPlural1 = (Run)target;
			return;
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
					tbPlural2 = (Run)target;
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
					btnSelect = (Button)target;
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
