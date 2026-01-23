using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Shapes;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.FormatPainter;

[DesignerGenerated]
public sealed class FormatTree : UserControl, IComponentConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> A;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> A;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> B;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> B;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> C;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> C;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> D;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> D;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> E;

		public static Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> E;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Key == MsoPictureEffectType.msoEffectSharpenSoften;
		}

		[SpecialName]
		internal List<float> A(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal bool B(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Key == MsoPictureEffectType.msoEffectBrightnessContrast;
		}

		[SpecialName]
		internal List<float> B(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal bool C(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Key == MsoPictureEffectType.msoEffectBrightnessContrast;
		}

		[SpecialName]
		internal List<float> C(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal bool D(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Key == MsoPictureEffectType.msoEffectSaturation;
		}

		[SpecialName]
		internal List<float> D(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Value;
		}

		[SpecialName]
		internal bool E(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Key == MsoPictureEffectType.msoEffectColorTemperature;
		}

		[SpecialName]
		internal List<float> E(KeyValuePair<MsoPictureEffectType, List<float>> A)
		{
			return A.Value;
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

	[CompilerGenerated]
	[AccessedThroughProperty("expLayout")]
	private Polygon m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLayout")]
	private CheckBox m_A;

	[AccessedThroughProperty("gridLayout")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("chkHeight")]
	[CompilerGenerated]
	private CheckBox m_B;

	[AccessedThroughProperty("txtHeight")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("chkWidth")]
	[CompilerGenerated]
	private CheckBox m_C;

	[AccessedThroughProperty("txtWidth")]
	[CompilerGenerated]
	private TextBlock m_B;

	[AccessedThroughProperty("chkLockAspectRatio")]
	[CompilerGenerated]
	private CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLockAspectRatio")]
	private TextBlock m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPositionY")]
	private CheckBox m_E;

	[AccessedThroughProperty("txtPositionY")]
	[CompilerGenerated]
	private TextBlock m_D;

	[AccessedThroughProperty("radTop")]
	[CompilerGenerated]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtTop")]
	private TextBlock m_E;

	[AccessedThroughProperty("radBottom")]
	[CompilerGenerated]
	private RadioButton m_B;

	[AccessedThroughProperty("txtBottom")]
	[CompilerGenerated]
	private TextBlock m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("radMidPointY")]
	private RadioButton m_C;

	[AccessedThroughProperty("txtMidpointY")]
	[CompilerGenerated]
	private TextBlock m_G;

	[AccessedThroughProperty("chkPositionX")]
	[CompilerGenerated]
	private CheckBox m_F;

	[AccessedThroughProperty("txtPositionX")]
	[CompilerGenerated]
	private TextBlock m_H;

	[AccessedThroughProperty("radLeft")]
	[CompilerGenerated]
	private RadioButton m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLeft")]
	private TextBlock m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("radRight")]
	private RadioButton m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("txtRight")]
	private TextBlock m_J;

	[CompilerGenerated]
	[AccessedThroughProperty("radMidPointX")]
	private RadioButton m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("txtMidpointX")]
	private TextBlock m_K;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRotation")]
	private CheckBox m_G;

	[AccessedThroughProperty("txtRotation")]
	[CompilerGenerated]
	private TextBlock m_L;

	[AccessedThroughProperty("expLine")]
	[CompilerGenerated]
	private Polygon m_B;

	[AccessedThroughProperty("chkLine")]
	[CompilerGenerated]
	private CheckBox m_H;

	[AccessedThroughProperty("gridLine")]
	[CompilerGenerated]
	private Grid m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLineColor")]
	private CheckBox m_I;

	[AccessedThroughProperty("rectLineColor")]
	[CompilerGenerated]
	private System.Windows.Shapes.Rectangle m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLineColor")]
	private TextBlock m_M;

	[AccessedThroughProperty("chkLineWeight")]
	[CompilerGenerated]
	private CheckBox m_J;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLineWeight")]
	private TextBlock m_N;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLineStyle")]
	private CheckBox m_K;

	[AccessedThroughProperty("txtLineStyle")]
	[CompilerGenerated]
	private TextBlock m_O;

	[CompilerGenerated]
	[AccessedThroughProperty("expFill")]
	private Polygon m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFill")]
	private CheckBox m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("gridFill")]
	private Grid m_C;

	[AccessedThroughProperty("chkFillColor")]
	[CompilerGenerated]
	private CheckBox m_M;

	[AccessedThroughProperty("rectFillColor")]
	[CompilerGenerated]
	private System.Windows.Shapes.Rectangle m_B;

	[AccessedThroughProperty("txtFillColor")]
	[CompilerGenerated]
	private TextBlock m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFillType")]
	private CheckBox m_N;

	[AccessedThroughProperty("txtFillType")]
	[CompilerGenerated]
	private TextBlock m_Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFillTransparency")]
	private CheckBox m_O;

	[AccessedThroughProperty("txtFillTransparency")]
	[CompilerGenerated]
	private TextBlock m_R;

	[AccessedThroughProperty("expFont")]
	[CompilerGenerated]
	private Polygon m_D;

	[AccessedThroughProperty("chkFont")]
	[CompilerGenerated]
	private CheckBox m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("gridFont")]
	private Grid m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFontColor")]
	private CheckBox m_Q;

	[AccessedThroughProperty("rectFontColor")]
	[CompilerGenerated]
	private System.Windows.Shapes.Rectangle m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFontColor")]
	private TextBlock m_S;

	[AccessedThroughProperty("chkFontSize")]
	[CompilerGenerated]
	private CheckBox m_R;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFontSize")]
	private TextBlock m_T;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFontName")]
	private CheckBox m_S;

	[AccessedThroughProperty("txtFontName")]
	[CompilerGenerated]
	private TextBlock m_U;

	[CompilerGenerated]
	[AccessedThroughProperty("chkDecoration")]
	private CheckBox m_T;

	[AccessedThroughProperty("txtDecoration")]
	[CompilerGenerated]
	private TextBlock m_V;

	[CompilerGenerated]
	[AccessedThroughProperty("expTextBox")]
	private Polygon m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTextBox")]
	private CheckBox m_U;

	[AccessedThroughProperty("gridTextBox")]
	[CompilerGenerated]
	private Grid m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBullets")]
	private CheckBox m_V;

	[AccessedThroughProperty("txtBullets")]
	[CompilerGenerated]
	private TextBlock m_W;

	[AccessedThroughProperty("chkIndents")]
	[CompilerGenerated]
	private CheckBox m_W;

	[AccessedThroughProperty("txtIndents")]
	[CompilerGenerated]
	private TextBlock m_X;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLineSpacing")]
	private CheckBox m_X;

	[CompilerGenerated]
	[AccessedThroughProperty("txtLineSpacing")]
	private TextBlock m_Y;

	[CompilerGenerated]
	[AccessedThroughProperty("chkMargins")]
	private CheckBox m_Y;

	[AccessedThroughProperty("txtMargins")]
	[CompilerGenerated]
	private TextBlock m_Z;

	[AccessedThroughProperty("chkAutoSize")]
	[CompilerGenerated]
	private CheckBox m_Z;

	[AccessedThroughProperty("txtAutoSize")]
	[CompilerGenerated]
	private TextBlock m_AB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWordWrap")]
	private CheckBox m_AB;

	[AccessedThroughProperty("txtWordWrap")]
	[CompilerGenerated]
	private TextBlock m_BB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAlignH")]
	private CheckBox m_BB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtAlignH")]
	private TextBlock m_CB;

	[AccessedThroughProperty("chkAlignV")]
	[CompilerGenerated]
	private CheckBox m_CB;

	[AccessedThroughProperty("txtAlignV")]
	[CompilerGenerated]
	private TextBlock m_DB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkOrientation")]
	private CheckBox m_DB;

	[AccessedThroughProperty("txtOrientation")]
	[CompilerGenerated]
	private TextBlock m_EB;

	[AccessedThroughProperty("expAutoShape")]
	[CompilerGenerated]
	private Polygon m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAutoShape")]
	private CheckBox m_EB;

	[AccessedThroughProperty("gridAutoShape")]
	[CompilerGenerated]
	private Grid m_F;

	[AccessedThroughProperty("chkAutoShapeType")]
	[CompilerGenerated]
	private CheckBox FB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtAutoShapeType")]
	private TextBlock FB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAdjustments")]
	private CheckBox GB;

	[AccessedThroughProperty("txtAdjustments")]
	[CompilerGenerated]
	private TextBlock GB;

	[AccessedThroughProperty("expPicture")]
	[CompilerGenerated]
	private Polygon m_G;

	[AccessedThroughProperty("chkPicture")]
	[CompilerGenerated]
	private CheckBox HB;

	[AccessedThroughProperty("gridPicture")]
	[CompilerGenerated]
	private Grid m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPictureScale")]
	private CheckBox IB;

	[AccessedThroughProperty("txtPictureScale")]
	[CompilerGenerated]
	private TextBlock HB;

	[AccessedThroughProperty("chkScaleHeight")]
	[CompilerGenerated]
	private CheckBox JB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtScaleHeight")]
	private TextBlock IB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkScaleWidth")]
	private CheckBox KB;

	[AccessedThroughProperty("txtScaleWidth")]
	[CompilerGenerated]
	private TextBlock JB;

	[AccessedThroughProperty("chkSharpness")]
	[CompilerGenerated]
	private CheckBox LB;

	[AccessedThroughProperty("txtSharpness")]
	[CompilerGenerated]
	private TextBlock KB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBrightness")]
	private CheckBox MB;

	[AccessedThroughProperty("txtBrightness")]
	[CompilerGenerated]
	private TextBlock LB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkContrast")]
	private CheckBox NB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtContrast")]
	private TextBlock MB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSaturation")]
	private CheckBox OB;

	[AccessedThroughProperty("txtSaturation")]
	[CompilerGenerated]
	private TextBlock NB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTemperature")]
	private CheckBox PB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtTemperature")]
	private TextBlock OB;

	[CompilerGenerated]
	[AccessedThroughProperty("expEffects")]
	private Polygon m_H;

	[AccessedThroughProperty("chkEffects")]
	[CompilerGenerated]
	private CheckBox QB;

	[CompilerGenerated]
	[AccessedThroughProperty("gridEffects")]
	private Grid m_H;

	[AccessedThroughProperty("chkShapeEffects")]
	[CompilerGenerated]
	private CheckBox RB;

	[AccessedThroughProperty("txtShapeEffects")]
	[CompilerGenerated]
	private TextBlock PB;

	[AccessedThroughProperty("chkShapeShadow")]
	[CompilerGenerated]
	private CheckBox SB;

	[CompilerGenerated]
	[AccessedThroughProperty("rectShapeShadow")]
	private System.Windows.Shapes.Rectangle m_D;

	[AccessedThroughProperty("txtShapeShadow")]
	[CompilerGenerated]
	private TextBlock QB;

	[AccessedThroughProperty("chkShapeReflection")]
	[CompilerGenerated]
	private CheckBox TB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtShapeReflection")]
	private TextBlock RB;

	[AccessedThroughProperty("chkShapeGlow")]
	[CompilerGenerated]
	private CheckBox UB;

	[CompilerGenerated]
	[AccessedThroughProperty("rectShapeGlow")]
	private System.Windows.Shapes.Rectangle m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("txtShapeGlow")]
	private TextBlock SB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShapeSoftEdge")]
	private CheckBox VB;

	[AccessedThroughProperty("txtShapeSoftEdge")]
	[CompilerGenerated]
	private TextBlock TB;

	[AccessedThroughProperty("chkShape3D")]
	[CompilerGenerated]
	private CheckBox WB;

	[AccessedThroughProperty("txtShape3D")]
	[CompilerGenerated]
	private TextBlock UB;

	[AccessedThroughProperty("chkTextEffects")]
	[CompilerGenerated]
	private CheckBox XB;

	[AccessedThroughProperty("txtTextEffects")]
	[CompilerGenerated]
	private TextBlock VB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTextShadow")]
	private CheckBox YB;

	[CompilerGenerated]
	[AccessedThroughProperty("rectTextShadow")]
	private System.Windows.Shapes.Rectangle m_F;

	[AccessedThroughProperty("txtTextShadow")]
	[CompilerGenerated]
	private TextBlock WB;

	[AccessedThroughProperty("chkTextReflection")]
	[CompilerGenerated]
	private CheckBox ZB;

	[CompilerGenerated]
	[AccessedThroughProperty("txtTextReflection")]
	private TextBlock XB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkTextGlow")]
	private CheckBox AC;

	[CompilerGenerated]
	[AccessedThroughProperty("rectTextGlow")]
	private System.Windows.Shapes.Rectangle m_G;

	[AccessedThroughProperty("txtTextGlow")]
	[CompilerGenerated]
	private TextBlock YB;

	[AccessedThroughProperty("chkTextSoftEdge")]
	[CompilerGenerated]
	private CheckBox BC;

	[AccessedThroughProperty("txtTextSoftEdge")]
	[CompilerGenerated]
	private TextBlock ZB;

	[AccessedThroughProperty("chkText3D")]
	[CompilerGenerated]
	private CheckBox CC;

	[AccessedThroughProperty("txtText3D")]
	[CompilerGenerated]
	private TextBlock AC;

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

	internal virtual Polygon expLayout
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

	internal virtual CheckBox chkLayout
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

	internal virtual Grid gridLayout
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
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual TextBlock txtHeight
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
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual TextBlock txtWidth
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

	internal virtual CheckBox chkLockAspectRatio
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

	internal virtual TextBlock txtLockAspectRatio
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

	internal virtual CheckBox chkPositionY
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

	internal virtual TextBlock txtPositionY
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

	internal virtual RadioButton radTop
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

	internal virtual TextBlock txtTop
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

	internal virtual RadioButton radBottom
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

	internal virtual TextBlock txtBottom
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

	internal virtual RadioButton radMidPointY
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

	internal virtual TextBlock txtMidpointY
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

	internal virtual CheckBox chkPositionX
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

	internal virtual TextBlock txtPositionX
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

	internal virtual RadioButton radLeft
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

	internal virtual TextBlock txtLeft
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

	internal virtual RadioButton radRight
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

	internal virtual TextBlock txtRight
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

	internal virtual RadioButton radMidPointX
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

	internal virtual TextBlock txtMidpointX
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

	internal virtual CheckBox chkRotation
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

	internal virtual TextBlock txtRotation
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

	internal virtual Polygon expLine
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

	internal virtual CheckBox chkLine
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
			RoutedEventHandler value2 = RemoveLineCheckBoxHandlers;
			RoutedEventHandler value3 = RemoveLineCheckBoxHandlers;
			CheckBox checkBox = this.m_H;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_H = value;
			checkBox = this.m_H;
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

	internal virtual Grid gridLine
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

	internal virtual CheckBox chkLineColor
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

	internal virtual System.Windows.Shapes.Rectangle rectLineColor
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

	internal virtual TextBlock txtLineColor
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

	internal virtual CheckBox chkLineWeight
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

	internal virtual TextBlock txtLineWeight
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

	internal virtual CheckBox chkLineStyle
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

	internal virtual TextBlock txtLineStyle
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

	internal virtual Polygon expFill
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

	internal virtual CheckBox chkFill
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
			RoutedEventHandler value2 = RemoveFillCheckBoxHandlers;
			RoutedEventHandler value3 = RemoveFillCheckBoxHandlers;
			CheckBox checkBox = this.m_L;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_L = value;
			checkBox = this.m_L;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual Grid gridFill
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

	internal virtual CheckBox chkFillColor
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

	internal virtual System.Windows.Shapes.Rectangle rectFillColor
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

	internal virtual TextBlock txtFillColor
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

	internal virtual CheckBox chkFillType
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

	internal virtual TextBlock txtFillType
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

	internal virtual CheckBox chkFillTransparency
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

	internal virtual TextBlock txtFillTransparency
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

	internal virtual Polygon expFont
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

	internal virtual CheckBox chkFont
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
			RoutedEventHandler value2 = RemoveFontCheckBoxHandlers;
			RoutedEventHandler value3 = RemoveFontCheckBoxHandlers;
			CheckBox checkBox = this.m_P;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_P = value;
			checkBox = this.m_P;
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

	internal virtual Grid gridFont
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

	internal virtual CheckBox chkFontColor
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

	internal virtual System.Windows.Shapes.Rectangle rectFontColor
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

	internal virtual TextBlock txtFontColor
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

	internal virtual CheckBox chkFontSize
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

	internal virtual TextBlock txtFontSize
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

	internal virtual CheckBox chkFontName
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

	internal virtual TextBlock txtFontName
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

	internal virtual CheckBox chkDecoration
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

	internal virtual TextBlock txtDecoration
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

	internal virtual Polygon expTextBox
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

	internal virtual CheckBox chkTextBox
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
			RoutedEventHandler value2 = TextBoxCheckChanged;
			RoutedEventHandler value3 = TextBoxCheckChanged;
			CheckBox checkBox = this.m_U;
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
			this.m_U = value;
			checkBox = this.m_U;
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

	internal virtual Grid gridTextBox
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

	internal virtual CheckBox chkBullets
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

	internal virtual TextBlock txtBullets
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

	internal virtual CheckBox chkIndents
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

	internal virtual TextBlock txtIndents
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

	internal virtual CheckBox chkLineSpacing
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

	internal virtual TextBlock txtLineSpacing
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

	internal virtual CheckBox chkMargins
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

	internal virtual TextBlock txtMargins
	{
		[CompilerGenerated]
		get
		{
			return this.m_Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Z = value;
		}
	}

	internal virtual CheckBox chkAutoSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Z = value;
		}
	}

	internal virtual TextBlock txtAutoSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_AB = value;
		}
	}

	internal virtual CheckBox chkWordWrap
	{
		[CompilerGenerated]
		get
		{
			return this.m_AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_AB = value;
		}
	}

	internal virtual TextBlock txtWordWrap
	{
		[CompilerGenerated]
		get
		{
			return this.m_BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_BB = value;
		}
	}

	internal virtual CheckBox chkAlignH
	{
		[CompilerGenerated]
		get
		{
			return this.m_BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_BB = value;
		}
	}

	internal virtual TextBlock txtAlignH
	{
		[CompilerGenerated]
		get
		{
			return this.m_CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_CB = value;
		}
	}

	internal virtual CheckBox chkAlignV
	{
		[CompilerGenerated]
		get
		{
			return this.m_CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_CB = value;
		}
	}

	internal virtual TextBlock txtAlignV
	{
		[CompilerGenerated]
		get
		{
			return this.m_DB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_DB = value;
		}
	}

	internal virtual CheckBox chkOrientation
	{
		[CompilerGenerated]
		get
		{
			return this.m_DB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_DB = value;
		}
	}

	internal virtual TextBlock txtOrientation
	{
		[CompilerGenerated]
		get
		{
			return this.m_EB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_EB = value;
		}
	}

	internal virtual Polygon expAutoShape
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

	internal virtual CheckBox chkAutoShape
	{
		[CompilerGenerated]
		get
		{
			return this.m_EB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = AutoShapeCheckChanged;
			RoutedEventHandler value3 = AutoShapeCheckChanged;
			CheckBox checkBox = this.m_EB;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_EB = value;
			checkBox = this.m_EB;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual Grid gridAutoShape
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

	internal virtual CheckBox chkAutoShapeType
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

	internal virtual TextBlock txtAutoShapeType
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

	internal virtual CheckBox chkAdjustments
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
			RoutedEventHandler value2 = AdjustmentsCheckChanged;
			CheckBox checkBox = this.GB;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
			}
			this.GB = value;
			checkBox = this.GB;
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
				return;
			}
		}
	}

	internal virtual TextBlock txtAdjustments
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

	internal virtual Polygon expPicture
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

	internal virtual CheckBox chkPicture
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
			RoutedEventHandler value2 = PictureCheckChanged;
			RoutedEventHandler value3 = PictureCheckChanged;
			CheckBox checkBox = this.HB;
			if (checkBox != null)
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

	internal virtual Grid gridPicture
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

	internal virtual CheckBox chkPictureScale
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
			RoutedEventHandler value2 = PictureScaleCheckChanged;
			RoutedEventHandler value3 = PictureScaleCheckChanged;
			CheckBox checkBox = this.IB;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.IB = value;
			checkBox = this.IB;
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

	internal virtual TextBlock txtPictureScale
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

	internal virtual CheckBox chkScaleHeight
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

	internal virtual TextBlock txtScaleHeight
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

	internal virtual CheckBox chkScaleWidth
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

	internal virtual TextBlock txtScaleWidth
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

	internal virtual CheckBox chkSharpness
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

	internal virtual TextBlock txtSharpness
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

	internal virtual CheckBox chkBrightness
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
			this.MB = value;
		}
	}

	internal virtual TextBlock txtBrightness
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

	internal virtual CheckBox chkContrast
	{
		[CompilerGenerated]
		get
		{
			return this.NB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.NB = value;
		}
	}

	internal virtual TextBlock txtContrast
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

	internal virtual CheckBox chkSaturation
	{
		[CompilerGenerated]
		get
		{
			return this.OB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.OB = value;
		}
	}

	internal virtual TextBlock txtSaturation
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

	internal virtual CheckBox chkTemperature
	{
		[CompilerGenerated]
		get
		{
			return this.PB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.PB = value;
		}
	}

	internal virtual TextBlock txtTemperature
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

	internal virtual Polygon expEffects
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

	internal virtual CheckBox chkEffects
	{
		[CompilerGenerated]
		get
		{
			return this.QB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.QB = value;
		}
	}

	internal virtual Grid gridEffects
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

	internal virtual CheckBox chkShapeEffects
	{
		[CompilerGenerated]
		get
		{
			return this.RB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.RB = value;
		}
	}

	internal virtual TextBlock txtShapeEffects
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
			PB = value;
		}
	}

	internal virtual CheckBox chkShapeShadow
	{
		[CompilerGenerated]
		get
		{
			return this.SB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.SB = value;
		}
	}

	internal virtual System.Windows.Shapes.Rectangle rectShapeShadow
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

	internal virtual TextBlock txtShapeShadow
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

	internal virtual CheckBox chkShapeReflection
	{
		[CompilerGenerated]
		get
		{
			return this.TB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.TB = value;
		}
	}

	internal virtual TextBlock txtShapeReflection
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

	internal virtual CheckBox chkShapeGlow
	{
		[CompilerGenerated]
		get
		{
			return this.UB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.UB = value;
		}
	}

	internal virtual System.Windows.Shapes.Rectangle rectShapeGlow
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

	internal virtual TextBlock txtShapeGlow
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
			SB = value;
		}
	}

	internal virtual CheckBox chkShapeSoftEdge
	{
		[CompilerGenerated]
		get
		{
			return this.VB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.VB = value;
		}
	}

	internal virtual TextBlock txtShapeSoftEdge
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

	internal virtual CheckBox chkShape3D
	{
		[CompilerGenerated]
		get
		{
			return this.WB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.WB = value;
		}
	}

	internal virtual TextBlock txtShape3D
	{
		[CompilerGenerated]
		get
		{
			return UB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			UB = value;
		}
	}

	internal virtual CheckBox chkTextEffects
	{
		[CompilerGenerated]
		get
		{
			return this.XB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.XB = value;
		}
	}

	internal virtual TextBlock txtTextEffects
	{
		[CompilerGenerated]
		get
		{
			return VB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			VB = value;
		}
	}

	internal virtual CheckBox chkTextShadow
	{
		[CompilerGenerated]
		get
		{
			return this.YB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.YB = value;
		}
	}

	internal virtual System.Windows.Shapes.Rectangle rectTextShadow
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

	internal virtual TextBlock txtTextShadow
	{
		[CompilerGenerated]
		get
		{
			return WB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			WB = value;
		}
	}

	internal virtual CheckBox chkTextReflection
	{
		[CompilerGenerated]
		get
		{
			return this.ZB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.ZB = value;
		}
	}

	internal virtual TextBlock txtTextReflection
	{
		[CompilerGenerated]
		get
		{
			return XB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			XB = value;
		}
	}

	internal virtual CheckBox chkTextGlow
	{
		[CompilerGenerated]
		get
		{
			return this.AC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.AC = value;
		}
	}

	internal virtual System.Windows.Shapes.Rectangle rectTextGlow
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

	internal virtual TextBlock txtTextGlow
	{
		[CompilerGenerated]
		get
		{
			return YB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			YB = value;
		}
	}

	internal virtual CheckBox chkTextSoftEdge
	{
		[CompilerGenerated]
		get
		{
			return BC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			BC = value;
		}
	}

	internal virtual TextBlock txtTextSoftEdge
	{
		[CompilerGenerated]
		get
		{
			return ZB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			ZB = value;
		}
	}

	internal virtual CheckBox chkText3D
	{
		[CompilerGenerated]
		get
		{
			return CC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			CC = value;
		}
	}

	internal virtual TextBlock txtText3D
	{
		[CompilerGenerated]
		get
		{
			return AC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			AC = value;
		}
	}

	public FormatTree()
	{
		base.Loaded += FormatTree_Loaded;
		this.m_A = AH.A(50090);
		this.m_B = AH.A(139562);
		this.m_C = AH.A(139577);
		this.m_D = AH.A(50099);
		this.m_E = AH.A(139592);
		this.m_F = AH.A(139607);
		InitializeComponent();
	}

	private void FormatTree_Loaded(object sender, RoutedEventArgs e)
	{
		F();
		D();
		B();
		H();
		J();
		L();
		P();
		N();
		T();
		R();
		V();
		X();
		Z();
		BB();
		DB();
	}

	private void A()
	{
		if (Visible)
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
					new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
					return;
				}
			}
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void A(Selection A)
	{
		btnCopy.IsEnabled = Pane.IsSingleShapeSelected(A);
		if (A.Type != PpSelectionType.ppSelectionShapes)
		{
			if (A.Type != PpSelectionType.ppSelectionText)
			{
				btnApply.IsEnabled = false;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		btnApply.IsEnabled = Pane.CopiedProperties != null;
	}

	private void ExpandCollapseFont(object sender, MouseButtonEventArgs e)
	{
		A(expFont, gridFont);
	}

	private void ExpandCollapseFill(object sender, MouseButtonEventArgs e)
	{
		A(expFill, gridFill);
	}

	private void ExpandCollapseLine(object sender, MouseButtonEventArgs e)
	{
		A(expLine, gridLine);
	}

	private void ExpandCollapseTextBox(object sender, MouseButtonEventArgs e)
	{
		A(expTextBox, gridTextBox);
	}

	private void ExpandCollapseLayout(object sender, MouseButtonEventArgs e)
	{
		A(expLayout, gridLayout);
	}

	private void ExpandCollapseAutoShape(object sender, MouseButtonEventArgs e)
	{
		A(expAutoShape, gridAutoShape);
	}

	private void ExpandCollapseEffects(object sender, MouseButtonEventArgs e)
	{
		A(expEffects, gridEffects);
	}

	private void ExpandCollapsePicture(object sender, MouseButtonEventArgs e)
	{
		A(expPicture, gridPicture);
	}

	private void A(Polygon A, Grid B)
	{
		if (B.Visibility == Visibility.Visible)
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
					C(A, B);
					return;
				}
			}
		}
		this.B(A, B);
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
		Pane.CopiedProperties = new Properties(Base.SelectedShapes()[1]);
		PopulateProperties();
	}

	private void btnApply_Click(object sender, RoutedEventArgs e)
	{
		Options options = new Options();
		bool flag = false;
		if (A(chkLine))
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
			Options.myLine line = options.Line;
			line.Color = chkLineColor.IsChecked.Value;
			line.Weight = chkLineWeight.IsChecked.Value;
			line.Style = chkLineStyle.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkFill))
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
			Options.myFill fill = options.Fill;
			fill.Color = chkFillColor.IsChecked.Value;
			fill.Type = chkFillType.IsChecked.Value;
			fill.Transparency = chkFillTransparency.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkFont))
		{
			Options.myFont font = options.Font;
			font.Size = chkFontSize.IsChecked.Value;
			font.Color = chkFontColor.IsChecked.Value;
			font.Name = chkFontName.IsChecked.Value;
			font.Decoration = chkDecoration.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkTextBox))
		{
			Options.myTextBox textBox = options.TextBox;
			textBox.Bullets = chkBullets.IsChecked.Value;
			textBox.Indents = chkIndents.IsChecked.Value;
			textBox.LineSpacing = chkLineSpacing.IsChecked.Value;
			textBox.HorizontalAlignment = chkAlignH.IsChecked.Value;
			textBox.Margins = chkMargins.IsChecked.Value;
			textBox.AutoSize = chkAutoSize.IsChecked.Value;
			textBox.WordWrap = chkWordWrap.IsChecked.Value;
			textBox.VerticalAlignment = chkAlignV.IsChecked.Value;
			textBox.Orientation = chkOrientation.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkLayout))
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
			Options.myLayout layout = options.Layout;
			layout.Height = chkHeight.IsChecked.Value;
			layout.Width = chkWidth.IsChecked.Value;
			layout.LockAspectRatio = chkLockAspectRatio.IsChecked.Value;
			layout.Top = radTop.IsChecked.Value;
			layout.Bottom = radBottom.IsChecked.Value;
			layout.MidpointY = radMidPointY.IsChecked.Value;
			layout.Left = radLeft.IsChecked.Value;
			layout.Right = radRight.IsChecked.Value;
			layout.MidpointX = radMidPointX.IsChecked.Value;
			layout.Rotation = chkRotation.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkShapeEffects))
		{
			Options.myShapeEffects shapeEffects = options.ShapeEffects;
			shapeEffects.Glow = chkShapeGlow.IsChecked.Value;
			shapeEffects.Reflection = chkShapeReflection.IsChecked.Value;
			shapeEffects.Shadow = chkShapeShadow.IsChecked.Value;
			shapeEffects.SoftEdge = chkShapeSoftEdge.IsChecked.Value;
			shapeEffects.ThreeD = chkShape3D.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkTextEffects))
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
			Options.myTextEffects textEffects = options.TextEffects;
			textEffects.Glow = chkTextGlow.IsChecked.Value;
			textEffects.Reflection = chkTextReflection.IsChecked.Value;
			textEffects.Shadow = chkTextShadow.IsChecked.Value;
			textEffects.SoftEdge = chkTextSoftEdge.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkAutoShape))
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
			Options.myAutoShape autoShape = options.AutoShape;
			autoShape.Type = chkAutoShapeType.IsChecked.Value;
			autoShape.Adjustments = chkAdjustments.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (A(chkPicture))
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
			Options.myPicture picture = options.Picture;
			picture.ScaleHeight = chkScaleHeight.IsChecked.Value;
			picture.ScaleWidth = chkScaleWidth.IsChecked.Value;
			picture.Sharpness = chkSharpness.IsChecked.Value;
			picture.Brightness = chkBrightness.IsChecked.Value;
			picture.Contrast = chkContrast.IsChecked.Value;
			picture.Saturation = chkSaturation.IsChecked.Value;
			picture.Temperature = chkTemperature.IsChecked.Value;
			_ = null;
			flag = true;
		}
		if (flag)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					Apply.ToSelection(Pane.CopiedProperties, options);
					return;
				}
			}
		}
		Forms.WarningMessage(AH.A(138834));
	}

	private bool A(CheckBox A)
	{
		bool? isChecked;
		bool? flag = (isChecked = A.IsChecked);
		bool? flag2;
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
			if (isChecked == true)
			{
				flag2 = true;
				goto IL_006b;
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
		if (A.IsChecked.HasValue)
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
			flag2 = isChecked;
		}
		else
		{
			flag2 = true;
		}
		goto IL_006b;
		IL_006b:
		isChecked = flag2;
		return isChecked.Value;
	}

	public void PopulateProperties()
	{
		Properties copiedProperties = Pane.CopiedProperties;
		bool isMetric = RegionInfo.CurrentRegion.IsMetric;
		string value = Conversions.ToString(Operators.CompareString(clsPublish.SystemDecimalSeparator(), AH.A(14417), TextCompare: false) == 0);
		try
		{
			A(copiedProperties);
			B(copiedProperties);
			A(copiedProperties, Conversions.ToBoolean(value));
			C(copiedProperties, Conversions.ToBoolean(value), isMetric);
			D(copiedProperties, Conversions.ToBoolean(value), isMetric);
			B(copiedProperties, Conversions.ToBoolean(value));
			C(copiedProperties, Conversions.ToBoolean(value));
			D(copiedProperties, Conversions.ToBoolean(value));
			F(copiedProperties, Conversions.ToBoolean(value), isMetric);
			if (copiedProperties.Shape.HasTextFrame == MsoTriState.msoTrue)
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
				A(A: true);
				B(expFont, gridFont);
				C(A: true);
				B(expTextBox, gridTextBox);
			}
			else
			{
				A(A: false);
				C(expFont, gridFont);
				G();
				chkFont.IsChecked = false;
				chkFontColor.IsChecked = false;
				chkFontName.IsChecked = false;
				chkFontSize.IsChecked = false;
				chkDecoration.IsChecked = false;
				F();
				C(A: false);
				C(expTextBox, gridTextBox);
				I();
				chkTextBox.IsChecked = false;
				chkBullets.IsChecked = false;
				chkIndents.IsChecked = false;
				chkLineSpacing.IsChecked = false;
				chkMargins.IsChecked = false;
				chkAutoSize.IsChecked = false;
				chkWordWrap.IsChecked = false;
				chkAlignH.IsChecked = false;
				chkAlignV.IsChecked = false;
				chkOrientation.IsChecked = false;
				H();
			}
			if (copiedProperties.Shape.HasTable == MsoTriState.msoTrue)
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
				B(A: false);
				C(expLine, gridLine);
				C();
				chkLine.IsEnabled = false;
				chkLine.IsChecked = false;
				chkLineColor.IsChecked = false;
				chkLineStyle.IsChecked = false;
				chkLineWeight.IsChecked = false;
				B();
				chkShapeGlow.IsChecked = false;
				chkShapeGlow.IsEnabled = false;
				chkShapeSoftEdge.IsChecked = false;
				chkShapeSoftEdge.IsEnabled = false;
			}
			else
			{
				chkLine.IsEnabled = true;
				B(A: true);
				chkShapeGlow.IsEnabled = true;
				chkShapeSoftEdge.IsEnabled = true;
			}
			if (copiedProperties.Shape.HasChart == MsoTriState.msoTrue)
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
				chkShapeReflection.IsEnabled = false;
				chkShapeReflection.IsChecked = false;
			}
			else
			{
				chkShapeReflection.IsEnabled = true;
			}
			if (copiedProperties.Shape.Type == MsoShapeType.msoAutoShape)
			{
				D(A: true);
				B(expAutoShape, gridAutoShape);
			}
			else
			{
				D(A: false);
				C(expAutoShape, gridAutoShape);
				AB();
				chkAutoShape.IsChecked = false;
				chkAutoShapeType.IsChecked = false;
				chkAdjustments.IsChecked = false;
				Z();
			}
			if (copiedProperties.Shape.HasPicture)
			{
				E(A: true);
				B(expPicture, gridPicture);
			}
			else
			{
				E(A: false);
				C(expPicture, gridPicture);
				CB();
				chkPicture.IsChecked = false;
				chkPictureScale.IsChecked = false;
				chkScaleHeight.IsChecked = false;
				chkScaleWidth.IsChecked = false;
				chkSharpness.IsChecked = false;
				chkBrightness.IsChecked = false;
				chkContrast.IsChecked = false;
				chkSaturation.IsChecked = false;
				chkTemperature.IsChecked = false;
				BB();
			}
			UpdateLayout();
			btnApply.IsEnabled = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private void A(Properties A)
	{
		List<string> list = new List<string>();
		if (A.Shape.HasTextFrame == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Properties.FontProperties font = A.Font;
					txtFontSize.Text = font.Size.ToString();
					System.Drawing.Color color = ColorTranslator.FromOle(font.ForeColor);
					this.A(rectFontColor, color);
					rectFontColor.Visibility = Visibility.Visible;
					txtFontColor.Text = this.A(color);
					txtFontName.Text = font.Name;
					if (font.Decoration.Bold == MsoTriState.msoTrue)
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
						list.Add(AH.A(50957));
					}
					if (font.Decoration.Italic == MsoTriState.msoTrue)
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
						list.Add(AH.A(50966));
					}
					if (font.Decoration.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
						list.Add(AH.A(50986));
					}
					font = default(Properties.FontProperties);
					if (!list.Any())
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
						txtDecoration.Text = AH.A(52691);
					}
					else
					{
						txtDecoration.Text = string.Join(AH.A(17773), list.ToArray());
					}
					list = null;
					return;
				}
				}
			}
		}
		rectFontColor.Visibility = Visibility.Collapsed;
		txtFontColor.Text = AH.A(17319);
		txtFontSize.Text = AH.A(17319);
		txtFontName.Text = AH.A(17319);
		txtDecoration.Text = AH.A(17319);
	}

	private void A(bool A)
	{
		chkFont.IsEnabled = A;
		chkFontColor.IsEnabled = A;
		chkFontName.IsEnabled = A;
		chkFontSize.IsEnabled = A;
		chkDecoration.IsEnabled = A;
	}

	private void B(Properties A)
	{
		Properties.FillProperties fill = A.Fill;
		if (fill.Visible == MsoTriState.msoTrue)
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
			if (fill.Type != MsoFillType.msoFillMixed)
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
				System.Drawing.Color color = ColorTranslator.FromOle(fill.ForeColor);
				this.A(rectFillColor, color);
				rectFillColor.Visibility = Visibility.Visible;
				txtFillColor.Text = this.A(color);
				switch (fill.Type)
				{
				case MsoFillType.msoFillSolid:
					txtFillType.Text = AH.A(138887);
					break;
				case MsoFillType.msoFillPatterned:
					txtFillType.Text = AH.A(138898);
					break;
				case MsoFillType.msoFillMixed:
					txtFillType.Text = AH.A(17319);
					break;
				case MsoFillType.msoFillTextured:
					txtFillType.Text = AH.A(138913);
					break;
				default:
					txtFillType.Text = fill.Type.ToString().Replace(AH.A(138928), "");
					break;
				}
				txtFillTransparency.Text = fill.Transparency.ToString(AH.A(14595));
				goto IL_01c6;
			}
		}
		rectFillColor.Visibility = Visibility.Collapsed;
		txtFillColor.Text = AH.A(17319);
		txtFillType.Text = AH.A(17319);
		txtFillTransparency.Text = AH.A(17319);
		goto IL_01c6;
		IL_01c6:
		fill = default(Properties.FillProperties);
	}

	private void A(Properties A, bool B)
	{
		Properties.LineProperties line = A.Line;
		if (line.Visible == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					System.Drawing.Color color = ColorTranslator.FromOle(line.ForeColor);
					this.A(rectLineColor, color);
					rectLineColor.Visibility = Visibility.Visible;
					txtLineColor.Text = this.A(color);
					string text;
					if (B)
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
						text = this.m_A;
					}
					else
					{
						text = this.m_D;
					}
					txtLineWeight.Text = line.Weight.ToString(text) + AH.A(138943);
					string text2;
					if (line.Style == MsoLineStyle.msoLineSingle)
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
						text2 = AH.A(50858);
					}
					else if (line.Style != MsoLineStyle.msoLineStyleMixed)
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
						text2 = line.Style.ToString().Replace(AH.A(138950), "");
						text2 = this.A(text2);
					}
					else
					{
						text2 = AH.A(138965);
					}
					string text3;
					if (line.DashStyle == MsoLineDashStyle.msoLineSolid)
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
						text3 = AH.A(138887);
					}
					else if (line.DashStyle != MsoLineDashStyle.msoLineDashStyleMixed)
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
						text3 = line.DashStyle.ToString().Replace(AH.A(138950), "");
						text3 = this.A(text3);
					}
					else
					{
						text3 = AH.A(138965);
					}
					txtLineStyle.Text = text2 + AH.A(17773) + text3;
					return;
				}
				}
			}
		}
		rectLineColor.Visibility = Visibility.Collapsed;
		txtLineColor.Text = AH.A(17319);
		txtLineWeight.Text = AH.A(17319);
		txtLineStyle.Text = AH.A(17319);
	}

	private void B(bool A)
	{
		chkLine.IsEnabled = A;
		chkLineColor.IsEnabled = A;
		chkLineStyle.IsEnabled = A;
		chkLineWeight.IsEnabled = A;
	}

	private void C(Properties A, bool B, bool C)
	{
		if (A.Shape.HasTextFrame == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					string text;
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
						text = AH.A(50115);
					}
					else
					{
						text = AH.A(50108);
					}
					string text2 = ((!B) ? this.m_D : this.m_A);
					Properties.TextBoxProperties textBox = A.TextBox;
					string text3 = "";
					if (textBox.Bullets.Count > 0)
					{
						chkBullets.IsEnabled = true;
						switch (textBox.Bullets.ElementAt(0).Value.Type)
						{
						case MsoBulletType.msoBulletNone:
							text3 = AH.A(52691);
							break;
						case MsoBulletType.msoBulletNumbered:
							text3 = AH.A(138976);
							break;
						case MsoBulletType.msoBulletUnnumbered:
							text3 = AH.A(138993);
							break;
						case MsoBulletType.msoBulletPicture:
							text3 = AH.A(3293);
							chkBullets.IsEnabled = false;
							chkBullets.IsChecked = false;
							break;
						default:
							text3 = AH.A(139014);
							break;
						}
					}
					else
					{
						text3 = AH.A(52691);
					}
					txtBullets.Text = text3;
					List<string> list = new List<string>();
					using (Dictionary<int, Properties.IndentProperties>.Enumerator enumerator = textBox.Indents.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							Properties.IndentProperties value = enumerator.Current.Value;
							if (C)
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
								list.Add(clsPublish.PointsToCentimeters(value.LeftIndent).ToString(text2) + AH.A(14622) + clsPublish.PointsToCentimeters(value.FirstLineIndent).ToString(text2));
							}
							else
							{
								list.Add(clsPublish.PointsToInches(value.LeftIndent).ToString(text2) + AH.A(14622) + clsPublish.PointsToInches(value.FirstLineIndent).ToString(text2));
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0228;
							}
							continue;
							end_IL_0228:
							break;
						}
					}
					if (list.Any())
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
						TextBlock textBlock = txtIndents;
						string text4 = string.Join(AH.A(14625), list.ToArray());
						string text5;
						if (!C)
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
							text5 = AH.A(50108);
						}
						else
						{
							text5 = AH.A(50115);
						}
						textBlock.Text = text4 + text5;
					}
					else
					{
						txtIndents.Text = AH.A(17319);
					}
					list = null;
					Properties.SpacingProperties lineSpacing = textBox.LineSpacing;
					if (lineSpacing.LineRuleWithin == MsoTriState.msoTrue)
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
						txtLineSpacing.Text = lineSpacing.SpaceBefore.ToString(text2) + AH.A(139019) + lineSpacing.SpaceAfter.ToString(text2) + AH.A(139019) + lineSpacing.SpaceWithin.ToString(text2) + AH.A(139032);
					}
					else
					{
						txtLineSpacing.Text = lineSpacing.SpaceBefore.ToString(text2) + AH.A(14622) + lineSpacing.SpaceAfter.ToString(text2) + AH.A(14622) + lineSpacing.SpaceWithin.ToString(text2) + AH.A(138943);
					}
					string text6 = textBox.HorizontalAlignment.ToString();
					text6 = text6.Replace(AH.A(139043), "");
					text6 = this.A(text6);
					txtAlignH.Text = text6;
					List<string> list2 = new List<string>();
					if (C)
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
						list2.Add(clsPublish.PointsToCentimeters(textBox.MarginLeft).ToString(text2));
						list2.Add(clsPublish.PointsToCentimeters(textBox.MarginRight).ToString(text2));
						list2.Add(clsPublish.PointsToCentimeters(textBox.MarginTop).ToString(text2));
						list2.Add(clsPublish.PointsToCentimeters(textBox.MarginBottom).ToString(text2));
					}
					else
					{
						list2.Add(clsPublish.PointsToInches(textBox.MarginLeft).ToString(text2));
						list2.Add(clsPublish.PointsToInches(textBox.MarginRight).ToString(text2));
						list2.Add(clsPublish.PointsToInches(textBox.MarginTop).ToString(text2));
						list2.Add(clsPublish.PointsToInches(textBox.MarginBottom).ToString(text2));
					}
					txtMargins.Text = string.Join(AH.A(14622), list2.ToArray()) + text;
					list2 = null;
					string text7 = "";
					switch (textBox.AutoSize)
					{
					case MsoAutoSize.msoAutoSizeNone:
						text7 = AH.A(139060);
						txtAutoSize.ToolTip = AH.A(32897);
						break;
					case MsoAutoSize.msoAutoSizeShapeToFitText:
						text7 = AH.A(139065);
						txtAutoSize.ToolTip = AH.A(32926);
						break;
					case MsoAutoSize.msoAutoSizeTextToFitShape:
						text7 = AH.A(139090);
						txtAutoSize.ToolTip = AH.A(32388);
						break;
					}
					txtAutoSize.Text = text7;
					string text8 = AH.A(17319);
					MsoTriState wordWrap = textBox.WordWrap;
					if (wordWrap != MsoTriState.msoTrue)
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
						if (wordWrap == MsoTriState.msoFalse)
						{
							text8 = AH.A(139060);
						}
					}
					else
					{
						text8 = AH.A(139113);
					}
					txtWordWrap.Text = text8;
					string text9 = AH.A(17319);
					switch (textBox.VerticalAnchor)
					{
					case MsoVerticalAnchor.msoAnchorTop:
						text9 = AH.A(120462);
						break;
					case MsoVerticalAnchor.msoAnchorMiddle:
						text9 = AH.A(139120);
						break;
					case MsoVerticalAnchor.msoAnchorBottom:
						text9 = AH.A(120469);
						break;
					}
					txtAlignV.Text = text9;
					string text10 = "";
					switch (textBox.Orientation)
					{
					case MsoTextOrientation.msoTextOrientationHorizontal:
						text10 = AH.A(139133);
						break;
					case MsoTextOrientation.msoTextOrientationDownward:
						text10 = AH.A(139154);
						break;
					case MsoTextOrientation.msoTextOrientationUpward:
						text10 = AH.A(139183);
						break;
					case MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast:
						text10 = AH.A(139214);
						break;
					case MsoTextOrientation.msoTextOrientationVertical:
						text10 = AH.A(139229);
						break;
					case MsoTextOrientation.msoTextOrientationVerticalFarEast:
						text10 = AH.A(139246);
						break;
					}
					txtOrientation.Text = text10;
					textBox = default(Properties.TextBoxProperties);
					return;
				}
				}
			}
		}
		txtBullets.Text = AH.A(17319);
		txtIndents.Text = AH.A(17319);
		txtLineSpacing.Text = AH.A(17319);
		txtAlignH.Text = AH.A(17319);
		txtMargins.Text = AH.A(17319);
		txtAutoSize.Text = AH.A(17319);
		txtWordWrap.Text = AH.A(17319);
		txtAlignV.Text = AH.A(17319);
		txtOrientation.Text = AH.A(17319);
	}

	private void C(bool A)
	{
		chkTextBox.IsEnabled = A;
		chkBullets.IsEnabled = A;
		chkIndents.IsEnabled = A;
		chkLineSpacing.IsEnabled = A;
		chkMargins.IsEnabled = A;
		chkAutoSize.IsEnabled = A;
		chkWordWrap.IsEnabled = A;
		chkAlignH.IsEnabled = A;
		chkAlignV.IsEnabled = A;
		chkOrientation.IsEnabled = A;
	}

	private void D(Properties A, bool B, bool C)
	{
		Properties.LayoutProperties layout = A.Layout;
		if (C)
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
			string text = (B ? this.m_C : this.m_F);
			txtHeight.Text = clsPublish.PointsToCentimeters(layout.Height).ToString(text);
			txtWidth.Text = clsPublish.PointsToCentimeters(layout.Width).ToString(text);
			txtLeft.Text = clsPublish.PointsToCentimeters(layout.Left).ToString(text);
			txtRight.Text = clsPublish.PointsToCentimeters(layout.Right).ToString(text);
			txtMidpointX.Text = clsPublish.PointsToCentimeters(layout.MidpointX).ToString(text);
			txtTop.Text = clsPublish.PointsToCentimeters(layout.Top).ToString(text);
			txtBottom.Text = clsPublish.PointsToCentimeters(layout.Bottom).ToString(text);
			txtMidpointY.Text = clsPublish.PointsToCentimeters(layout.MidpointY).ToString(text);
		}
		else
		{
			string text = (B ? this.m_B : this.m_E);
			txtHeight.Text = clsPublish.PointsToInches(layout.Height).ToString(text);
			txtWidth.Text = clsPublish.PointsToInches(layout.Width).ToString(text);
			txtLeft.Text = clsPublish.PointsToInches(layout.Left).ToString(text);
			txtRight.Text = clsPublish.PointsToInches(layout.Right).ToString(text);
			txtMidpointX.Text = clsPublish.PointsToInches(layout.MidpointX).ToString(text);
			txtTop.Text = clsPublish.PointsToInches(layout.Top).ToString(text);
			txtBottom.Text = clsPublish.PointsToInches(layout.Bottom).ToString(text);
			txtMidpointY.Text = clsPublish.PointsToInches(layout.MidpointY).ToString(text);
		}
		if (layout.LockAspectRatio == MsoTriState.msoTrue)
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
			txtLockAspectRatio.Text = AH.A(139113);
		}
		else
		{
			txtLockAspectRatio.Text = AH.A(139060);
		}
		txtRotation.Text = layout.Rotation.ToString(AH.A(139285)) + AH.A(139288);
	}

	private void B(Properties A, bool B)
	{
		Properties.ShapeEffectsProperties shapeEffects = A.ShapeEffects;
		this.A(txtShapeGlow, rectShapeGlow, shapeEffects.Glow, B);
		txtShapeReflection.Text = this.A(shapeEffects.Reflection);
		if (shapeEffects.Shadow.Visible == MsoTriState.msoFalse)
		{
			rectShapeShadow.Visibility = Visibility.Collapsed;
			txtShapeShadow.Text = AH.A(17319);
		}
		else
		{
			System.Drawing.Color color = ColorTranslator.FromOle(shapeEffects.Shadow.ForeColor);
			this.A(rectShapeShadow, color);
			rectShapeShadow.Visibility = Visibility.Visible;
			txtShapeShadow.Text = this.A(color);
		}
		txtShape3D.Text = this.A(shapeEffects.ThreeD);
		if (shapeEffects.SoftEdge.Type == MsoSoftEdgeType.msoSoftEdgeTypeNone)
		{
			txtShapeSoftEdge.Text = AH.A(17319);
		}
		else
		{
			txtShapeSoftEdge.Text = this.A(shapeEffects.SoftEdge.Type.ToString().Replace(AH.A(139297), ""));
		}
		shapeEffects = default(Properties.ShapeEffectsProperties);
	}

	private void C(Properties A, bool B)
	{
		Properties.TextEffectsProperties textEffects = A.TextEffects;
		this.A(txtTextGlow, rectTextGlow, textEffects.Glow, B);
		txtTextReflection.Text = this.A(textEffects.Reflection);
		if (textEffects.Shadow.Visible == MsoTriState.msoFalse)
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
			rectTextShadow.Visibility = Visibility.Collapsed;
			txtTextShadow.Text = AH.A(17319);
		}
		else
		{
			System.Drawing.Color color = ColorTranslator.FromOle(textEffects.Shadow.ForeColor);
			this.A(rectTextShadow, color);
			rectTextShadow.Visibility = Visibility.Visible;
			txtTextShadow.Text = this.A(color);
		}
		txtText3D.Text = this.A(textEffects.ThreeD);
		if (textEffects.SoftEdge == MsoSoftEdgeType.msoSoftEdgeTypeNone)
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
			txtTextSoftEdge.Text = AH.A(17319);
		}
		else
		{
			txtTextSoftEdge.Text = this.A(textEffects.SoftEdge.ToString().Replace(AH.A(139297), ""));
		}
		textEffects = default(Properties.TextEffectsProperties);
	}

	private void D(Properties A, bool B)
	{
		string text = ((!B) ? this.m_D : this.m_A);
		if (A.Shape.Type == MsoShapeType.msoAutoShape)
		{
			List<string> list = new List<string>();
			Properties.AutoShapeProperties autoShape = A.AutoShape;
			string a = autoShape.Type.ToString().Replace(AH.A(139320), "");
			a = this.A(a);
			a = Regex.Replace(a, AH.A(139337), AH.A(139368));
			txtAutoShapeType.Text = a;
			if (autoShape.Adjustments != null)
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
				foreach (float adjustment in autoShape.Adjustments)
				{
					list.Add(adjustment.ToString(text));
				}
			}
			autoShape = default(Properties.AutoShapeProperties);
			if (list.Any())
			{
				txtAdjustments.Text = string.Join(AH.A(17773), list.ToArray());
			}
			else
			{
				txtAdjustments.Text = AH.A(17319);
			}
			list = null;
		}
		else
		{
			txtAutoShapeType.Text = AH.A(17319);
			txtAdjustments.Text = AH.A(17319);
		}
	}

	private void D(bool A)
	{
		chkAutoShape.IsEnabled = A;
		chkAutoShapeType.IsEnabled = A;
		chkAdjustments.IsEnabled = A;
	}

	private void E(Properties A, bool B, bool C)
	{
	}

	private void F(Properties A, bool B, bool C)
	{
		string text = ((!B) ? this.m_D : this.m_A);
		if (A.Shape.HasPicture)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Properties.PictureProperties picture = A.Picture;
					txtScaleHeight.Text = picture.ScaleHeight.ToString(AH.A(14595));
					txtScaleWidth.Text = picture.ScaleWidth.ToString(AH.A(14595));
					try
					{
						TextBlock textBlock = txtSharpness;
						Dictionary<MsoPictureEffectType, List<float>> pictureEffects = picture.PictureEffects;
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> predicate;
						if (_Closure_0024__.A == null)
						{
							predicate = (_Closure_0024__.A = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Key == MsoPictureEffectType.msoEffectSharpenSoften);
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
							predicate = _Closure_0024__.A;
						}
						IEnumerable<KeyValuePair<MsoPictureEffectType, List<float>>> source = pictureEffects.Where(predicate);
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> selector;
						if (_Closure_0024__.A == null)
						{
							selector = (_Closure_0024__.A = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Value);
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
							selector = _Closure_0024__.A;
						}
						textBlock.Text = source.Select(selector).ToList()[0][0].ToString(AH.A(14595));
						chkSharpness.IsEnabled = true;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						txtSharpness.Text = AH.A(17319);
						chkSharpness.IsChecked = false;
						chkSharpness.IsEnabled = false;
						ProjectData.ClearProjectError();
					}
					try
					{
						TextBlock textBlock2 = txtBrightness;
						IEnumerable<KeyValuePair<MsoPictureEffectType, List<float>>> source2 = picture.PictureEffects.Where([SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Key == MsoPictureEffectType.msoEffectBrightnessContrast);
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> selector2;
						if (_Closure_0024__.B == null)
						{
							selector2 = (_Closure_0024__.B = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Value);
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
							selector2 = _Closure_0024__.B;
						}
						textBlock2.Text = source2.Select(selector2).ToList()[0][0].ToString(AH.A(14595));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						txtBrightness.Text = (2f * picture.Brightness - 1f).ToString(AH.A(14595));
						ProjectData.ClearProjectError();
					}
					try
					{
						TextBlock textBlock3 = txtContrast;
						Dictionary<MsoPictureEffectType, List<float>> pictureEffects2 = picture.PictureEffects;
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> predicate2;
						if (_Closure_0024__.C == null)
						{
							predicate2 = (_Closure_0024__.C = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Key == MsoPictureEffectType.msoEffectBrightnessContrast);
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
							predicate2 = _Closure_0024__.C;
						}
						IEnumerable<KeyValuePair<MsoPictureEffectType, List<float>>> source3 = pictureEffects2.Where(predicate2);
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> selector3;
						if (_Closure_0024__.C == null)
						{
							selector3 = (_Closure_0024__.C = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Value);
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
							selector3 = _Closure_0024__.C;
						}
						textBlock3.Text = source3.Select(selector3).ToList()[0][1].ToString(AH.A(14595));
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						txtContrast.Text = (2f * picture.Contrast - 1f).ToString(AH.A(14595));
						ProjectData.ClearProjectError();
					}
					try
					{
						TextBlock textBlock4 = txtSaturation;
						Dictionary<MsoPictureEffectType, List<float>> pictureEffects3 = picture.PictureEffects;
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> predicate3;
						if (_Closure_0024__.D == null)
						{
							predicate3 = (_Closure_0024__.D = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Key == MsoPictureEffectType.msoEffectSaturation);
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
							predicate3 = _Closure_0024__.D;
						}
						IEnumerable<KeyValuePair<MsoPictureEffectType, List<float>>> source4 = pictureEffects3.Where(predicate3);
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> selector4;
						if (_Closure_0024__.D == null)
						{
							selector4 = (_Closure_0024__.D = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Value);
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
							selector4 = _Closure_0024__.D;
						}
						textBlock4.Text = source4.Select(selector4).ToList()[0][0].ToString(AH.A(14595));
						chkSaturation.IsEnabled = true;
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						txtSaturation.Text = AH.A(17319);
						chkSaturation.IsChecked = false;
						chkSaturation.IsEnabled = false;
						ProjectData.ClearProjectError();
					}
					try
					{
						TextBlock textBlock5 = txtTemperature;
						Dictionary<MsoPictureEffectType, List<float>> pictureEffects4 = picture.PictureEffects;
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, bool> predicate4;
						if (_Closure_0024__.E == null)
						{
							predicate4 = (_Closure_0024__.E = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Key == MsoPictureEffectType.msoEffectColorTemperature);
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
							predicate4 = _Closure_0024__.E;
						}
						IEnumerable<KeyValuePair<MsoPictureEffectType, List<float>>> source5 = pictureEffects4.Where(predicate4);
						Func<KeyValuePair<MsoPictureEffectType, List<float>>, List<float>> selector5;
						if (_Closure_0024__.E == null)
						{
							selector5 = (_Closure_0024__.E = [SpecialName] (KeyValuePair<MsoPictureEffectType, List<float>> keyValuePair) => keyValuePair.Value);
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
							selector5 = _Closure_0024__.E;
						}
						textBlock5.Text = source5.Select(selector5).ToList()[0][0].ToString(text);
						chkTemperature.IsEnabled = true;
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						txtTemperature.Text = AH.A(17319);
						chkTemperature.IsChecked = false;
						chkTemperature.IsEnabled = false;
						ProjectData.ClearProjectError();
					}
					picture = default(Properties.PictureProperties);
					return;
				}
				}
			}
		}
		txtScaleHeight.Text = AH.A(17319);
		txtScaleWidth.Text = AH.A(17319);
		txtSharpness.Text = AH.A(17319);
		txtBrightness.Text = AH.A(17319);
		txtContrast.Text = AH.A(17319);
		txtSaturation.Text = AH.A(17319);
		txtTemperature.Text = AH.A(17319);
	}

	private void E(bool A)
	{
		chkPicture.IsEnabled = A;
		chkPictureScale.IsEnabled = A;
		chkScaleHeight.IsEnabled = A;
		chkScaleWidth.IsEnabled = A;
		chkSharpness.IsEnabled = A;
		chkBrightness.IsEnabled = A;
		chkContrast.IsEnabled = A;
		chkSaturation.IsEnabled = A;
		chkTemperature.IsEnabled = A;
	}

	private string A(string A)
	{
		return Regex.Replace(A, AH.A(139385), AH.A(139400), RegexOptions.Compiled).Trim();
	}

	private string A(Properties.ThreeDProperties A)
	{
		if (A.Visible == MsoTriState.msoFalse)
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
					return AH.A(17319);
				}
			}
		}
		return AH.A(139407);
	}

	private string A(Properties.ReflectionProperties A)
	{
		if (A.Type != MsoReflectionType.msoReflectionTypeNone)
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
			if (A.Type != MsoReflectionType.msoReflectionTypeMixed)
			{
				return this.A(A.Type.ToString().Replace(AH.A(139412), ""));
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
		return AH.A(17319);
	}

	private void A(TextBlock A, System.Windows.Shapes.Rectangle B, Properties.GlowProperties C, bool D)
	{
		if (C.Radius == 0f)
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
					B.Visibility = Visibility.Collapsed;
					A.Text = AH.A(17319);
					return;
				}
			}
		}
		if (D)
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
			_ = this.m_A;
		}
		else
		{
			_ = this.m_D;
		}
		System.Drawing.Color color = ColorTranslator.FromOle(C.Color);
		this.A(B, color);
		B.Visibility = Visibility.Visible;
		A.Text = this.A(color);
	}

	private void A(System.Windows.Shapes.Rectangle A, System.Drawing.Color B)
	{
		System.Windows.Shapes.Rectangle rectangle = A;
		rectangle.Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(B.R, B.G, B.B));
		if (this.A(B) < 180)
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
		return clsColors.Color2RGB(A).Replace(AH.A(12717), AH.A(14258));
	}

	private void B()
	{
		chkLineColor.Checked += LineChildCheckChanged;
		chkLineColor.Unchecked += LineChildCheckChanged;
		chkLineWeight.Checked += LineChildCheckChanged;
		chkLineWeight.Unchecked += LineChildCheckChanged;
		chkLineStyle.Checked += LineChildCheckChanged;
		chkLineStyle.Unchecked += LineChildCheckChanged;
	}

	private void C()
	{
		chkLineColor.Checked -= LineChildCheckChanged;
		chkLineColor.Unchecked -= LineChildCheckChanged;
		chkLineWeight.Checked -= LineChildCheckChanged;
		chkLineWeight.Unchecked -= LineChildCheckChanged;
		chkLineStyle.Checked -= LineChildCheckChanged;
		chkLineStyle.Unchecked -= LineChildCheckChanged;
	}

	private void RemoveLineCheckBoxHandlers(object sender, RoutedEventArgs e)
	{
		bool value = chkLine.IsChecked.Value;
		C();
		chkLineColor.IsChecked = value;
		chkLineWeight.IsChecked = value;
		chkLineStyle.IsChecked = value;
		B();
	}

	private void LineChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkLineColor.IsChecked);
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
		bool? isChecked2;
		flag = (isChecked2 = chkLineWeight.IsChecked);
		if (!flag.HasValue)
		{
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0089;
		IL_01cf:
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
				goto IL_025c;
			}
		}
		isChecked = chkLineStyle.IsChecked;
		if (((!isChecked) ?? isChecked) == true)
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
			if (flag2.HasValue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						chkLine.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_025c;
		IL_00f6:
		flag = chkLineColor.IsChecked;
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
		isChecked2 = flag3;
		flag = flag3;
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
				goto IL_01cf;
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
		flag = chkLineWeight.IsChecked;
		bool? flag4;
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
			obj2 = null;
		}
		else if (isChecked != true)
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
			obj2 = isChecked2;
		}
		goto IL_01cf;
		IL_0089:
		flag2 = obj;
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
				goto IL_00f6;
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
		if (chkLineStyle.IsChecked == true && flag2.HasValue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkLine.IsChecked = true;
					return;
				}
			}
		}
		goto IL_00f6;
		IL_025c:
		chkLine.IsChecked = null;
	}

	private void D()
	{
		chkFillColor.Checked += FillChildCheckChanged;
		chkFillColor.Unchecked += FillChildCheckChanged;
		chkFillType.Checked += FillChildCheckChanged;
		chkFillType.Unchecked += FillChildCheckChanged;
		chkFillTransparency.Checked += FillChildCheckChanged;
		chkFillTransparency.Unchecked += FillChildCheckChanged;
	}

	private void E()
	{
		chkFillColor.Checked -= FillChildCheckChanged;
		chkFillColor.Unchecked -= FillChildCheckChanged;
		chkFillType.Checked -= FillChildCheckChanged;
		chkFillType.Unchecked -= FillChildCheckChanged;
		chkFillTransparency.Checked -= FillChildCheckChanged;
		chkFillTransparency.Unchecked -= FillChildCheckChanged;
	}

	private void RemoveFillCheckBoxHandlers(object sender, RoutedEventArgs e)
	{
		bool value = chkFill.IsChecked.Value;
		E();
		chkFillColor.IsChecked = value;
		chkFillType.IsChecked = value;
		chkFillTransparency.IsChecked = value;
		D();
	}

	private void FillChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkFillColor.IsChecked);
		bool? obj;
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_009b;
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
		bool? isChecked2;
		flag = (isChecked2 = chkFillType.IsChecked);
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_009b;
		IL_01e1:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (flag2.HasValue)
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
			if (flag2 != true)
			{
				goto IL_0279;
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
		isChecked = chkFillTransparency.IsChecked;
		bool? flag3;
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
			flag3 = isChecked;
		}
		else
		{
			flag3 = isChecked != true;
		}
		isChecked = flag3;
		if (isChecked == true && flag2.HasValue)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkFill.IsChecked = false;
					return;
				}
			}
		}
		goto IL_0279;
		IL_0102:
		flag = chkFillColor.IsChecked;
		bool? flag4;
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
			flag4 = flag;
		}
		else
		{
			flag4 = flag != true;
		}
		isChecked2 = flag4;
		flag = flag4;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj2 = false;
				goto IL_01e1;
			}
		}
		flag = chkFillType.IsChecked;
		bool? flag5;
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
			flag5 = flag;
		}
		else
		{
			flag5 = flag != true;
		}
		isChecked = flag5;
		flag = flag5;
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
			obj2 = null;
		}
		else if (isChecked != true)
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
			obj2 = false;
		}
		else
		{
			obj2 = isChecked2;
		}
		goto IL_01e1;
		IL_0279:
		chkFill.IsChecked = null;
		return;
		IL_009b:
		flag2 = obj;
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
				goto IL_0102;
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
		if (chkFillTransparency.IsChecked == true)
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
			if (flag2.HasValue)
			{
				chkFill.IsChecked = true;
				return;
			}
		}
		goto IL_0102;
	}

	private void F()
	{
		chkFontSize.Checked += FontChildCheckChanged;
		chkFontSize.Unchecked += FontChildCheckChanged;
		chkFontColor.Checked += FontChildCheckChanged;
		chkFontColor.Unchecked += FontChildCheckChanged;
		chkFontName.Checked += FontChildCheckChanged;
		chkFontName.Unchecked += FontChildCheckChanged;
		chkDecoration.Checked += FontChildCheckChanged;
		chkDecoration.Unchecked += FontChildCheckChanged;
	}

	private void G()
	{
		chkFontSize.Checked -= FontChildCheckChanged;
		chkFontSize.Unchecked -= FontChildCheckChanged;
		chkFontColor.Checked -= FontChildCheckChanged;
		chkFontColor.Unchecked -= FontChildCheckChanged;
		chkFontName.Checked -= FontChildCheckChanged;
		chkFontName.Unchecked -= FontChildCheckChanged;
		chkDecoration.Checked -= FontChildCheckChanged;
		chkDecoration.Unchecked -= FontChildCheckChanged;
	}

	private void RemoveFontCheckBoxHandlers(object sender, RoutedEventArgs e)
	{
		bool value = chkFont.IsChecked.Value;
		G();
		chkFontSize.IsChecked = value;
		chkFontColor.IsChecked = value;
		chkFontName.IsChecked = value;
		chkDecoration.IsChecked = value;
		F();
	}

	private void FontChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkFontSize.IsChecked);
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
				goto IL_0095;
			}
		}
		bool? isChecked2;
		flag = (isChecked2 = chkFontColor.IsChecked);
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0095;
		IL_03bc:
		chkFont.IsChecked = null;
		return;
		IL_0188:
		flag = chkFontSize.IsChecked;
		bool? flag2;
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
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			if (isChecked2 != true)
			{
				obj2 = false;
				goto IL_0272;
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
		flag = chkFontColor.IsChecked;
		bool? flag3;
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
			while (true)
			{
				switch (1)
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
				switch (1)
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
		goto IL_0272;
		IL_0095:
		bool? flag4 = obj;
		isChecked2 = obj;
		bool? obj3;
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
				obj3 = false;
				goto IL_0119;
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
		bool? isChecked3;
		isChecked2 = (isChecked3 = chkFontName.IsChecked);
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
			obj3 = null;
		}
		else if (isChecked3 != true)
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
			obj3 = flag4;
		}
		goto IL_0119;
		IL_0119:
		bool? flag5 = obj3;
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
				goto IL_0188;
			}
		}
		if (chkDecoration.IsChecked == true)
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
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						chkFont.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_0188;
		IL_0272:
		isChecked3 = obj2;
		isChecked = obj2;
		bool? obj4;
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
			if (isChecked3 != true)
			{
				obj4 = false;
				goto IL_0322;
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
		isChecked = chkFontName.IsChecked;
		bool? flag6;
		if (!isChecked.HasValue)
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
			flag6 = isChecked;
		}
		else
		{
			flag6 = isChecked != true;
		}
		flag4 = flag6;
		isChecked = flag6;
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
			obj4 = null;
		}
		else if (flag4 != true)
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
			obj4 = isChecked3;
		}
		goto IL_0322;
		IL_0322:
		flag5 = obj4;
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
				goto IL_03bc;
			}
		}
		flag4 = chkDecoration.IsChecked;
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
				switch (1)
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
					switch (2)
					{
					case 0:
						break;
					default:
						chkFont.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_03bc;
	}

	private void H()
	{
		chkBullets.Checked += TextBoxChildCheckChanged;
		chkBullets.Unchecked += TextBoxChildCheckChanged;
		chkIndents.Checked += TextBoxChildCheckChanged;
		chkIndents.Unchecked += TextBoxChildCheckChanged;
		chkLineSpacing.Checked += TextBoxChildCheckChanged;
		chkLineSpacing.Unchecked += TextBoxChildCheckChanged;
		chkMargins.Checked += TextBoxChildCheckChanged;
		chkMargins.Unchecked += TextBoxChildCheckChanged;
		chkAutoSize.Checked += TextBoxChildCheckChanged;
		chkAutoSize.Unchecked += TextBoxChildCheckChanged;
		chkWordWrap.Checked += TextBoxChildCheckChanged;
		chkWordWrap.Unchecked += TextBoxChildCheckChanged;
		chkAlignH.Checked += TextBoxChildCheckChanged;
		chkAlignH.Unchecked += TextBoxChildCheckChanged;
		chkAlignV.Checked += TextBoxChildCheckChanged;
		chkAlignV.Unchecked += TextBoxChildCheckChanged;
		chkOrientation.Checked += TextBoxChildCheckChanged;
		chkOrientation.Unchecked += TextBoxChildCheckChanged;
	}

	private void I()
	{
		chkBullets.Checked -= TextBoxChildCheckChanged;
		chkBullets.Unchecked -= TextBoxChildCheckChanged;
		chkIndents.Checked -= TextBoxChildCheckChanged;
		chkIndents.Unchecked -= TextBoxChildCheckChanged;
		chkLineSpacing.Checked -= TextBoxChildCheckChanged;
		chkLineSpacing.Unchecked -= TextBoxChildCheckChanged;
		chkMargins.Checked -= TextBoxChildCheckChanged;
		chkMargins.Unchecked -= TextBoxChildCheckChanged;
		chkAutoSize.Checked -= TextBoxChildCheckChanged;
		chkAutoSize.Unchecked -= TextBoxChildCheckChanged;
		chkWordWrap.Checked -= TextBoxChildCheckChanged;
		chkWordWrap.Unchecked -= TextBoxChildCheckChanged;
		chkAlignH.Checked -= TextBoxChildCheckChanged;
		chkAlignH.Unchecked -= TextBoxChildCheckChanged;
		chkAlignV.Checked -= TextBoxChildCheckChanged;
		chkAlignV.Unchecked -= TextBoxChildCheckChanged;
		chkOrientation.Checked -= TextBoxChildCheckChanged;
		chkOrientation.Unchecked -= TextBoxChildCheckChanged;
	}

	private void TextBoxCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkTextBox.IsChecked.Value;
		I();
		chkBullets.IsChecked = value;
		chkIndents.IsChecked = value;
		chkLineSpacing.IsChecked = value;
		chkMargins.IsChecked = value;
		chkAutoSize.IsChecked = value;
		chkWordWrap.IsChecked = value;
		chkAlignH.IsChecked = value;
		chkAlignV.IsChecked = value;
		chkOrientation.IsChecked = value;
		H();
	}

	private void TextBoxChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkBullets.IsChecked);
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
				goto IL_009d;
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
		bool? isChecked2;
		flag = (isChecked2 = chkIndents.IsChecked);
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_009d;
		IL_04e2:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		isChecked = (bool?)obj2;
		bool? obj3;
		if (isChecked.HasValue)
		{
			if (flag2 != true)
			{
				obj3 = false;
				goto IL_058e;
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
		isChecked = chkLineSpacing.IsChecked;
		bool? flag3;
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
		else if (flag4 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag2;
		}
		goto IL_058e;
		IL_0396:
		bool? obj4;
		bool? flag5 = (bool?)obj4;
		if (flag5.HasValue)
		{
			if (flag5 != true)
			{
				goto IL_03ff;
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
		if (chkOrientation.IsChecked == true)
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
			if (flag5.HasValue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						chkTextBox.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_03ff;
		IL_058e:
		bool? flag6 = obj3;
		flag4 = obj3;
		bool? obj5;
		if (flag4.HasValue)
		{
			if (flag6 != true)
			{
				obj5 = false;
				goto IL_063c;
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
		flag4 = chkAlignH.IsChecked;
		bool? flag7;
		if (!flag4.HasValue)
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
			flag7 = flag4;
		}
		else
		{
			flag7 = flag4 != true;
		}
		bool? flag8 = flag7;
		flag4 = flag7;
		if (!flag4.HasValue)
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
			obj5 = null;
		}
		else if (flag8 != true)
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
			obj5 = false;
		}
		else
		{
			obj5 = flag6;
		}
		goto IL_063c;
		IL_06f2:
		bool? obj6;
		bool? flag9 = (bool?)obj6;
		bool? flag10 = (bool?)obj6;
		bool? obj7;
		if (flag10.HasValue)
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
			if (flag9 != true)
			{
				obj7 = false;
				goto IL_07a2;
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
		flag10 = chkAutoSize.IsChecked;
		bool? flag11;
		if (!flag10.HasValue)
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
			flag11 = flag10;
		}
		else
		{
			flag11 = flag10 != true;
		}
		bool? flag12 = flag11;
		flag10 = flag11;
		if (!flag10.HasValue)
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
			obj7 = null;
		}
		else if (flag12 != true)
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
			obj7 = false;
		}
		else
		{
			obj7 = flag9;
		}
		goto IL_07a2;
		IL_0855:
		bool? obj8;
		bool? flag13 = (bool?)obj8;
		bool? flag14 = (bool?)obj8;
		bool? obj9;
		if (flag14.HasValue)
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
			if (flag13 != true)
			{
				obj9 = false;
				goto IL_0903;
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
		flag14 = chkAlignV.IsChecked;
		bool? flag15;
		if (!flag14.HasValue)
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
			flag15 = flag14;
		}
		else
		{
			flag15 = flag14 != true;
		}
		bool? flag16 = flag15;
		flag14 = flag15;
		if (!flag14.HasValue)
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
			obj9 = null;
		}
		else if (flag16 != true)
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
			obj9 = flag13;
		}
		goto IL_0903;
		IL_09a3:
		chkTextBox.IsChecked = null;
		return;
		IL_021a:
		bool? obj10;
		flag12 = (bool?)obj10;
		bool? flag17 = (bool?)obj10;
		bool? obj11;
		if (flag17.HasValue)
		{
			if (flag12 != true)
			{
				obj11 = false;
				goto IL_0291;
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
		flag17 = (flag9 = chkAutoSize.IsChecked);
		if (!flag17.HasValue)
		{
			obj11 = null;
		}
		else if (flag9 != true)
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
			obj11 = false;
		}
		else
		{
			obj11 = flag12;
		}
		goto IL_0291;
		IL_019d:
		bool? obj12;
		flag10 = (bool?)obj12;
		flag6 = (bool?)obj12;
		if (flag6.HasValue)
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
			if (flag10 != true)
			{
				obj10 = false;
				goto IL_021a;
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
		flag6 = (flag17 = chkMargins.IsChecked);
		if (!flag6.HasValue)
		{
			obj10 = null;
		}
		else if (flag17 != true)
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
			obj10 = false;
		}
		else
		{
			obj10 = flag10;
		}
		goto IL_021a;
		IL_0903:
		flag5 = obj9;
		if (flag5.HasValue)
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
			if (flag5 != true)
			{
				goto IL_09a3;
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
		flag16 = chkOrientation.IsChecked;
		bool? flag18;
		if (!flag16.HasValue)
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
			flag18 = flag16;
		}
		else
		{
			flag18 = flag16 != true;
		}
		flag16 = flag18;
		if (flag16 == true)
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
					switch (3)
					{
					case 0:
						break;
					default:
						chkTextBox.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_09a3;
		IL_0312:
		bool? obj13;
		flag16 = (bool?)obj13;
		bool? flag19 = (bool?)obj13;
		if (flag19.HasValue)
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
				obj4 = false;
				goto IL_0396;
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
		flag19 = (flag13 = chkAlignV.IsChecked);
		if (!flag19.HasValue)
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
			obj4 = null;
		}
		else if (flag13 != true)
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
			obj4 = flag16;
		}
		goto IL_0396;
		IL_0291:
		flag14 = obj11;
		flag9 = obj11;
		if (flag9.HasValue)
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
				obj13 = false;
				goto IL_0312;
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
		flag9 = (flag19 = chkWordWrap.IsChecked);
		if (!flag9.HasValue)
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
			obj13 = null;
		}
		else if (flag19 != true)
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
			obj13 = false;
		}
		else
		{
			obj13 = flag14;
		}
		goto IL_0312;
		IL_07a2:
		flag19 = obj7;
		flag12 = obj7;
		if (flag12.HasValue)
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
			if (flag19 != true)
			{
				obj8 = false;
				goto IL_0855;
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
		flag12 = chkWordWrap.IsChecked;
		bool? flag20;
		if (!flag12.HasValue)
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
			flag20 = flag12;
		}
		else
		{
			flag20 = flag12 != true;
		}
		flag14 = flag20;
		flag12 = flag20;
		if (!flag12.HasValue)
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
			obj8 = null;
		}
		else if (flag14 != true)
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
			obj8 = false;
		}
		else
		{
			obj8 = flag19;
		}
		goto IL_0855;
		IL_011e:
		bool? obj14;
		flag8 = (bool?)obj14;
		flag2 = (bool?)obj14;
		if (flag2.HasValue)
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
			if (flag8 != true)
			{
				obj12 = false;
				goto IL_019d;
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
		flag2 = (flag6 = chkAlignH.IsChecked);
		if (!flag2.HasValue)
		{
			obj12 = null;
		}
		else if (flag6 != true)
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
			obj12 = false;
		}
		else
		{
			obj12 = flag8;
		}
		goto IL_019d;
		IL_009d:
		flag4 = obj;
		isChecked2 = obj;
		if (isChecked2.HasValue)
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
				obj14 = false;
				goto IL_011e;
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
		isChecked2 = (flag2 = chkLineSpacing.IsChecked);
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
			obj14 = null;
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
			obj14 = false;
		}
		else
		{
			obj14 = flag4;
		}
		goto IL_011e;
		IL_03ff:
		flag = chkBullets.IsChecked;
		bool? flag21;
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
			flag21 = flag;
		}
		else
		{
			flag21 = flag != true;
		}
		isChecked2 = flag21;
		flag = flag21;
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
				goto IL_04e2;
			}
		}
		flag = chkIndents.IsChecked;
		bool? flag22;
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
			flag22 = flag;
		}
		else
		{
			flag22 = flag != true;
		}
		isChecked = flag22;
		flag = flag22;
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
		goto IL_04e2;
		IL_063c:
		flag17 = obj5;
		flag8 = obj5;
		if (flag8.HasValue)
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
			if (flag17 != true)
			{
				obj6 = false;
				goto IL_06f2;
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
		flag8 = chkMargins.IsChecked;
		bool? flag23;
		if (!flag8.HasValue)
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
			flag23 = flag8;
		}
		else
		{
			flag23 = flag8 != true;
		}
		flag10 = flag23;
		flag8 = flag23;
		if (!flag8.HasValue)
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
			obj6 = null;
		}
		else if (flag10 != true)
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
			obj6 = false;
		}
		else
		{
			obj6 = flag17;
		}
		goto IL_06f2;
	}

	private void J()
	{
		chkPositionX.Checked += XPositionCheckChanged;
		chkPositionX.Unchecked += XPositionCheckChanged;
		radLeft.Checked += XPositionChildCheckChanged;
		radLeft.Unchecked += XPositionChildCheckChanged;
		radRight.Checked += XPositionChildCheckChanged;
		radRight.Unchecked += XPositionChildCheckChanged;
		radMidPointX.Checked += XPositionChildCheckChanged;
		radMidPointX.Unchecked += XPositionChildCheckChanged;
	}

	private void K()
	{
		chkPositionX.Checked -= XPositionCheckChanged;
		chkPositionX.Unchecked -= XPositionCheckChanged;
		radLeft.Checked -= XPositionChildCheckChanged;
		radLeft.Unchecked -= XPositionChildCheckChanged;
		radRight.Checked -= XPositionChildCheckChanged;
		radRight.Unchecked -= XPositionChildCheckChanged;
		radMidPointX.Checked -= XPositionChildCheckChanged;
		radMidPointX.Unchecked -= XPositionChildCheckChanged;
	}

	private void XPositionCheckChanged(object sender, RoutedEventArgs e)
	{
		K();
		if (chkPositionX.IsChecked == true)
		{
			radLeft.IsChecked = true;
		}
		else
		{
			radLeft.IsChecked = false;
			radRight.IsChecked = false;
			radMidPointX.IsChecked = false;
		}
		J();
	}

	private void XPositionChildCheckChanged(object sender, RoutedEventArgs e)
	{
		K();
		if (radLeft.IsChecked != true)
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
			if (radRight.IsChecked != true)
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
				if (radMidPointX.IsChecked != true)
				{
					chkPositionX.IsChecked = false;
					goto IL_00a0;
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
		}
		chkPositionX.IsChecked = true;
		goto IL_00a0;
		IL_00a0:
		J();
	}

	private void L()
	{
		chkPositionY.Checked += YPositionCheckChanged;
		chkPositionY.Unchecked += YPositionCheckChanged;
		radTop.Checked += YPositionChildCheckChanged;
		radTop.Unchecked += YPositionChildCheckChanged;
		radBottom.Checked += YPositionChildCheckChanged;
		radBottom.Unchecked += YPositionChildCheckChanged;
		radMidPointY.Checked += YPositionChildCheckChanged;
		radMidPointY.Unchecked += YPositionChildCheckChanged;
	}

	private void M()
	{
		chkPositionY.Checked -= YPositionCheckChanged;
		chkPositionY.Unchecked -= YPositionCheckChanged;
		radTop.Checked -= YPositionChildCheckChanged;
		radTop.Unchecked -= YPositionChildCheckChanged;
		radBottom.Checked -= YPositionChildCheckChanged;
		radBottom.Unchecked -= YPositionChildCheckChanged;
		radMidPointY.Checked -= YPositionChildCheckChanged;
		radMidPointY.Unchecked -= YPositionChildCheckChanged;
	}

	private void YPositionCheckChanged(object sender, RoutedEventArgs e)
	{
		M();
		if (chkPositionY.IsChecked == true)
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
			radTop.IsChecked = true;
		}
		else
		{
			radTop.IsChecked = false;
			radBottom.IsChecked = false;
			radMidPointY.IsChecked = false;
		}
		L();
	}

	private void YPositionChildCheckChanged(object sender, RoutedEventArgs e)
	{
		M();
		if (radTop.IsChecked != true)
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
			if (radBottom.IsChecked != true)
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
				if (radMidPointY.IsChecked != true)
				{
					chkPositionY.IsChecked = false;
					goto IL_009e;
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
		}
		chkPositionY.IsChecked = true;
		goto IL_009e;
		IL_009e:
		L();
	}

	private void N()
	{
		chkWidth.Checked += LayoutChildCheckChanged;
		chkWidth.Unchecked += LayoutChildCheckChanged;
		chkHeight.Checked += LayoutChildCheckChanged;
		chkHeight.Unchecked += LayoutChildCheckChanged;
		chkLockAspectRatio.Checked += LayoutChildCheckChanged;
		chkLockAspectRatio.Unchecked += LayoutChildCheckChanged;
		chkPositionX.Checked += LayoutChildCheckChanged;
		chkPositionX.Unchecked += LayoutChildCheckChanged;
		chkPositionY.Checked += LayoutChildCheckChanged;
		chkPositionY.Unchecked += LayoutChildCheckChanged;
		chkRotation.Checked += LayoutChildCheckChanged;
		chkRotation.Unchecked += LayoutChildCheckChanged;
	}

	private void O()
	{
		chkWidth.Checked -= LayoutChildCheckChanged;
		chkWidth.Unchecked -= LayoutChildCheckChanged;
		chkHeight.Checked -= LayoutChildCheckChanged;
		chkHeight.Unchecked -= LayoutChildCheckChanged;
		chkLockAspectRatio.Checked -= LayoutChildCheckChanged;
		chkLockAspectRatio.Unchecked -= LayoutChildCheckChanged;
		chkPositionX.Checked -= LayoutChildCheckChanged;
		chkPositionX.Unchecked -= LayoutChildCheckChanged;
		chkPositionY.Checked -= LayoutChildCheckChanged;
		chkPositionY.Unchecked -= LayoutChildCheckChanged;
		chkRotation.Checked -= LayoutChildCheckChanged;
		chkRotation.Unchecked -= LayoutChildCheckChanged;
	}

	private void P()
	{
		chkLayout.Checked += LayoutCheckChanged;
		chkLayout.Unchecked += LayoutCheckChanged;
	}

	private void Q()
	{
		chkLayout.Checked -= LayoutCheckChanged;
		chkLayout.Unchecked -= LayoutCheckChanged;
	}

	private void LayoutCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkLayout.IsChecked.Value;
		O();
		chkWidth.IsChecked = value;
		chkHeight.IsChecked = value;
		chkPositionX.IsChecked = value;
		chkPositionY.IsChecked = value;
		chkRotation.IsChecked = value;
		chkLockAspectRatio.IsChecked = value;
		N();
	}

	private void LayoutChildCheckChanged(object sender, RoutedEventArgs e)
	{
		Q();
		bool? isChecked;
		bool? flag = (isChecked = chkWidth.IsChecked);
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
				goto IL_00a7;
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
		bool? isChecked2;
		flag = (isChecked2 = chkHeight.IsChecked);
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
		goto IL_00a7;
		IL_061d:
		chkLayout.IsChecked = null;
		goto IL_0633;
		IL_00a7:
		bool? flag2 = obj;
		isChecked2 = obj;
		bool? obj2;
		if (isChecked2.HasValue)
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
				obj2 = false;
				goto IL_0126;
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
		bool? isChecked3;
		isChecked2 = (isChecked3 = chkLockAspectRatio.IsChecked);
		if (!isChecked2.HasValue)
		{
			obj2 = null;
		}
		else if (isChecked3 != true)
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
			obj2 = flag2;
		}
		goto IL_0126;
		IL_01af:
		bool? obj3;
		bool? flag3 = (bool?)obj3;
		bool? flag4 = (bool?)obj3;
		bool? obj4;
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
			if (flag3 != true)
			{
				obj4 = false;
				goto IL_0235;
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
		bool? isChecked4;
		flag4 = (isChecked4 = chkPositionY.IsChecked);
		if (!flag4.HasValue)
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
		else if (isChecked4 != true)
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
			obj4 = flag3;
		}
		goto IL_0235;
		IL_0633:
		P();
		return;
		IL_0435:
		bool? obj5;
		flag4 = (bool?)obj5;
		flag2 = (bool?)obj5;
		bool? obj6;
		if (flag2.HasValue)
		{
			if (flag4 != true)
			{
				obj6 = false;
				goto IL_04d0;
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
		flag2 = chkPositionX.IsChecked;
		bool? flag5;
		if (!flag2.HasValue)
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
			flag5 = flag2;
		}
		else
		{
			flag5 = flag2 != true;
		}
		bool? flag6 = flag5;
		flag2 = flag5;
		if (flag2.HasValue)
		{
			obj6 = (flag6 == true) & flag4;
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
			obj6 = null;
		}
		goto IL_04d0;
		IL_0580:
		bool? obj7;
		bool? flag7 = (bool?)obj7;
		if (flag7.HasValue)
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
			if (flag7 != true)
			{
				goto IL_061d;
			}
		}
		flag3 = chkRotation.IsChecked;
		bool? flag8;
		if (!flag3.HasValue)
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
			flag8 = flag3;
		}
		else
		{
			flag8 = flag3 != true;
		}
		flag3 = flag8;
		if (flag3 == true)
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
			if (flag7.HasValue)
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
				chkLayout.IsChecked = false;
				goto IL_0633;
			}
		}
		goto IL_061d;
		IL_0126:
		flag6 = obj2;
		isChecked3 = obj2;
		if (isChecked3.HasValue)
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
			if (flag6 != true)
			{
				obj3 = false;
				goto IL_01af;
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
		isChecked3 = (flag4 = chkPositionX.IsChecked);
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
			obj3 = null;
		}
		else if (flag4 != true)
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
			obj3 = false;
		}
		else
		{
			obj3 = flag6;
		}
		goto IL_01af;
		IL_0235:
		flag7 = obj4;
		if (flag7.HasValue)
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
			if (flag7 != true)
			{
				goto IL_029e;
			}
		}
		if (chkRotation.IsChecked != true || !flag7.HasValue)
		{
			goto IL_029e;
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
		chkLayout.IsChecked = true;
		goto IL_0633;
		IL_029e:
		flag = chkWidth.IsChecked;
		bool? flag9;
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
			flag9 = flag;
		}
		else
		{
			flag9 = flag != true;
		}
		isChecked2 = flag9;
		flag = flag9;
		bool? obj8;
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
			if (isChecked2 != true)
			{
				obj8 = false;
				goto IL_037d;
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
		flag = chkHeight.IsChecked;
		bool? flag10;
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
			flag10 = flag;
		}
		else
		{
			flag10 = flag != true;
		}
		isChecked = flag10;
		flag = flag10;
		if (!flag.HasValue)
		{
			obj8 = null;
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
			obj8 = false;
		}
		else
		{
			obj8 = isChecked2;
		}
		goto IL_037d;
		IL_037d:
		isChecked3 = obj8;
		isChecked = obj8;
		if (isChecked.HasValue)
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
				obj5 = false;
				goto IL_0435;
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
		isChecked = chkLockAspectRatio.IsChecked;
		bool? flag11;
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
			flag11 = isChecked;
		}
		else
		{
			flag11 = isChecked != true;
		}
		flag2 = flag11;
		isChecked = flag11;
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
			obj5 = null;
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
			obj5 = false;
		}
		else
		{
			obj5 = isChecked3;
		}
		goto IL_0435;
		IL_04d0:
		isChecked4 = obj6;
		flag6 = obj6;
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
			if (isChecked4 != true)
			{
				obj7 = false;
				goto IL_0580;
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
		flag6 = chkPositionY.IsChecked;
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
		flag3 = flag12;
		flag6 = flag12;
		if (!flag6.HasValue)
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
			obj7 = null;
		}
		else if (flag3 != true)
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
			obj7 = false;
		}
		else
		{
			obj7 = isChecked4;
		}
		goto IL_0580;
	}

	private void R()
	{
		chkShapeEffects.Checked += EffectsChildCheckChanged;
		chkShapeEffects.Unchecked += EffectsChildCheckChanged;
		chkShapeEffects.Indeterminate += EffectsChildCheckChanged;
		chkTextEffects.Checked += EffectsChildCheckChanged;
		chkTextEffects.Unchecked += EffectsChildCheckChanged;
		chkTextEffects.Indeterminate += EffectsChildCheckChanged;
	}

	private void S()
	{
		chkShapeEffects.Checked -= EffectsChildCheckChanged;
		chkShapeEffects.Unchecked -= EffectsChildCheckChanged;
		chkShapeEffects.Indeterminate -= EffectsChildCheckChanged;
		chkTextEffects.Checked -= EffectsChildCheckChanged;
		chkTextEffects.Unchecked -= EffectsChildCheckChanged;
		chkTextEffects.Indeterminate -= EffectsChildCheckChanged;
	}

	private void T()
	{
		chkEffects.Checked += EffectsCheckChanged;
		chkEffects.Unchecked += EffectsCheckChanged;
	}

	private void U()
	{
		chkEffects.Checked -= EffectsCheckChanged;
		chkEffects.Unchecked -= EffectsCheckChanged;
	}

	private void EffectsCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkEffects.IsChecked.Value;
		S();
		chkShapeEffects.IsChecked = value;
		chkTextEffects.IsChecked = value;
		R();
	}

	private void EffectsChildCheckChanged(object sender, RoutedEventArgs e)
	{
		U();
		bool? isChecked = chkShapeEffects.IsChecked;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (isChecked != true)
			{
				goto IL_0085;
			}
		}
		if (chkTextEffects.IsChecked == true)
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
			if (isChecked.HasValue)
			{
				chkEffects.IsChecked = true;
				goto IL_0163;
			}
		}
		goto IL_0085;
		IL_0163:
		T();
		return;
		IL_014d:
		chkEffects.IsChecked = null;
		goto IL_0163;
		IL_0085:
		bool? isChecked2 = chkShapeEffects.IsChecked;
		isChecked = (!isChecked2) ?? isChecked2;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_014d;
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
		isChecked2 = chkTextEffects.IsChecked;
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
		isChecked2 = flag;
		if (isChecked2 == true)
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
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				chkEffects.IsChecked = false;
				goto IL_0163;
			}
		}
		goto IL_014d;
	}

	private void V()
	{
		chkShapeEffects.Checked += ShapeEffectsCheckChanged;
		chkShapeEffects.Unchecked += ShapeEffectsCheckChanged;
		chkShapeEffects.Indeterminate += ShapeEffectsIndeterminate;
		chkShape3D.Checked += ShapeEffectsChildCheckChanged;
		chkShape3D.Unchecked += ShapeEffectsChildCheckChanged;
		chkShapeGlow.Checked += ShapeEffectsChildCheckChanged;
		chkShapeGlow.Unchecked += ShapeEffectsChildCheckChanged;
		chkShapeReflection.Checked += ShapeEffectsChildCheckChanged;
		chkShapeReflection.Unchecked += ShapeEffectsChildCheckChanged;
		chkShapeShadow.Checked += ShapeEffectsChildCheckChanged;
		chkShapeShadow.Unchecked += ShapeEffectsChildCheckChanged;
		chkShapeSoftEdge.Checked += ShapeEffectsChildCheckChanged;
		chkShapeSoftEdge.Unchecked += ShapeEffectsChildCheckChanged;
	}

	private void W()
	{
		chkShapeEffects.Checked -= ShapeEffectsCheckChanged;
		chkShapeEffects.Unchecked -= ShapeEffectsCheckChanged;
		chkShapeEffects.Indeterminate -= ShapeEffectsIndeterminate;
		chkShape3D.Checked -= ShapeEffectsChildCheckChanged;
		chkShape3D.Unchecked -= ShapeEffectsChildCheckChanged;
		chkShapeGlow.Checked -= ShapeEffectsChildCheckChanged;
		chkShapeGlow.Unchecked -= ShapeEffectsChildCheckChanged;
		chkShapeReflection.Checked -= ShapeEffectsChildCheckChanged;
		chkShapeReflection.Unchecked -= ShapeEffectsChildCheckChanged;
		chkShapeShadow.Checked -= ShapeEffectsChildCheckChanged;
		chkShapeShadow.Unchecked -= ShapeEffectsChildCheckChanged;
		chkShapeSoftEdge.Checked -= ShapeEffectsChildCheckChanged;
		chkShapeSoftEdge.Unchecked -= ShapeEffectsChildCheckChanged;
	}

	private void ShapeEffectsCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkShapeEffects.IsChecked.Value;
		W();
		chkShape3D.IsChecked = value;
		chkShapeGlow.IsChecked = value;
		chkShapeReflection.IsChecked = value;
		chkShapeShadow.IsChecked = value;
		chkShapeSoftEdge.IsChecked = value;
		V();
	}

	private void ShapeEffectsIndeterminate(object sender, RoutedEventArgs e)
	{
		chkEffects.IsChecked = null;
	}

	private void ShapeEffectsChildCheckChanged(object sender, RoutedEventArgs e)
	{
		W();
		bool? isChecked;
		bool? flag = (isChecked = chkShape3D.IsChecked);
		bool? obj;
		if (flag.HasValue)
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
			if (isChecked != true)
			{
				obj = false;
				goto IL_00a1;
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
		bool? isChecked2;
		flag = (isChecked2 = chkShapeGlow.IsChecked);
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
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_00a1;
		IL_01a6:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (!((!flag2) ?? false) && chkShapeSoftEdge.IsChecked == true && flag2.HasValue)
		{
			chkShapeEffects.IsChecked = true;
			goto IL_046d;
		}
		flag = chkShape3D.IsChecked;
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
			if (isChecked2 != true)
			{
				obj3 = false;
				goto IL_02b7;
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
		flag = chkShapeGlow.IsChecked;
		flag = (isChecked = (!flag) ?? flag);
		obj3 = ((!flag.HasValue) ? ((bool?)null) : ((isChecked == true) & isChecked2));
		goto IL_02b7;
		IL_02b7:
		bool? flag4 = obj3;
		isChecked = obj3;
		bool? obj4;
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
				goto IL_0358;
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
		isChecked = chkShapeReflection.IsChecked;
		bool? flag5;
		isChecked = (flag5 = (!isChecked) ?? isChecked);
		if (!isChecked.HasValue)
		{
			obj4 = null;
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
			obj4 = false;
		}
		else
		{
			obj4 = flag4;
		}
		goto IL_0358;
		IL_0358:
		bool? flag6 = obj4;
		flag5 = obj4;
		bool? obj5;
		bool? flag7;
		if (!flag5.HasValue || flag6 == true)
		{
			flag5 = chkShapeShadow.IsChecked;
			flag5 = (flag7 = (!flag5) ?? flag5);
			obj5 = ((!flag5.HasValue) ? ((bool?)null) : ((flag7 == true) & flag6));
		}
		else
		{
			obj5 = false;
		}
		flag2 = obj5;
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
				goto IL_0457;
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
		flag7 = chkShapeSoftEdge.IsChecked;
		if (((!flag7) ?? flag7) == true)
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
				chkShapeEffects.IsChecked = false;
				goto IL_046d;
			}
		}
		goto IL_0457;
		IL_046d:
		V();
		return;
		IL_00a1:
		flag5 = obj;
		isChecked2 = obj;
		bool? obj6;
		if (isChecked2.HasValue)
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
				obj6 = false;
				goto IL_0120;
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
		isChecked2 = (flag4 = chkShapeReflection.IsChecked);
		if (isChecked2.HasValue)
		{
			obj6 = (flag4 == true) & flag5;
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
			obj6 = null;
		}
		goto IL_0120;
		IL_0120:
		flag7 = obj6;
		flag4 = obj6;
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
			if (flag7 != true)
			{
				obj2 = false;
				goto IL_01a6;
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
		flag4 = (flag6 = chkShapeShadow.IsChecked);
		if (!flag4.HasValue)
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
		else if (flag6 != true)
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
			obj2 = false;
		}
		else
		{
			obj2 = flag7;
		}
		goto IL_01a6;
		IL_0457:
		chkShapeEffects.IsChecked = null;
		goto IL_046d;
	}

	private void X()
	{
		chkTextEffects.Checked += TextEffectsCheckChanged;
		chkTextEffects.Unchecked += TextEffectsCheckChanged;
		chkTextEffects.Indeterminate += TextEffectsIndeterminate;
		chkTextGlow.Checked += TextEffectsChildCheckChanged;
		chkTextGlow.Unchecked += TextEffectsChildCheckChanged;
		chkTextReflection.Checked += TextEffectsChildCheckChanged;
		chkTextReflection.Unchecked += TextEffectsChildCheckChanged;
		chkTextShadow.Checked += TextEffectsChildCheckChanged;
		chkTextShadow.Unchecked += TextEffectsChildCheckChanged;
		chkTextSoftEdge.Checked += TextEffectsChildCheckChanged;
		chkTextSoftEdge.Unchecked += TextEffectsChildCheckChanged;
		chkText3D.Checked += TextEffectsChildCheckChanged;
		chkText3D.Unchecked += TextEffectsChildCheckChanged;
	}

	private void Y()
	{
		chkTextEffects.Checked -= TextEffectsCheckChanged;
		chkTextEffects.Unchecked -= TextEffectsCheckChanged;
		chkTextEffects.Indeterminate -= TextEffectsIndeterminate;
		chkTextGlow.Checked -= TextEffectsChildCheckChanged;
		chkTextGlow.Unchecked -= TextEffectsChildCheckChanged;
		chkTextReflection.Checked -= TextEffectsChildCheckChanged;
		chkTextReflection.Unchecked -= TextEffectsChildCheckChanged;
		chkTextShadow.Checked -= TextEffectsChildCheckChanged;
		chkTextShadow.Unchecked -= TextEffectsChildCheckChanged;
		chkText3D.Checked -= TextEffectsChildCheckChanged;
		chkText3D.Unchecked -= TextEffectsChildCheckChanged;
	}

	private void TextEffectsCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkTextEffects.IsChecked.Value;
		Y();
		chkTextGlow.IsChecked = value;
		chkTextReflection.IsChecked = value;
		chkTextShadow.IsChecked = value;
		chkTextSoftEdge.IsChecked = value;
		chkText3D.IsChecked = value;
		X();
	}

	private void TextEffectsIndeterminate(object sender, RoutedEventArgs e)
	{
		chkEffects.IsChecked = null;
	}

	private void TextEffectsChildCheckChanged(object sender, RoutedEventArgs e)
	{
		Y();
		bool? isChecked;
		bool? flag = (isChecked = chkTextGlow.IsChecked);
		bool? obj;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (isChecked != true)
			{
				obj = false;
				goto IL_0099;
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
		bool? isChecked2;
		flag = (isChecked2 = chkTextReflection.IsChecked);
		if (flag.HasValue)
		{
			obj = (isChecked2 == true) & isChecked;
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
			obj = null;
		}
		goto IL_0099;
		IL_03a5:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		bool? flag3 = (bool?)obj2;
		bool? obj3;
		if (flag3.HasValue)
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
				goto IL_043f;
			}
		}
		flag3 = chkTextSoftEdge.IsChecked;
		bool? flag4;
		if (!flag3.HasValue)
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
			flag4 = flag3;
		}
		else
		{
			flag4 = flag3 != true;
		}
		bool? flag5 = flag4;
		flag3 = flag4;
		if (flag3.HasValue)
		{
			obj3 = (flag5 == true) & flag2;
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
			obj3 = null;
		}
		goto IL_043f;
		IL_0122:
		bool? obj4;
		flag5 = (bool?)obj4;
		bool? flag6 = (bool?)obj4;
		bool? obj5;
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
			if (flag5 != true)
			{
				obj5 = false;
				goto IL_01a6;
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
		flag6 = (flag2 = chkTextSoftEdge.IsChecked);
		if (!flag6.HasValue)
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
			obj5 = null;
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
			obj5 = false;
		}
		else
		{
			obj5 = flag5;
		}
		goto IL_01a6;
		IL_0099:
		flag3 = obj;
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
			if (flag3 != true)
			{
				obj4 = false;
				goto IL_0122;
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
		isChecked2 = (flag6 = chkTextShadow.IsChecked);
		if (!isChecked2.HasValue)
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
			obj4 = flag3;
		}
		goto IL_0122;
		IL_01a6:
		bool? flag7 = obj5;
		if (flag7.HasValue)
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
			if (flag7 != true)
			{
				goto IL_021d;
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
		if (chkText3D.IsChecked == true)
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
			if (flag7.HasValue)
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
				chkTextEffects.IsChecked = true;
				goto IL_04c7;
			}
		}
		goto IL_021d;
		IL_04b3:
		chkTextEffects.IsChecked = null;
		goto IL_04c7;
		IL_021d:
		flag = chkTextGlow.IsChecked;
		bool? flag8;
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
			flag8 = flag;
		}
		else
		{
			flag8 = flag != true;
		}
		isChecked2 = flag8;
		flag = flag8;
		bool? obj6;
		if (!flag.HasValue || isChecked2 == true)
		{
			flag = chkTextReflection.IsChecked;
			flag = (isChecked = (!flag) ?? flag);
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
				obj6 = null;
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
				obj6 = false;
			}
			else
			{
				obj6 = isChecked2;
			}
		}
		else
		{
			obj6 = false;
		}
		flag6 = obj6;
		isChecked = obj6;
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
			if (flag6 != true)
			{
				obj2 = false;
				goto IL_03a5;
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
		isChecked = chkTextShadow.IsChecked;
		bool? flag9;
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
			flag9 = isChecked;
		}
		else
		{
			flag9 = isChecked != true;
		}
		flag3 = flag9;
		isChecked = flag9;
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
			obj2 = null;
		}
		else if (flag3 != true)
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
			obj2 = flag6;
		}
		goto IL_03a5;
		IL_043f:
		flag7 = obj3;
		if (flag7.HasValue)
		{
			if (flag7 != true)
			{
				goto IL_04b3;
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
		flag5 = chkText3D.IsChecked;
		if (((!flag5) ?? flag5) != true || !flag7.HasValue)
		{
			goto IL_04b3;
		}
		chkTextEffects.IsChecked = false;
		goto IL_04c7;
		IL_04c7:
		X();
	}

	private void Z()
	{
		chkAutoShapeType.Checked += AutoShapeChildCheckChanged;
		chkAutoShapeType.Unchecked += AutoShapeChildCheckChanged;
		chkAdjustments.Checked += AutoShapeChildCheckChanged;
		chkAdjustments.Unchecked += AutoShapeChildCheckChanged;
	}

	private void AB()
	{
		chkAutoShapeType.Checked -= AutoShapeChildCheckChanged;
		chkAutoShapeType.Unchecked -= AutoShapeChildCheckChanged;
		chkAdjustments.Checked -= AutoShapeChildCheckChanged;
		chkAdjustments.Unchecked -= AutoShapeChildCheckChanged;
	}

	private void AutoShapeCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkAutoShape.IsChecked.Value;
		AB();
		chkAutoShapeType.IsChecked = value;
		chkAdjustments.IsChecked = value;
		Z();
	}

	private void AutoShapeChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkAutoShapeType.IsChecked;
		if (isChecked.HasValue)
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
			if (isChecked != true)
			{
				goto IL_0083;
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
		if (chkAdjustments.IsChecked == true && isChecked.HasValue)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkAutoShape.IsChecked = true;
					return;
				}
			}
		}
		goto IL_0083;
		IL_014c:
		chkAutoShape.IsChecked = null;
		return;
		IL_0083:
		bool? isChecked2 = chkAutoShapeType.IsChecked;
		isChecked = (!isChecked2) ?? isChecked2;
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
				goto IL_014c;
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
		isChecked2 = chkAdjustments.IsChecked;
		bool? flag;
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
			flag = isChecked2;
		}
		else
		{
			flag = isChecked2 != true;
		}
		isChecked2 = flag;
		if (isChecked2 == true)
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
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						chkAutoShape.IsChecked = false;
						return;
					}
				}
			}
		}
		goto IL_014c;
	}

	private void AdjustmentsCheckChanged(object sender, RoutedEventArgs e)
	{
		chkAutoShapeType.IsChecked = true;
	}

	private void BB()
	{
		chkPictureScale.Checked += PictureChildCheckChanged;
		chkPictureScale.Unchecked += PictureChildCheckChanged;
		chkSharpness.Checked += PictureChildCheckChanged;
		chkSharpness.Unchecked += PictureChildCheckChanged;
		chkBrightness.Checked += PictureChildCheckChanged;
		chkBrightness.Unchecked += PictureChildCheckChanged;
		chkContrast.Checked += PictureChildCheckChanged;
		chkContrast.Unchecked += PictureChildCheckChanged;
		chkSaturation.Checked += PictureChildCheckChanged;
		chkSaturation.Unchecked += PictureChildCheckChanged;
		chkTemperature.Checked += PictureChildCheckChanged;
		chkTemperature.Unchecked += PictureChildCheckChanged;
	}

	private void CB()
	{
		chkPictureScale.Checked -= PictureChildCheckChanged;
		chkPictureScale.Unchecked -= PictureChildCheckChanged;
		chkSharpness.Checked -= PictureChildCheckChanged;
		chkSharpness.Unchecked -= PictureChildCheckChanged;
		chkBrightness.Checked -= PictureChildCheckChanged;
		chkBrightness.Unchecked -= PictureChildCheckChanged;
		chkContrast.Checked -= PictureChildCheckChanged;
		chkContrast.Unchecked -= PictureChildCheckChanged;
		chkSaturation.Checked -= PictureChildCheckChanged;
		chkSaturation.Unchecked -= PictureChildCheckChanged;
		chkTemperature.Checked -= PictureChildCheckChanged;
		chkTemperature.Unchecked -= PictureChildCheckChanged;
	}

	private void PictureCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkPicture.IsChecked.Value;
		CB();
		chkPictureScale.IsChecked = value;
		chkSharpness.IsChecked = value;
		chkBrightness.IsChecked = value;
		chkContrast.IsChecked = value;
		chkSaturation.IsChecked = value;
		chkTemperature.IsChecked = value;
		BB();
	}

	private void PictureChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked;
		bool? flag = (isChecked = chkPictureScale.IsChecked);
		bool? obj;
		if (flag.HasValue)
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
		}
		bool? isChecked2;
		flag = (isChecked2 = chkSharpness.IsChecked);
		if (!flag.HasValue)
		{
			obj = null;
		}
		else if (isChecked2 != true)
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
			obj = false;
		}
		else
		{
			obj = isChecked;
		}
		goto IL_0095;
		IL_0201:
		bool? obj2;
		bool? flag2 = (bool?)obj2;
		if (flag2.HasValue)
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
			if (flag2 != true)
			{
				goto IL_0270;
			}
		}
		if (chkTemperature.IsChecked == true)
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
					switch (6)
					{
					case 0:
						break;
					default:
						chkPicture.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_0270;
		IL_05c9:
		chkPicture.IsChecked = null;
		return;
		IL_054c:
		bool? obj3;
		flag2 = (bool?)obj3;
		if (flag2.HasValue)
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
			if (flag2 != true)
			{
				goto IL_05c9;
			}
		}
		bool? isChecked3 = chkTemperature.IsChecked;
		if (((!isChecked3) ?? isChecked3) == true)
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
			if (flag2.HasValue)
			{
				chkPicture.IsChecked = false;
				return;
			}
		}
		goto IL_05c9;
		IL_0095:
		bool? flag3 = obj;
		isChecked2 = obj;
		bool? obj4;
		if (isChecked2.HasValue)
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
			if (flag3 != true)
			{
				obj4 = false;
				goto IL_0110;
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
		bool? isChecked4;
		isChecked2 = (isChecked4 = chkBrightness.IsChecked);
		if (isChecked2.HasValue)
		{
			obj4 = (isChecked4 == true) & flag3;
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
			obj4 = null;
		}
		goto IL_0110;
		IL_0270:
		flag = chkPictureScale.IsChecked;
		bool? flag4;
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
			flag4 = flag;
		}
		else
		{
			flag4 = flag != true;
		}
		isChecked2 = flag4;
		flag = flag4;
		bool? obj5;
		if (flag.HasValue)
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
			if (isChecked2 != true)
			{
				obj5 = false;
				goto IL_0361;
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
		flag = chkSharpness.IsChecked;
		bool? flag5;
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
			flag5 = flag;
		}
		else
		{
			flag5 = flag != true;
		}
		isChecked = flag5;
		flag = flag5;
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
			obj5 = null;
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
			obj5 = false;
		}
		else
		{
			obj5 = isChecked2;
		}
		goto IL_0361;
		IL_0110:
		bool? flag6 = obj4;
		isChecked4 = obj4;
		bool? obj6;
		if (isChecked4.HasValue)
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
				obj6 = false;
				goto IL_018f;
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
		bool? isChecked5;
		isChecked4 = (isChecked5 = chkContrast.IsChecked);
		if (!isChecked4.HasValue)
		{
			obj6 = null;
		}
		else if (isChecked5 != true)
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
			obj6 = false;
		}
		else
		{
			obj6 = flag6;
		}
		goto IL_018f;
		IL_018f:
		isChecked3 = obj6;
		isChecked5 = obj6;
		if (isChecked5.HasValue)
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
			if (isChecked3 != true)
			{
				obj2 = false;
				goto IL_0201;
			}
		}
		bool? isChecked6;
		isChecked5 = (isChecked6 = chkSaturation.IsChecked);
		if (isChecked5.HasValue)
		{
			obj2 = (isChecked6 == true) & isChecked3;
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
			obj2 = null;
		}
		goto IL_0201;
		IL_0411:
		bool? obj7;
		isChecked5 = (bool?)obj7;
		flag3 = (bool?)obj7;
		bool? obj8;
		if (!flag3.HasValue || isChecked5 == true)
		{
			flag3 = chkContrast.IsChecked;
			bool? flag7;
			if (!flag3.HasValue)
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
				flag7 = flag3;
			}
			else
			{
				flag7 = flag3 != true;
			}
			flag6 = flag7;
			flag3 = flag7;
			if (!flag3.HasValue)
			{
				obj8 = null;
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
				obj8 = false;
			}
			else
			{
				obj8 = isChecked5;
			}
		}
		else
		{
			obj8 = false;
		}
		isChecked6 = obj8;
		flag6 = obj8;
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
			if (isChecked6 != true)
			{
				obj3 = false;
				goto IL_054c;
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
		flag6 = chkSaturation.IsChecked;
		bool? flag8;
		if (!flag6.HasValue)
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
			flag8 = flag6;
		}
		else
		{
			flag8 = flag6 != true;
		}
		isChecked3 = flag8;
		flag6 = flag8;
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
			obj3 = null;
		}
		else if (isChecked3 != true)
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
			obj3 = isChecked6;
		}
		goto IL_054c;
		IL_0361:
		isChecked4 = obj5;
		isChecked = obj5;
		if (isChecked.HasValue)
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
			if (isChecked4 != true)
			{
				obj7 = false;
				goto IL_0411;
			}
		}
		isChecked = chkBrightness.IsChecked;
		bool? flag9;
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
			flag9 = isChecked;
		}
		else
		{
			flag9 = isChecked != true;
		}
		flag3 = flag9;
		isChecked = flag9;
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
			obj7 = null;
		}
		else if (flag3 != true)
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
			obj7 = false;
		}
		else
		{
			obj7 = isChecked4;
		}
		goto IL_0411;
	}

	private void DB()
	{
		chkPictureScale.Checked += PictureScaleCheckChanged;
		chkPictureScale.Unchecked += PictureScaleCheckChanged;
		chkPictureScale.Indeterminate += PictureScaleIndeterminate;
		chkScaleHeight.Checked += PictureScaleChildCheckChanged;
		chkScaleHeight.Unchecked += PictureScaleChildCheckChanged;
		chkScaleWidth.Checked += PictureScaleChildCheckChanged;
		chkScaleWidth.Unchecked += PictureScaleChildCheckChanged;
	}

	private void EB()
	{
		chkPictureScale.Checked -= PictureScaleCheckChanged;
		chkPictureScale.Unchecked -= PictureScaleCheckChanged;
		chkPictureScale.Indeterminate -= PictureScaleIndeterminate;
		chkScaleHeight.Checked -= PictureScaleChildCheckChanged;
		chkScaleHeight.Unchecked -= PictureScaleChildCheckChanged;
		chkScaleWidth.Checked -= PictureScaleChildCheckChanged;
		chkScaleWidth.Unchecked -= PictureScaleChildCheckChanged;
	}

	private void PictureScaleCheckChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkPictureScale.IsChecked.Value;
		EB();
		chkScaleHeight.IsChecked = value;
		chkScaleWidth.IsChecked = value;
		DB();
	}

	private void PictureScaleIndeterminate(object sender, RoutedEventArgs e)
	{
		chkPicture.IsChecked = null;
	}

	private void PictureScaleChildCheckChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkScaleHeight.IsChecked;
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
				goto IL_0085;
			}
		}
		if (chkScaleWidth.IsChecked == true)
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
			if (isChecked.HasValue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						chkPictureScale.IsChecked = true;
						return;
					}
				}
			}
		}
		goto IL_0085;
		IL_0085:
		bool? isChecked2 = chkScaleHeight.IsChecked;
		bool? flag;
		if (!isChecked2.HasValue)
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
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (isChecked != true)
			{
				goto IL_0147;
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
		isChecked2 = chkScaleWidth.IsChecked;
		if (((!isChecked2) ?? isChecked2) == true && isChecked.HasValue)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkPictureScale.IsChecked = false;
					return;
				}
			}
		}
		goto IL_0147;
		IL_0147:
		chkPictureScale.IsChecked = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
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
			Uri resourceLocator = new Uri(AH.A(139439), UriKind.Relative);
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
			btnCopy = (Button)target;
			return;
		}
		if (connectionId == 2)
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
					((StackPanel)target).MouseDown += ExpandCollapseLayout;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			expLayout = (Polygon)target;
			return;
		}
		if (connectionId == 5)
		{
			chkLayout = (CheckBox)target;
			return;
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
					gridLayout = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			chkHeight = (CheckBox)target;
			return;
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
					txtHeight = (TextBlock)target;
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
					chkWidth = (CheckBox)target;
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
					txtWidth = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			chkLockAspectRatio = (CheckBox)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtLockAspectRatio = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPositionY = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtPositionY = (TextBlock)target;
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
					radTop = (RadioButton)target;
					return;
				}
			}
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
					txtTop = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			radBottom = (RadioButton)target;
			return;
		}
		if (connectionId == 18)
		{
			txtBottom = (TextBlock)target;
			return;
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radMidPointY = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			txtMidpointY = (TextBlock)target;
			return;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkPositionX = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtPositionX = (TextBlock)target;
					return;
				}
			}
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
					radLeft = (RadioButton)target;
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
					txtLeft = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radRight = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtRight = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			radMidPointX = (RadioButton)target;
			return;
		}
		if (connectionId == 28)
		{
			txtMidpointX = (TextBlock)target;
			return;
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
					chkRotation = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtRotation = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseLine;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			expLine = (Polygon)target;
			return;
		}
		if (connectionId == 33)
		{
			chkLine = (CheckBox)target;
			return;
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					gridLine = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			chkLineColor = (CheckBox)target;
			return;
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					rectLineColor = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			txtLineColor = (TextBlock)target;
			return;
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkLineWeight = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtLineWeight = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLineStyle = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtLineStyle = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseFill;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			expFill = (Polygon)target;
			return;
		}
		if (connectionId == 44)
		{
			chkFill = (CheckBox)target;
			return;
		}
		if (connectionId == 45)
		{
			gridFill = (Grid)target;
			return;
		}
		if (connectionId == 46)
		{
			chkFillColor = (CheckBox)target;
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
					rectFillColor = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtFillColor = (TextBlock)target;
					return;
				}
			}
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
					chkFillType = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 50)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtFillType = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 51)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkFillTransparency = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 52)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtFillTransparency = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 53)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseFont;
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
					expFont = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 55)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFont = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 56)
		{
			gridFont = (Grid)target;
			return;
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkFontColor = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 58)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					rectFontColor = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 59)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtFontColor = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 60)
		{
			chkFontSize = (CheckBox)target;
			return;
		}
		if (connectionId == 61)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtFontSize = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 62)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkFontName = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 63)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtFontName = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 64)
		{
			chkDecoration = (CheckBox)target;
			return;
		}
		if (connectionId == 65)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtDecoration = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 66)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseTextBox;
					return;
				}
			}
		}
		if (connectionId == 67)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					expTextBox = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 68)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkTextBox = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 69)
		{
			gridTextBox = (Grid)target;
			return;
		}
		if (connectionId == 70)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkBullets = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 71)
		{
			txtBullets = (TextBlock)target;
			return;
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
					chkIndents = (CheckBox)target;
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
					txtIndents = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 74)
		{
			chkLineSpacing = (CheckBox)target;
			return;
		}
		if (connectionId == 75)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtLineSpacing = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 76)
		{
			chkMargins = (CheckBox)target;
			return;
		}
		if (connectionId == 77)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtMargins = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 78)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkAutoSize = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 79)
		{
			txtAutoSize = (TextBlock)target;
			return;
		}
		if (connectionId == 80)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkWordWrap = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 81)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtWordWrap = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 82)
		{
			chkAlignH = (CheckBox)target;
			return;
		}
		if (connectionId == 83)
		{
			txtAlignH = (TextBlock)target;
			return;
		}
		if (connectionId == 84)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkAlignV = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 85)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtAlignV = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 86)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkOrientation = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 87)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtOrientation = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 88)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseAutoShape;
					return;
				}
			}
		}
		if (connectionId == 89)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					expAutoShape = (Polygon)target;
					return;
				}
			}
		}
		if (connectionId == 90)
		{
			chkAutoShape = (CheckBox)target;
			return;
		}
		if (connectionId == 91)
		{
			gridAutoShape = (Grid)target;
			return;
		}
		if (connectionId == 92)
		{
			chkAutoShapeType = (CheckBox)target;
			return;
		}
		if (connectionId == 93)
		{
			txtAutoShapeType = (TextBlock)target;
			return;
		}
		if (connectionId == 94)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkAdjustments = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 95)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtAdjustments = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 96)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapsePicture;
					return;
				}
			}
		}
		if (connectionId == 97)
		{
			expPicture = (Polygon)target;
			return;
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
					chkPicture = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 99)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					gridPicture = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 100)
		{
			chkPictureScale = (CheckBox)target;
			return;
		}
		if (connectionId == 101)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtPictureScale = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 102)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkScaleHeight = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 103)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtScaleHeight = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 104)
		{
			chkScaleWidth = (CheckBox)target;
			return;
		}
		if (connectionId == 105)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtScaleWidth = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 106)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkSharpness = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 107)
		{
			txtSharpness = (TextBlock)target;
			return;
		}
		if (connectionId == 108)
		{
			chkBrightness = (CheckBox)target;
			return;
		}
		if (connectionId == 109)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtBrightness = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 110)
		{
			chkContrast = (CheckBox)target;
			return;
		}
		if (connectionId == 111)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtContrast = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 112)
		{
			chkSaturation = (CheckBox)target;
			return;
		}
		if (connectionId == 113)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtSaturation = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 114)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkTemperature = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 115)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtTemperature = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 116)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((StackPanel)target).MouseDown += ExpandCollapseEffects;
					return;
				}
			}
		}
		if (connectionId == 117)
		{
			expEffects = (Polygon)target;
			return;
		}
		if (connectionId == 118)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkEffects = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 119)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					gridEffects = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 120)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkShapeEffects = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 121)
		{
			txtShapeEffects = (TextBlock)target;
			return;
		}
		if (connectionId == 122)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkShapeShadow = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 123)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					rectShapeShadow = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 124)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtShapeShadow = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 125)
		{
			chkShapeReflection = (CheckBox)target;
			return;
		}
		if (connectionId == 126)
		{
			txtShapeReflection = (TextBlock)target;
			return;
		}
		if (connectionId == 127)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkShapeGlow = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 128)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					rectShapeGlow = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 129)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtShapeGlow = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 130)
		{
			chkShapeSoftEdge = (CheckBox)target;
			return;
		}
		if (connectionId == 131)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtShapeSoftEdge = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 132)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkShape3D = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 133)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					txtShape3D = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 134)
		{
			chkTextEffects = (CheckBox)target;
			return;
		}
		if (connectionId == 135)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtTextEffects = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 136)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkTextShadow = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 137)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					rectTextShadow = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 138)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtTextShadow = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 139)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkTextReflection = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 140)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtTextReflection = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 141)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkTextGlow = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 142)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					rectTextGlow = (System.Windows.Shapes.Rectangle)target;
					return;
				}
			}
		}
		if (connectionId == 143)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtTextGlow = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 144)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkTextSoftEdge = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 145)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					txtTextSoftEdge = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 146)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkText3D = (CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 147)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					txtText3D = (TextBlock)target;
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
