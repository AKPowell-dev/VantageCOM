using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.UI;

namespace PowerPointAddIn1.DeckCheck.Reformat;

[DesignerGenerated]
public sealed class wpfReformat : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Tuple<int, IndexedObject>, int> A;

		public static Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> A;

		public static Func<int, IEnumerable<Tuple<int, IndexedObject>>, V<int, IEnumerable<Tuple<int, IndexedObject>>>> A;

		public static Func<FontColorItem, int> A;

		public static Func<Tuple<int, IndexedObject>, int> B;

		public static Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> B;

		public static Func<int, IEnumerable<Tuple<int, IndexedObject>>, W<int, IEnumerable<Tuple<int, IndexedObject>>>> A;

		public static Func<W<int, IEnumerable<Tuple<int, IndexedObject>>>, int> A;

		public static Func<Tuple<int, IndexedObject>, int> C;

		public static Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> C;

		public static Func<int, IEnumerable<Tuple<int, IndexedObject>>, W<int, IEnumerable<Tuple<int, IndexedObject>>>> B;

		public static Func<W<int, IEnumerable<Tuple<int, IndexedObject>>>, int> B;

		public static Func<Tuple<float, IndexedObject>, float> A;

		public static Func<Tuple<float, IndexedObject>, Tuple<float, IndexedObject>> A;

		public static Func<float, IEnumerable<Tuple<float, IndexedObject>>, W<float, IEnumerable<Tuple<float, IndexedObject>>>> A;

		public static Func<W<float, IEnumerable<Tuple<float, IndexedObject>>>, int> A;

		public static Func<Tuple<string, IndexedObject>, string> A;

		public static Func<Tuple<string, IndexedObject>, Tuple<string, IndexedObject>> A;

		public static Func<string, IEnumerable<Tuple<string, IndexedObject>>, X<string, IEnumerable<Tuple<string, IndexedObject>>>> A;

		public static Func<FontFamilyItem, int> A;

		public static Func<FontFamilyItem, string> A;

		public static Func<Tuple<FontStyle, IndexedObject>, FontStyle> A;

		public static Func<Tuple<FontStyle, IndexedObject>, Tuple<FontStyle, IndexedObject>> A;

		public static Func<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>, X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>> A;

		public static Func<FontStyleItem, int> A;

		public static Func<FontStyleItem, FontStyleOption> A;

		public static Func<Tuple<TextDecoration, IndexedObject>, TextDecoration> A;

		public static Func<Tuple<TextDecoration, IndexedObject>, Tuple<TextDecoration, IndexedObject>> A;

		public static Func<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>, X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>> A;

		public static Func<TextDecorationItem, int> A;

		public static Func<TextDecorationItem, TextDecorationOption> A;

		public static Func<Tuple<int, IndexedObject>, int> D;

		public static Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> D;

		public static Func<int, IEnumerable<Tuple<int, IndexedObject>>, W<int, IEnumerable<Tuple<int, IndexedObject>>>> C;

		public static Func<W<int, IEnumerable<Tuple<int, IndexedObject>>>, int> C;

		public static Func<Tuple<float, IndexedObject>, float> B;

		public static Func<Tuple<float, IndexedObject>, Tuple<float, IndexedObject>> B;

		public static Func<float, IEnumerable<Tuple<float, IndexedObject>>, W<float, IEnumerable<Tuple<float, IndexedObject>>>> B;

		public static Func<W<float, IEnumerable<Tuple<float, IndexedObject>>>, int> B;

		public static Func<Tuple<LineSpacing, IndexedObject>, LineSpacing> A;

		public static Func<Tuple<LineSpacing, IndexedObject>, Tuple<LineSpacing, IndexedObject>> A;

		public static Func<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>, W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>> A;

		public static Func<W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>, int> A;

		public static Func<Tuple<Indent, IndexedObject>, Indent> A;

		public static Func<Tuple<Indent, IndexedObject>, Tuple<Indent, IndexedObject>> A;

		public static Func<Indent, IEnumerable<Tuple<Indent, IndexedObject>>, W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>> A;

		public static Func<W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>, int> A;

		public static Func<Tuple<BulletStyle, IndexedObject>, BulletStyle> A;

		public static Func<Tuple<BulletStyle, IndexedObject>, Tuple<BulletStyle, IndexedObject>> A;

		public static Func<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>, W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>> A;

		public static Func<W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>, int> A;

		public static Func<Tuple<Margins, IndexedObject>, Margins> A;

		public static Func<Tuple<Margins, IndexedObject>, Tuple<Margins, IndexedObject>> A;

		public static Func<Margins, IEnumerable<Tuple<Margins, IndexedObject>>, W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> A;

		public static Func<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>, int> A;

		public static Func<Tuple<Margins, IndexedObject>, Margins> B;

		public static Func<Tuple<Margins, IndexedObject>, Tuple<Margins, IndexedObject>> B;

		public static Func<Margins, IEnumerable<Tuple<Margins, IndexedObject>>, W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> B;

		public static Func<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>, int> B;

		public static Func<Rect, double> A;

		public static Func<Rect, double> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<int, IndexedObject> A(Tuple<int, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal V<int, IEnumerable<Tuple<int, IndexedObject>>> A(int A, IEnumerable<Tuple<int, IndexedObject>> B)
		{
			return new V<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(FontColorItem A)
		{
			return ((BaseItem)A).Quantity;
		}

		[SpecialName]
		internal int B(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<int, IndexedObject> B(Tuple<int, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<int, IEnumerable<Tuple<int, IndexedObject>>> A(int A, IEnumerable<Tuple<int, IndexedObject>> B)
		{
			return new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<int, IEnumerable<Tuple<int, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal int C(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<int, IndexedObject> C(Tuple<int, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<int, IEnumerable<Tuple<int, IndexedObject>>> B(int A, IEnumerable<Tuple<int, IndexedObject>> B)
		{
			return new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int B(W<int, IEnumerable<Tuple<int, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal float A(Tuple<float, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<float, IndexedObject> D(Tuple<float, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<float, IEnumerable<Tuple<float, IndexedObject>>> A(float A, IEnumerable<Tuple<float, IndexedObject>> B)
		{
			return new W<float, IEnumerable<Tuple<float, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<float, IEnumerable<Tuple<float, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal string A(Tuple<string, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<string, IndexedObject> E(Tuple<string, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal X<string, IEnumerable<Tuple<string, IndexedObject>>> A(string A, IEnumerable<Tuple<string, IndexedObject>> B)
		{
			return new X<string, IEnumerable<Tuple<string, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(FontFamilyItem A)
		{
			return ((BaseItem)A).Quantity;
		}

		[SpecialName]
		internal string A(FontFamilyItem A)
		{
			return ((BaseItem)A).Label;
		}

		[SpecialName]
		internal FontStyle A(Tuple<FontStyle, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<FontStyle, IndexedObject> F(Tuple<FontStyle, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>> A(FontStyle A, IEnumerable<Tuple<FontStyle, IndexedObject>> B)
		{
			return new X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(FontStyleItem A)
		{
			return ((BaseItem)A).Quantity;
		}

		[SpecialName]
		internal FontStyleOption A(FontStyleItem A)
		{
			return new FontStyleOption(A.Style, ((BaseItem)A).Label);
		}

		[SpecialName]
		internal TextDecoration A(Tuple<TextDecoration, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<TextDecoration, IndexedObject> G(Tuple<TextDecoration, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>> A(TextDecoration A, IEnumerable<Tuple<TextDecoration, IndexedObject>> B)
		{
			return new X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(TextDecorationItem A)
		{
			return ((BaseItem)A).Quantity;
		}

		[SpecialName]
		internal TextDecorationOption A(TextDecorationItem A)
		{
			return new TextDecorationOption(A.Decoration, ((BaseItem)A).Label);
		}

		[SpecialName]
		internal int D(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<int, IndexedObject> H(Tuple<int, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<int, IEnumerable<Tuple<int, IndexedObject>>> C(int A, IEnumerable<Tuple<int, IndexedObject>> B)
		{
			return new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int C(W<int, IEnumerable<Tuple<int, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal float B(Tuple<float, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<float, IndexedObject> I(Tuple<float, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<float, IEnumerable<Tuple<float, IndexedObject>>> B(float A, IEnumerable<Tuple<float, IndexedObject>> B)
		{
			return new W<float, IEnumerable<Tuple<float, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int B(W<float, IEnumerable<Tuple<float, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal LineSpacing A(Tuple<LineSpacing, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<LineSpacing, IndexedObject> J(Tuple<LineSpacing, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>> A(LineSpacing A, IEnumerable<Tuple<LineSpacing, IndexedObject>> B)
		{
			return new W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal Indent A(Tuple<Indent, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<Indent, IndexedObject> K(Tuple<Indent, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>> A(Indent A, IEnumerable<Tuple<Indent, IndexedObject>> B)
		{
			return new W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal BulletStyle A(Tuple<BulletStyle, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<BulletStyle, IndexedObject> L(Tuple<BulletStyle, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>> A(BulletStyle A, IEnumerable<Tuple<BulletStyle, IndexedObject>> B)
		{
			return new W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal Margins A(Tuple<Margins, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<Margins, IndexedObject> M(Tuple<Margins, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> A(Margins A, IEnumerable<Tuple<Margins, IndexedObject>> B)
		{
			return new W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int A(W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal Margins B(Tuple<Margins, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal Tuple<Margins, IndexedObject> N(Tuple<Margins, IndexedObject> A)
		{
			return A;
		}

		[SpecialName]
		internal W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> B(Margins A, IEnumerable<Tuple<Margins, IndexedObject>> B)
		{
			return new W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>(A, B);
		}

		[SpecialName]
		internal int B(W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> A)
		{
			return A.g.Count();
		}

		[SpecialName]
		internal double A(Rect A)
		{
			return A.Left;
		}

		[SpecialName]
		internal double B(Rect A)
		{
			return A.Top;
		}
	}

	[CompilerGenerated]
	internal sealed class FC
	{
		public FontStyleItem A;

		public FC(FC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(FontStyleOption A)
		{
			if (Operators.CompareString(A.Style.Family, this.A.Style.Family, TextCompare: false) == 0)
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
						return A.Style.Size == this.A.Style.Size;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class GC
	{
		public TextDecorationItem A;

		public GC(GC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(TextDecorationOption A)
		{
			if (A.Decoration.Bold == this.A.Decoration.Bold)
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
				if (A.Decoration.Italic == this.A.Decoration.Italic)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							return A.Decoration.UnderlineStyle == this.A.Decoration.UnderlineStyle;
						}
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class HC
	{
		public int A;

		[SpecialName]
		internal bool A(PaletteColor A)
		{
			return A.OLE == this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class IC
	{
		public MsoLineDashStyle A;

		[SpecialName]
		internal bool A(LineDashStyle A)
		{
			return A.Style == this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class JC
	{
		public float A;

		[SpecialName]
		internal bool A(LineWeight A)
		{
			return A.Weight == this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class KC
	{
		public int A;

		[SpecialName]
		internal bool A(NavigationItem A)
		{
			if (A.IndexedObject.SlideOrLayout is Slide)
			{
				return ((Slide)A.IndexedObject.SlideOrLayout).SlideIndex == this.A;
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class LC
	{
		public object A;

		public float A;

		public float B;

		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public LC(LC A)
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
				this.A = A.A;
				this.A = A.A;
				this.B = A.B;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetTextRangeRectangle((TextRange2)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect B()
		{
			Type typeFromHandle = typeof(MarchingAnts);
			string memberName = AH.A(170589);
			object[] obj = new object[1] { this.A };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (obj3 == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return default(Rect);
					}
				}
			}
			return (Rect)obj3;
		}

		[SpecialName]
		internal Rect C()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(((Cell)this.A).Shape);
		}

		[SpecialName]
		internal Rect D()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(this.A);
		}

		[SpecialName]
		internal Rect E()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle((PlotArea)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect F()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle(this.A.Chart.PlotArea, this.A, this.B);
		}

		[SpecialName]
		internal Rect G()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLabelRectangle((IMsoDataLabel)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect H()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle(this.A.Chart.PlotArea, this.A, this.B);
		}

		[SpecialName]
		internal Rect I()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetChartTitleRectangle((ChartTitle)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect J()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisRectangle((Axis)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect K()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisTitleRectangle((AxisTitle)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect L()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaOuterRectangle(this.A);
		}

		[SpecialName]
		internal Rect M()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(this.A);
		}

		[SpecialName]
		internal Rect N()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendRectangle((Legend)this.A, this.A, this.B);
		}

		[SpecialName]
		internal Rect O()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendRectangle(this.A.Chart.Legend, this.A, this.B);
		}

		[SpecialName]
		internal Rect P()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(this.A);
		}

		[SpecialName]
		internal Rect Q()
		{
			Type typeFromHandle = typeof(MarchingAnts);
			string memberName = AH.A(170589);
			object[] obj = new object[1] { this.A };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
			if (array2[0])
			{
				this.A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (obj3 == null)
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
						return default(Rect);
					}
				}
			}
			return (Rect)obj3;
		}

		[SpecialName]
		internal Rect R()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendKeyRectangle(this.A, (IMsoLegendKey)this.A);
		}

		[SpecialName]
		internal Rect S()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetChartPointRectangle(this.A, (ChartPoint)this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class MC
	{
		public TextRange2 A;

		public NC A;

		public MC(MC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetObjectRectangle(this.A.A.A + this.A.BoundLeft - this.A.A, this.A.A.B + this.A.BoundTop, this.A.A, this.A.BoundHeight);
		}
	}

	[CompilerGenerated]
	internal sealed class NC
	{
		public float A;

		public LC A;

		public NC(NC A)
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
	}

	[CompilerGenerated]
	internal sealed class OC
	{
		public IMsoDataLabel A;

		public LC A;

		public OC(OC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLabelRectangle(this.A, this.A.A, this.A.B);
		}
	}

	[CompilerGenerated]
	internal sealed class PC
	{
		public Axis A;

		public LC A;

		public PC(PC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Rect A()
		{
			return PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisRectangle(this.A, this.A.A, this.A.B);
		}
	}

	[CompilerGenerated]
	internal sealed class QC
	{
		public int A;

		public wpfReformat A;

		public QC(QC A)
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
		internal bool A(PaletteColor A)
		{
			return A.Color == this.A.ColorOptions[this.A].Color;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private bool m_A;

	private bool m_B;

	[CompilerGenerated]
	private Conventions m_A;

	private ICollectionView m_A;

	[CompilerGenerated]
	private List<BaseItem> m_A;

	private DataTemplate m_A;

	private static wpfMarchingAnts m_A;

	private ObservableCollection<FillColorItem> m_A;

	private ObservableCollection<FillTransparencyItem> m_A;

	private ObservableCollection<BorderColorItem> m_A;

	private ObservableCollection<FontColorItem> m_A;

	private ObservableCollection<FontFamilyItem> m_A;

	private ObservableCollection<FontStyleItem> m_A;

	private ObservableCollection<TextDecorationItem> m_A;

	private ObservableCollection<BorderDashItem> m_A;

	private ObservableCollection<BorderWeightItem> m_A;

	private ObservableCollection<TextBoxMarginsItem> m_A;

	private ObservableCollection<CellMarginsItem> m_A;

	private ObservableCollection<ParagraphSpacingItem> m_A;

	private ObservableCollection<IndentItem> m_A;

	private ObservableCollection<BulletStyleItem> m_A;

	private ObservableCollection<PaletteColor> m_A;

	private ObservableCollection<float> m_A;

	private ObservableCollection<LineDashStyle> m_A;

	private ObservableCollection<LineWeight> m_A;

	private ObservableCollection<string> m_A;

	private ObservableCollection<FontStyleOption> m_A;

	private ObservableCollection<TextDecorationOption> m_A;

	private ObservableCollection<MarginsOption> m_A;

	private ObservableCollection<MarginsOption> m_B;

	private ObservableCollection<ParagraphSpacingOption> m_A;

	private ObservableCollection<IndentOption> m_A;

	private ObservableCollection<BulletStyleOption> m_A;

	private Visibility m_A;

	private Visibility m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxResults")]
	private System.Windows.Controls.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("bdrWorking")]
	private Border m_A;

	private bool m_C;

	private Conventions Conventions
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

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10961));
		}
	}

	private List<BaseItem> AllItems
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

	private DataTemplate NavItemDataTemplate
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51458));
		}
	}

	public wpfMarchingAnts MarchingAnts
	{
		get
		{
			return wpfReformat.m_A;
		}
		set
		{
			wpfReformat.m_A = value;
		}
	}

	public ObservableCollection<FillColorItem> FillColors
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51497));
		}
	}

	public ObservableCollection<FillTransparencyItem> FillTransparencies
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51518));
			int transparenciesVisibility;
			if (value.Count <= 1)
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
				transparenciesVisibility = 2;
			}
			else
			{
				transparenciesVisibility = 0;
			}
			TransparenciesVisibility = (Visibility)transparenciesVisibility;
		}
	}

	public ObservableCollection<BorderColorItem> BorderColors
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51555));
		}
	}

	public ObservableCollection<FontColorItem> FontColors
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51580));
		}
	}

	public ObservableCollection<FontFamilyItem> FontFamilies
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51601));
			int fontFamiliesVisibility;
			if (value.Count <= 1)
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
				fontFamiliesVisibility = 2;
			}
			else
			{
				fontFamiliesVisibility = 0;
			}
			FontFamiliesVisibility = (Visibility)fontFamiliesVisibility;
		}
	}

	public ObservableCollection<FontStyleItem> FontStyles
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51626));
		}
	}

	public ObservableCollection<TextDecorationItem> TextDecorations
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51647));
		}
	}

	public ObservableCollection<BorderDashItem> BorderDashStyles
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51678));
		}
	}

	public ObservableCollection<BorderWeightItem> BorderWeights
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51711));
		}
	}

	public ObservableCollection<TextBoxMarginsItem> TextBoxMargins
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51738));
		}
	}

	public ObservableCollection<CellMarginsItem> CellMargins
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51767));
		}
	}

	public ObservableCollection<ParagraphSpacingItem> ParagraphSpacing
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51790));
		}
	}

	public ObservableCollection<IndentItem> Indents
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51823));
		}
	}

	public ObservableCollection<BulletStyleItem> BulletStyles
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51838));
		}
	}

	public ObservableCollection<PaletteColor> ColorOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51863));
		}
	}

	public ObservableCollection<float> FillTransparencyOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51888));
		}
	}

	public ObservableCollection<LineDashStyle> LineDashStyleOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51935));
		}
	}

	public ObservableCollection<LineWeight> LineWeightOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(51976));
		}
	}

	public ObservableCollection<string> FontFamilyOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52011));
		}
	}

	public ObservableCollection<FontStyleOption> FontStyleOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52046));
		}
	}

	public ObservableCollection<TextDecorationOption> TextDecorationOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52079));
		}
	}

	public ObservableCollection<MarginsOption> TextBoxMarginsOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52122));
		}
	}

	public ObservableCollection<MarginsOption> CellMarginsOptions
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(52165));
		}
	}

	public ObservableCollection<ParagraphSpacingOption> ParagraphSpacingOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52202));
		}
	}

	public ObservableCollection<IndentOption> IndentOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52249));
		}
	}

	public ObservableCollection<BulletStyleOption> BulletStyleOptions
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52276));
		}
	}

	public Visibility TransparenciesVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(52313));
		}
	}

	public Visibility FontFamiliesVisibility
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(52362));
		}
	}

	internal virtual System.Windows.Controls.ListBox lbxResults
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
			System.Windows.Input.KeyEventHandler value2 = lbxResults_PreviewKeyDown;
			System.Windows.Controls.ListBox listBox = this.m_A;
			if (listBox != null)
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
				listBox.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			listBox = this.m_A;
			if (listBox != null)
			{
				listBox.PreviewKeyDown += value2;
			}
		}
	}

	internal virtual Border bdrWorking
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
	}

	public wpfReformat(Conventions conv)
	{
		base.Loaded += wpfReformat_Loaded;
		this.m_A = false;
		this.m_B = false;
		this.m_A = null;
		InitializeComponent();
		Conventions = conv;
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

	private void wpfReformat_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Expected O, but got Unknown
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Expected O, but got Unknown
		//IL_00bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Expected O, but got Unknown
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d2: Expected O, but got Unknown
		//IL_00d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Expected O, but got Unknown
		//IL_00e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ea: Expected O, but got Unknown
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Expected O, but got Unknown
		//IL_0111: Unknown result type (might be due to invalid IL or missing references)
		//IL_011b: Expected O, but got Unknown
		//IL_0121: Unknown result type (might be due to invalid IL or missing references)
		//IL_012b: Expected O, but got Unknown
		//IL_0131: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Expected O, but got Unknown
		//IL_0141: Unknown result type (might be due to invalid IL or missing references)
		//IL_014b: Expected O, but got Unknown
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_015b: Expected O, but got Unknown
		//IL_0161: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Expected O, but got Unknown
		//IL_0171: Unknown result type (might be due to invalid IL or missing references)
		//IL_017b: Expected O, but got Unknown
		//IL_0181: Unknown result type (might be due to invalid IL or missing references)
		//IL_018b: Expected O, but got Unknown
		//IL_0190: Unknown result type (might be due to invalid IL or missing references)
		//IL_019a: Expected O, but got Unknown
		ColorOptions = new ObservableCollection<PaletteColor>();
		using (List<PaletteColor>.Enumerator enumerator = ((Conventions)Conventions).ColorPalette.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				PaletteColor current = enumerator.Current;
				ColorOptions.Add(new PaletteColor(current.Color, current.IsUsed, true, current.Name));
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
				break;
			}
		}
		LineDashStyleOptions = new ObservableCollection<LineDashStyle>();
		ObservableCollection<LineDashStyle> lineDashStyleOptions = LineDashStyleOptions;
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineSolid));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineSysDot));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineSysDash));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineDash));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineDashDot));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineLongDash));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineLongDashDot));
		lineDashStyleOptions.Add(new LineDashStyle(MsoLineDashStyle.msoLineLongDashDotDot));
		_ = null;
		LineWeightOptions = new ObservableCollection<LineWeight>();
		ObservableCollection<LineWeight> lineWeightOptions = LineWeightOptions;
		lineWeightOptions.Add(new LineWeight(0.25f));
		lineWeightOptions.Add(new LineWeight(0.5f));
		lineWeightOptions.Add(new LineWeight(0.75f));
		lineWeightOptions.Add(new LineWeight(1f));
		lineWeightOptions.Add(new LineWeight(1.5f));
		lineWeightOptions.Add(new LineWeight(2.25f));
		lineWeightOptions.Add(new LineWeight(3f));
		lineWeightOptions.Add(new LineWeight(4.5f));
		lineWeightOptions.Add(new LineWeight(6f));
		_ = null;
		AllItems = new List<BaseItem>();
		NavItemDataTemplate = A(AH.A(52407));
		A();
		E();
		F();
		G();
		B();
		D();
		C();
		H();
		I();
		J();
		K();
		L();
		M();
		N();
		AllItems = new List<BaseItem>();
		List<BaseItem> allItems = AllItems;
		using (IEnumerator<FontColorItem> enumerator2 = FontColors.GetEnumerator())
		{
			while (enumerator2.MoveNext())
			{
				BaseItem current2 = enumerator2.Current;
				allItems.Add(current2);
				allItems.AddRange(current2.Objects);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_0262;
				}
				continue;
				end_IL_0262:
				break;
			}
		}
		IEnumerator<FontFamilyItem> enumerator3 = default(IEnumerator<FontFamilyItem>);
		try
		{
			enumerator3 = FontFamilies.GetEnumerator();
			while (enumerator3.MoveNext())
			{
				BaseItem current3 = enumerator3.Current;
				allItems.Add(current3);
				allItems.AddRange(current3.Objects);
			}
		}
		finally
		{
			if (enumerator3 != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					enumerator3.Dispose();
					break;
				}
			}
		}
		using (IEnumerator<FontStyleItem> enumerator4 = FontStyles.GetEnumerator())
		{
			while (enumerator4.MoveNext())
			{
				BaseItem current4 = enumerator4.Current;
				allItems.Add(current4);
				allItems.AddRange(current4.Objects);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0310;
				}
				continue;
				end_IL_0310:
				break;
			}
		}
		foreach (TextDecorationItem textDecoration in TextDecorations)
		{
			allItems.Add(textDecoration);
			allItems.AddRange(textDecoration.Objects);
		}
		IEnumerator<FillColorItem> enumerator6 = default(IEnumerator<FillColorItem>);
		try
		{
			enumerator6 = FillColors.GetEnumerator();
			while (enumerator6.MoveNext())
			{
				BaseItem current6 = enumerator6.Current;
				allItems.Add(current6);
				allItems.AddRange(current6.Objects);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_03b4;
				}
				continue;
				end_IL_03b4:
				break;
			}
		}
		finally
		{
			if (enumerator6 != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator6.Dispose();
					break;
				}
			}
		}
		using (IEnumerator<FillTransparencyItem> enumerator7 = FillTransparencies.GetEnumerator())
		{
			while (enumerator7.MoveNext())
			{
				BaseItem current7 = enumerator7.Current;
				allItems.Add(current7);
				allItems.AddRange(current7.Objects);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0415;
				}
				continue;
				end_IL_0415:
				break;
			}
		}
		using (IEnumerator<BorderColorItem> enumerator8 = BorderColors.GetEnumerator())
		{
			while (enumerator8.MoveNext())
			{
				BaseItem current8 = enumerator8.Current;
				allItems.Add(current8);
				allItems.AddRange(current8.Objects);
			}
			while (true)
			{
				switch (3)
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
		IEnumerator<BorderDashItem> enumerator9 = default(IEnumerator<BorderDashItem>);
		try
		{
			enumerator9 = BorderDashStyles.GetEnumerator();
			while (enumerator9.MoveNext())
			{
				BaseItem current9 = enumerator9.Current;
				allItems.Add(current9);
				allItems.AddRange(current9.Objects);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_04bf;
				}
				continue;
				end_IL_04bf:
				break;
			}
		}
		finally
		{
			if (enumerator9 != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					enumerator9.Dispose();
					break;
				}
			}
		}
		IEnumerator<BorderWeightItem> enumerator10 = default(IEnumerator<BorderWeightItem>);
		try
		{
			enumerator10 = BorderWeights.GetEnumerator();
			while (enumerator10.MoveNext())
			{
				BaseItem current10 = enumerator10.Current;
				allItems.Add(current10);
				allItems.AddRange(current10.Objects);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0520;
				}
				continue;
				end_IL_0520:
				break;
			}
		}
		finally
		{
			if (enumerator10 != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator10.Dispose();
					break;
				}
			}
		}
		IEnumerator<ParagraphSpacingItem> enumerator11 = default(IEnumerator<ParagraphSpacingItem>);
		try
		{
			enumerator11 = ParagraphSpacing.GetEnumerator();
			while (enumerator11.MoveNext())
			{
				BaseItem current11 = enumerator11.Current;
				allItems.Add(current11);
				allItems.AddRange(current11.Objects);
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0581;
				}
				continue;
				end_IL_0581:
				break;
			}
		}
		finally
		{
			if (enumerator11 != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					enumerator11.Dispose();
					break;
				}
			}
		}
		IEnumerator<IndentItem> enumerator12 = default(IEnumerator<IndentItem>);
		try
		{
			enumerator12 = Indents.GetEnumerator();
			while (enumerator12.MoveNext())
			{
				BaseItem current12 = enumerator12.Current;
				allItems.Add(current12);
				allItems.AddRange(current12.Objects);
			}
		}
		finally
		{
			if (enumerator12 != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					enumerator12.Dispose();
					break;
				}
			}
		}
		IEnumerator<BulletStyleItem> enumerator13 = default(IEnumerator<BulletStyleItem>);
		try
		{
			enumerator13 = BulletStyles.GetEnumerator();
			while (enumerator13.MoveNext())
			{
				BaseItem current13 = enumerator13.Current;
				allItems.Add(current13);
				allItems.AddRange(current13.Objects);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0639;
				}
				continue;
				end_IL_0639:
				break;
			}
		}
		finally
		{
			if (enumerator13 != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator13.Dispose();
					break;
				}
			}
		}
		using (IEnumerator<TextBoxMarginsItem> enumerator14 = TextBoxMargins.GetEnumerator())
		{
			while (enumerator14.MoveNext())
			{
				BaseItem current14 = enumerator14.Current;
				allItems.Add(current14);
				allItems.AddRange(current14.Objects);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_069a;
				}
				continue;
				end_IL_069a:
				break;
			}
		}
		IEnumerator<CellMarginsItem> enumerator15 = default(IEnumerator<CellMarginsItem>);
		try
		{
			enumerator15 = CellMargins.GetEnumerator();
			while (enumerator15.MoveNext())
			{
				BaseItem current15 = enumerator15.Current;
				allItems.Add(current15);
				allItems.AddRange(current15.Objects);
			}
		}
		finally
		{
			if (enumerator15 != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					enumerator15.Dispose();
					break;
				}
			}
		}
		allItems = null;
		SourceCollection = CollectionViewSource.GetDefaultView(AllItems);
		SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(AH.A(52438)));
		SourceCollection.Filter = A;
		lbxResults.SelectionChanged += NavigateObjects;
	}

	private DataTemplate A(string A)
	{
		DataTemplate result = default(DataTemplate);
		try
		{
			result = (DataTemplate)lbxResults.FindResource(A);
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private bool A(object A)
	{
		return true;
	}

	private void A()
	{
		FontColors = new ObservableCollection<FontColorItem>();
		if (!Conventions.UsedFontColors.Any())
		{
			return;
		}
		IEnumerator<V<int, IEnumerable<Tuple<int, IndexedObject>>>> enumerator = default(IEnumerator<V<int, IEnumerable<Tuple<int, IndexedObject>>>>);
		IEnumerator<Tuple<int, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<int, IndexedObject>>);
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
			List<FontColorItem> list = new List<FontColorItem>();
			DataTemplate template = A(AH.A(52451));
			try
			{
				List<Tuple<int, IndexedObject>> usedFontColors = Conventions.UsedFontColors;
				Func<Tuple<int, IndexedObject>, int> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1);
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
				Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> elementSelector = [SpecialName] (Tuple<int, IndexedObject> A) => A;
				Func<int, IEnumerable<Tuple<int, IndexedObject>>, V<int, IEnumerable<Tuple<int, IndexedObject>>>> resultSelector;
				if (_Closure_0024__.A == null)
				{
					resultSelector = (_Closure_0024__.A = [SpecialName] (int A, IEnumerable<Tuple<int, IndexedObject>> B) => new V<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B));
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
					resultSelector = _Closure_0024__.A;
				}
				IEnumerable<V<int, IEnumerable<Tuple<int, IndexedObject>>>> enumerable = usedFontColors.GroupBy(keySelector, elementSelector, resultSelector);
				int intTotal = enumerable.Count();
				try
				{
					enumerator = enumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						V<int, IEnumerable<Tuple<int, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list2 = new List<IndexedObject>();
						try
						{
							enumerator2 = current.IntegerGroup.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Tuple<int, IndexedObject> current2 = enumerator2.Current;
								list2.Add(current2.Item2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0145;
								}
								continue;
								end_IL_0145:
								break;
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									enumerator2.Dispose();
									break;
								}
							}
						}
						list.Add(new FontColorItem(current.Item1, list2, template, NavItemDataTemplate, A(current.Item1), intTotal));
						list2 = null;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_01a5;
						}
						continue;
						end_IL_01a5:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				enumerable = null;
				FontColors = new ObservableCollection<FontColorItem>(list.OrderByDescending([SpecialName] (FontColorItem A) => ((BaseItem)A).Quantity));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			list = null;
			template = null;
			return;
		}
	}

	private void B()
	{
		FillColors = new ObservableCollection<FillColorItem>();
		if (!Conventions.UsedFillColors.Any())
		{
			return;
		}
		DataTemplate template = A(AH.A(52486));
		try
		{
			List<Tuple<int, IndexedObject>> usedFillColors = Conventions.UsedFillColors;
			Func<Tuple<int, IndexedObject>, int> keySelector;
			if (_Closure_0024__.B == null)
			{
				keySelector = (_Closure_0024__.B = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				keySelector = _Closure_0024__.B;
			}
			Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> elementSelector = [SpecialName] (Tuple<int, IndexedObject> A) => A;
			Func<int, IEnumerable<Tuple<int, IndexedObject>>, W<int, IEnumerable<Tuple<int, IndexedObject>>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (int A, IEnumerable<Tuple<int, IndexedObject>> B) => new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B));
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
				resultSelector = _Closure_0024__.A;
			}
			IOrderedEnumerable<W<int, IEnumerable<Tuple<int, IndexedObject>>>> orderedEnumerable = from A in usedFillColors.GroupBy(keySelector, elementSelector, resultSelector)
				orderby A.g.Count() descending
				select A;
			int intTotal = orderedEnumerable.Count();
			IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>> enumerator = default(IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>>);
			try
			{
				enumerator = orderedEnumerable.GetEnumerator();
				IEnumerator<Tuple<int, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<int, IndexedObject>>);
				while (enumerator.MoveNext())
				{
					W<int, IEnumerable<Tuple<int, IndexedObject>>> current = enumerator.Current;
					List<IndexedObject> list = new List<IndexedObject>();
					try
					{
						enumerator2 = current.g.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Tuple<int, IndexedObject> current2 = enumerator2.Current;
							list.Add(current2.Item2);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0164;
							}
							continue;
							end_IL_0164:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
					FillColors.Add(new FillColorItem(current.Item1, list, template, NavItemDataTemplate, A(current.Item1), intTotal));
					list = null;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_01c8;
					}
					continue;
					end_IL_01c8:
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			orderedEnumerable = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		template = null;
	}

	private void C()
	{
		BorderColors = new ObservableCollection<BorderColorItem>();
		if (!Conventions.UsedBorderColors.Any())
		{
			return;
		}
		IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>> enumerator = default(IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>>);
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
			DataTemplate template = A(AH.A(52486));
			try
			{
				List<Tuple<int, IndexedObject>> usedBorderColors = Conventions.UsedBorderColors;
				Func<Tuple<int, IndexedObject>, int> keySelector = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1;
				Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> elementSelector;
				if (_Closure_0024__.C == null)
				{
					elementSelector = (_Closure_0024__.C = [SpecialName] (Tuple<int, IndexedObject> A) => A);
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
					elementSelector = _Closure_0024__.C;
				}
				IEnumerable<W<int, IEnumerable<Tuple<int, IndexedObject>>>> source = usedBorderColors.GroupBy(keySelector, elementSelector, [SpecialName] (int A, IEnumerable<Tuple<int, IndexedObject>> B) => new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B));
				Func<W<int, IEnumerable<Tuple<int, IndexedObject>>>, int> keySelector2;
				if (_Closure_0024__.B == null)
				{
					keySelector2 = (_Closure_0024__.B = [SpecialName] (W<int, IEnumerable<Tuple<int, IndexedObject>>> A) => A.g.Count());
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
					keySelector2 = _Closure_0024__.B;
				}
				IOrderedEnumerable<W<int, IEnumerable<Tuple<int, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
				int intTotal = orderedEnumerable.Count();
				try
				{
					enumerator = orderedEnumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						W<int, IEnumerable<Tuple<int, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list = new List<IndexedObject>();
						using (IEnumerator<Tuple<int, IndexedObject>> enumerator2 = current.g.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								Tuple<int, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0166;
								}
								continue;
								end_IL_0166:
								break;
							}
						}
						BorderColors.Add(new BorderColorItem(current.Item1, list, template, NavItemDataTemplate, A(current.Item1), intTotal));
						list = null;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_01c4;
						}
						continue;
						end_IL_01c4:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				orderedEnumerable = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			template = null;
			return;
		}
	}

	private void D()
	{
		FillTransparencies = new ObservableCollection<FillTransparencyItem>();
		FillTransparencyOptions = new ObservableCollection<float>();
		if (!Conventions.UsedFillTransparencies.Any())
		{
			return;
		}
		IEnumerator<W<float, IEnumerable<Tuple<float, IndexedObject>>>> enumerator = default(IEnumerator<W<float, IEnumerable<Tuple<float, IndexedObject>>>>);
		IEnumerator<Tuple<float, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<float, IndexedObject>>);
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
			DataTemplate template = A(AH.A(52513));
			int num = 0;
			try
			{
				List<Tuple<float, IndexedObject>> usedFillTransparencies = Conventions.UsedFillTransparencies;
				Func<Tuple<float, IndexedObject>, float> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<float, IndexedObject> A) => A.Item1);
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
					keySelector = _Closure_0024__.A;
				}
				Func<Tuple<float, IndexedObject>, Tuple<float, IndexedObject>> elementSelector;
				if (_Closure_0024__.A == null)
				{
					elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<float, IndexedObject> A) => A);
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
					elementSelector = _Closure_0024__.A;
				}
				IEnumerable<W<float, IEnumerable<Tuple<float, IndexedObject>>>> source = usedFillTransparencies.GroupBy(keySelector, elementSelector, [SpecialName] (float A, IEnumerable<Tuple<float, IndexedObject>> B) => new W<float, IEnumerable<Tuple<float, IndexedObject>>>(A, B));
				Func<W<float, IEnumerable<Tuple<float, IndexedObject>>>, int> keySelector2;
				if (_Closure_0024__.A == null)
				{
					keySelector2 = (_Closure_0024__.A = [SpecialName] (W<float, IEnumerable<Tuple<float, IndexedObject>>> A) => A.g.Count());
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
					keySelector2 = _Closure_0024__.A;
				}
				IOrderedEnumerable<W<float, IEnumerable<Tuple<float, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
				int intTotal = orderedEnumerable.Count();
				try
				{
					enumerator = orderedEnumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						W<float, IEnumerable<Tuple<float, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list = new List<IndexedObject>();
						try
						{
							enumerator2 = current.g.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Tuple<float, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									enumerator2.Dispose();
									break;
								}
							}
						}
						FillTransparencies.Add(new FillTransparencyItem(current.Item1, list, template, NavItemDataTemplate, intTotal, num));
						FillTransparencyOptions.Add(current.Item1);
						list = null;
						num = checked(num + 1);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_01f4;
						}
						continue;
						end_IL_01f4:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				orderedEnumerable = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			template = null;
			return;
		}
	}

	private void E()
	{
		FontFamilies = new ObservableCollection<FontFamilyItem>();
		if (!Conventions.UsedFontFamilies.Any())
		{
			return;
		}
		List<FontFamilyItem> list = new List<FontFamilyItem>();
		DataTemplate template = A(AH.A(52566));
		try
		{
			List<Tuple<string, IndexedObject>> usedFontFamilies = Conventions.UsedFontFamilies;
			Func<Tuple<string, IndexedObject>, string> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<string, IndexedObject> A) => A.Item1);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				keySelector = _Closure_0024__.A;
			}
			Func<Tuple<string, IndexedObject>, Tuple<string, IndexedObject>> elementSelector;
			if (_Closure_0024__.A == null)
			{
				elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<string, IndexedObject> A) => A);
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
				elementSelector = _Closure_0024__.A;
			}
			Func<string, IEnumerable<Tuple<string, IndexedObject>>, X<string, IEnumerable<Tuple<string, IndexedObject>>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (string A, IEnumerable<Tuple<string, IndexedObject>> B) => new X<string, IEnumerable<Tuple<string, IndexedObject>>>(A, B));
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
				resultSelector = _Closure_0024__.A;
			}
			IEnumerable<X<string, IEnumerable<Tuple<string, IndexedObject>>>> enumerable = usedFontFamilies.GroupBy(keySelector, elementSelector, resultSelector);
			int intTotal = enumerable.Count();
			IEnumerator<X<string, IEnumerable<Tuple<string, IndexedObject>>>> enumerator = default(IEnumerator<X<string, IEnumerable<Tuple<string, IndexedObject>>>>);
			try
			{
				enumerator = enumerable.GetEnumerator();
				IEnumerator<Tuple<string, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<string, IndexedObject>>);
				while (enumerator.MoveNext())
				{
					X<string, IEnumerable<Tuple<string, IndexedObject>>> current = enumerator.Current;
					List<IndexedObject> list2 = new List<IndexedObject>();
					try
					{
						enumerator2 = current.StringGroup.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Tuple<string, IndexedObject> current2 = enumerator2.Current;
							list2.Add(current2.Item2);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0147;
							}
							continue;
							end_IL_0147:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
					list.Add(new FontFamilyItem(current.Item1, list2, template, NavItemDataTemplate, intTotal));
					list2 = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_019a;
					}
					continue;
					end_IL_019a:
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			enumerable = null;
			FontFamilies = new ObservableCollection<FontFamilyItem>(list.OrderByDescending([SpecialName] (FontFamilyItem A) => ((BaseItem)A).Quantity));
			FontFamilyOptions = new ObservableCollection<string>(FontFamilies.Select([SpecialName] (FontFamilyItem A) => ((BaseItem)A).Label));
			IEnumerator<FontFamilyItem> enumerator3 = default(IEnumerator<FontFamilyItem>);
			try
			{
				enumerator3 = FontFamilies.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					FontFamilyItem current3 = enumerator3.Current;
					((BaseItem)current3).SelectedIndex = FontFamilyOptions.IndexOf(((BaseItem)current3).Label);
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0275;
					}
					continue;
					end_IL_0275:
					break;
				}
			}
			finally
			{
				if (enumerator3 != null)
				{
					while (true)
					{
						switch (2)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		template = null;
	}

	private void F()
	{
		FontStyles = new ObservableCollection<FontStyleItem>();
		if (!Conventions.UsedFontStyles.Any())
		{
			return;
		}
		IEnumerator<X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>> enumerator = default(IEnumerator<X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>>);
		IEnumerator<FontStyleItem> enumerator3 = default(IEnumerator<FontStyleItem>);
		FC fC = default(FC);
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
			List<FontStyleItem> list = new List<FontStyleItem>();
			DataTemplate template = A(AH.A(52607));
			try
			{
				List<Tuple<FontStyle, IndexedObject>> usedFontStyles = Conventions.UsedFontStyles;
				Func<Tuple<FontStyle, IndexedObject>, FontStyle> keySelector = [SpecialName] (Tuple<FontStyle, IndexedObject> A) => A.Item1;
				Func<Tuple<FontStyle, IndexedObject>, Tuple<FontStyle, IndexedObject>> elementSelector = [SpecialName] (Tuple<FontStyle, IndexedObject> A) => A;
				Func<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>, X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>> resultSelector;
				if (_Closure_0024__.A == null)
				{
					resultSelector = (_Closure_0024__.A = [SpecialName] (FontStyle A, IEnumerable<Tuple<FontStyle, IndexedObject>> B) => new X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>(A, B));
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
					resultSelector = _Closure_0024__.A;
				}
				IEnumerable<X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>>> enumerable = usedFontStyles.GroupBy(keySelector, elementSelector, resultSelector);
				int intTotal = enumerable.Count();
				try
				{
					enumerator = enumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						X<FontStyle, IEnumerable<Tuple<FontStyle, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list2 = new List<IndexedObject>();
						using (IEnumerator<Tuple<FontStyle, IndexedObject>> enumerator2 = current.StringGroup.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								Tuple<FontStyle, IndexedObject> current2 = enumerator2.Current;
								list2.Add(current2.Item2);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0143;
								}
								continue;
								end_IL_0143:
								break;
							}
						}
						list.Add(new FontStyleItem(current.Item1, list2, template, NavItemDataTemplate, intTotal));
						list2 = null;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_018c;
						}
						continue;
						end_IL_018c:
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
				enumerable = null;
				List<FontStyleItem> source = list;
				Func<FontStyleItem, int> keySelector2;
				if (_Closure_0024__.A == null)
				{
					keySelector2 = (_Closure_0024__.A = [SpecialName] (FontStyleItem A) => ((BaseItem)A).Quantity);
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
					keySelector2 = _Closure_0024__.A;
				}
				FontStyles = new ObservableCollection<FontStyleItem>(source.OrderByDescending(keySelector2));
				ObservableCollection<FontStyleItem> fontStyles = FontStyles;
				Func<FontStyleItem, FontStyleOption> selector;
				if (_Closure_0024__.A == null)
				{
					selector = (_Closure_0024__.A = [SpecialName] (FontStyleItem A) => new FontStyleOption(A.Style, ((BaseItem)A).Label));
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
					selector = _Closure_0024__.A;
				}
				FontStyleOptions = new ObservableCollection<FontStyleOption>(fontStyles.Select(selector));
				try
				{
					enumerator3 = FontStyles.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						fC = new FC(fC);
						fC.A = enumerator3.Current;
						((BaseItem)fC.A).SelectedIndex = FontStyleOptions.IndexOf(FontStyleOptions.First(fC.A));
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_02a1;
						}
						continue;
						end_IL_02a1:
						break;
					}
				}
				finally
				{
					if (enumerator3 != null)
					{
						while (true)
						{
							switch (3)
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
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			list = null;
			template = null;
			return;
		}
	}

	private void G()
	{
		TextDecorations = new ObservableCollection<TextDecorationItem>();
		if (!Conventions.UsedTextDecorations.Any())
		{
			return;
		}
		List<TextDecorationItem> list = new List<TextDecorationItem>();
		DataTemplate template = A(AH.A(52644));
		try
		{
			List<Tuple<TextDecoration, IndexedObject>> usedTextDecorations = Conventions.UsedTextDecorations;
			Func<Tuple<TextDecoration, IndexedObject>, TextDecoration> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<TextDecoration, IndexedObject> A) => A.Item1);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				keySelector = _Closure_0024__.A;
			}
			Func<Tuple<TextDecoration, IndexedObject>, Tuple<TextDecoration, IndexedObject>> elementSelector;
			if (_Closure_0024__.A == null)
			{
				elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<TextDecoration, IndexedObject> A) => A);
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
				elementSelector = _Closure_0024__.A;
			}
			Func<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>, X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>> resultSelector;
			if (_Closure_0024__.A == null)
			{
				resultSelector = (_Closure_0024__.A = [SpecialName] (TextDecoration A, IEnumerable<Tuple<TextDecoration, IndexedObject>> B) => new X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>(A, B));
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
				resultSelector = _Closure_0024__.A;
			}
			IEnumerable<X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>> enumerable = usedTextDecorations.GroupBy(keySelector, elementSelector, resultSelector);
			int intTotal = enumerable.Count();
			IEnumerator<X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>> enumerator = default(IEnumerator<X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>>>);
			try
			{
				enumerator = enumerable.GetEnumerator();
				IEnumerator<Tuple<TextDecoration, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<TextDecoration, IndexedObject>>);
				while (enumerator.MoveNext())
				{
					X<TextDecoration, IEnumerable<Tuple<TextDecoration, IndexedObject>>> current = enumerator.Current;
					List<IndexedObject> list2 = new List<IndexedObject>();
					try
					{
						enumerator2 = current.StringGroup.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Tuple<TextDecoration, IndexedObject> current2 = enumerator2.Current;
							list2.Add(current2.Item2);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0143;
							}
							continue;
							end_IL_0143:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
					list.Add(new TextDecorationItem(current.Item1, list2, template, NavItemDataTemplate, intTotal));
					list2 = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0196;
					}
					continue;
					end_IL_0196:
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
			enumerable = null;
			List<TextDecorationItem> source = list;
			Func<TextDecorationItem, int> keySelector2;
			if (_Closure_0024__.A == null)
			{
				keySelector2 = (_Closure_0024__.A = [SpecialName] (TextDecorationItem A) => ((BaseItem)A).Quantity);
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
				keySelector2 = _Closure_0024__.A;
			}
			TextDecorations = new ObservableCollection<TextDecorationItem>(source.OrderByDescending(keySelector2));
			ObservableCollection<TextDecorationItem> textDecorations = TextDecorations;
			Func<TextDecorationItem, TextDecorationOption> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (TextDecorationItem A) => new TextDecorationOption(A.Decoration, ((BaseItem)A).Label));
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
				selector = _Closure_0024__.A;
			}
			TextDecorationOptions = new ObservableCollection<TextDecorationOption>(textDecorations.Select(selector));
			TextDecoration decor = new TextDecoration
			{
				Bold = false,
				Italic = false,
				UnderlineStyle = MsoTextUnderlineType.msoNoUnderline
			};
			TextDecorationOptions.Insert(0, new TextDecorationOption(decor, AH.A(52691)));
			using IEnumerator<TextDecorationItem> enumerator3 = TextDecorations.GetEnumerator();
			GC gC = default(GC);
			while (enumerator3.MoveNext())
			{
				gC = new GC(gC);
				gC.A = enumerator3.Current;
				((BaseItem)gC.A).SelectedIndex = TextDecorationOptions.IndexOf(TextDecorationOptions.First(gC.A));
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_02e8;
				}
				continue;
				end_IL_02e8:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		template = null;
	}

	private void H()
	{
		BorderDashStyles = new ObservableCollection<BorderDashItem>();
		if (!Conventions.UsedBorderDashStyles.Any())
		{
			return;
		}
		IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>> enumerator = default(IEnumerator<W<int, IEnumerable<Tuple<int, IndexedObject>>>>);
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
			DataTemplate template = A(AH.A(52700));
			try
			{
				List<Tuple<int, IndexedObject>> usedBorderDashStyles = Conventions.UsedBorderDashStyles;
				Func<Tuple<int, IndexedObject>, int> keySelector;
				if (_Closure_0024__.D == null)
				{
					keySelector = (_Closure_0024__.D = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1);
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
					keySelector = _Closure_0024__.D;
				}
				Func<Tuple<int, IndexedObject>, Tuple<int, IndexedObject>> elementSelector;
				if (_Closure_0024__.D == null)
				{
					elementSelector = (_Closure_0024__.D = [SpecialName] (Tuple<int, IndexedObject> A) => A);
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
					elementSelector = _Closure_0024__.D;
				}
				Func<int, IEnumerable<Tuple<int, IndexedObject>>, W<int, IEnumerable<Tuple<int, IndexedObject>>>> resultSelector;
				if (_Closure_0024__.C == null)
				{
					resultSelector = (_Closure_0024__.C = [SpecialName] (int A, IEnumerable<Tuple<int, IndexedObject>> B) => new W<int, IEnumerable<Tuple<int, IndexedObject>>>(A, B));
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
					resultSelector = _Closure_0024__.C;
				}
				IEnumerable<W<int, IEnumerable<Tuple<int, IndexedObject>>>> source = usedBorderDashStyles.GroupBy(keySelector, elementSelector, resultSelector);
				Func<W<int, IEnumerable<Tuple<int, IndexedObject>>>, int> keySelector2;
				if (_Closure_0024__.C == null)
				{
					keySelector2 = (_Closure_0024__.C = [SpecialName] (W<int, IEnumerable<Tuple<int, IndexedObject>>> A) => A.g.Count());
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
					keySelector2 = _Closure_0024__.C;
				}
				IOrderedEnumerable<W<int, IEnumerable<Tuple<int, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
				int intTotal = orderedEnumerable.Count();
				try
				{
					enumerator = orderedEnumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						W<int, IEnumerable<Tuple<int, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list = new List<IndexedObject>();
						using (IEnumerator<Tuple<int, IndexedObject>> enumerator2 = current.g.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								Tuple<int, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0178;
								}
								continue;
								end_IL_0178:
								break;
							}
						}
						BorderDashStyles.Add(new BorderDashItem(current.Item1, list, template, NavItemDataTemplate, intTotal, A((MsoLineDashStyle)current.Item1)));
						list = null;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_01d6;
						}
						continue;
						end_IL_01d6:
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
				orderedEnumerable = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			template = null;
			return;
		}
	}

	private void I()
	{
		BorderWeights = new ObservableCollection<BorderWeightItem>();
		if (!Conventions.UsedBorderWeights.Any())
		{
			return;
		}
		IEnumerator<W<float, IEnumerable<Tuple<float, IndexedObject>>>> enumerator = default(IEnumerator<W<float, IEnumerable<Tuple<float, IndexedObject>>>>);
		IEnumerator<Tuple<float, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<float, IndexedObject>>);
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
			DataTemplate template = A(AH.A(52741));
			try
			{
				List<Tuple<float, IndexedObject>> usedBorderWeights = Conventions.UsedBorderWeights;
				Func<Tuple<float, IndexedObject>, float> keySelector;
				if (_Closure_0024__.B == null)
				{
					keySelector = (_Closure_0024__.B = [SpecialName] (Tuple<float, IndexedObject> A) => A.Item1);
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
					keySelector = _Closure_0024__.B;
				}
				Func<Tuple<float, IndexedObject>, Tuple<float, IndexedObject>> elementSelector;
				if (_Closure_0024__.B == null)
				{
					elementSelector = (_Closure_0024__.B = [SpecialName] (Tuple<float, IndexedObject> A) => A);
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
					elementSelector = _Closure_0024__.B;
				}
				IEnumerable<W<float, IEnumerable<Tuple<float, IndexedObject>>>> source = usedBorderWeights.GroupBy(keySelector, elementSelector, [SpecialName] (float A, IEnumerable<Tuple<float, IndexedObject>> B) => new W<float, IEnumerable<Tuple<float, IndexedObject>>>(A, B));
				Func<W<float, IEnumerable<Tuple<float, IndexedObject>>>, int> keySelector2;
				if (_Closure_0024__.B == null)
				{
					keySelector2 = (_Closure_0024__.B = [SpecialName] (W<float, IEnumerable<Tuple<float, IndexedObject>>> A) => A.g.Count());
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
					keySelector2 = _Closure_0024__.B;
				}
				IOrderedEnumerable<W<float, IEnumerable<Tuple<float, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
				int intTotal = orderedEnumerable.Count();
				try
				{
					enumerator = orderedEnumerable.GetEnumerator();
					while (enumerator.MoveNext())
					{
						W<float, IEnumerable<Tuple<float, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list = new List<IndexedObject>();
						try
						{
							enumerator2 = current.g.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Tuple<float, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0176;
								}
								continue;
								end_IL_0176:
								break;
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									enumerator2.Dispose();
									break;
								}
							}
						}
						BorderWeights.Add(new BorderWeightItem(current.Item1, list, template, NavItemDataTemplate, intTotal, A(current.Item1)));
						list = null;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				orderedEnumerable = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			template = null;
			return;
		}
	}

	private void J()
	{
		ParagraphSpacing = new ObservableCollection<ParagraphSpacingItem>();
		ParagraphSpacingOptions = new ObservableCollection<ParagraphSpacingOption>();
		if (!Conventions.UsedParagraphSpacing.Any())
		{
			return;
		}
		checked
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
				DataTemplate template = A(AH.A(52782));
				int num = 0;
				try
				{
					List<Tuple<LineSpacing, IndexedObject>> usedParagraphSpacing = Conventions.UsedParagraphSpacing;
					Func<Tuple<LineSpacing, IndexedObject>, LineSpacing> keySelector = [SpecialName] (Tuple<LineSpacing, IndexedObject> A) => A.Item1;
					Func<Tuple<LineSpacing, IndexedObject>, Tuple<LineSpacing, IndexedObject>> elementSelector;
					if (_Closure_0024__.A == null)
					{
						elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<LineSpacing, IndexedObject> A) => A);
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
						elementSelector = _Closure_0024__.A;
					}
					Func<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>, W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>> resultSelector;
					if (_Closure_0024__.A == null)
					{
						resultSelector = (_Closure_0024__.A = [SpecialName] (LineSpacing A, IEnumerable<Tuple<LineSpacing, IndexedObject>> B) => new W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>(A, B));
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
						resultSelector = _Closure_0024__.A;
					}
					IEnumerable<W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>> source = usedParagraphSpacing.GroupBy(keySelector, elementSelector, resultSelector);
					Func<W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>, int> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>> A) => A.g.Count());
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
						keySelector2 = _Closure_0024__.A;
					}
					IOrderedEnumerable<W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
					int intTotal = orderedEnumerable.Count();
					using (IEnumerator<W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>>> enumerator = orderedEnumerable.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							W<LineSpacing, IEnumerable<Tuple<LineSpacing, IndexedObject>>> current = enumerator.Current;
							List<IndexedObject> list = new List<IndexedObject>();
							using (IEnumerator<Tuple<LineSpacing, IndexedObject>> enumerator2 = current.g.GetEnumerator())
							{
								while (enumerator2.MoveNext())
								{
									Tuple<LineSpacing, IndexedObject> current2 = enumerator2.Current;
									list.Add(current2.Item2);
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0185;
									}
									continue;
									end_IL_0185:
									break;
								}
							}
							ObservableCollection<ParagraphSpacingItem> paragraphSpacing = ParagraphSpacing;
							paragraphSpacing.Add(new ParagraphSpacingItem(current.Item1, list, template, NavItemDataTemplate, intTotal, num));
							ParagraphSpacingItem paragraphSpacingItem = paragraphSpacing[paragraphSpacing.Count - 1];
							ParagraphSpacingOptions.Add(new ParagraphSpacingOption(paragraphSpacingItem.Spacing, ((BaseItem)paragraphSpacingItem).Label));
							paragraphSpacingItem = null;
							_ = null;
							list = null;
							num++;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0213;
							}
							continue;
							end_IL_0213:
							break;
						}
					}
					orderedEnumerable = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				template = null;
				return;
			}
		}
	}

	private void K()
	{
		Indents = new ObservableCollection<IndentItem>();
		IndentOptions = new ObservableCollection<IndentOption>();
		if (!Conventions.UsedIndents.Any())
		{
			return;
		}
		checked
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
				DataTemplate template = A(AH.A(52831));
				int num = 0;
				try
				{
					List<Tuple<Indent, IndexedObject>> usedIndents = Conventions.UsedIndents;
					Func<Tuple<Indent, IndexedObject>, Indent> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<Indent, IndexedObject> A) => A.Item1);
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
						keySelector = _Closure_0024__.A;
					}
					Func<Tuple<Indent, IndexedObject>, Tuple<Indent, IndexedObject>> elementSelector;
					if (_Closure_0024__.A == null)
					{
						elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<Indent, IndexedObject> A) => A);
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
						elementSelector = _Closure_0024__.A;
					}
					IEnumerable<W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>> source = usedIndents.GroupBy(keySelector, elementSelector, [SpecialName] (Indent A, IEnumerable<Tuple<Indent, IndexedObject>> B) => new W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>(A, B));
					Func<W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>, int> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>> A) => A.g.Count());
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
						keySelector2 = _Closure_0024__.A;
					}
					IOrderedEnumerable<W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
					int intTotal = orderedEnumerable.Count();
					foreach (W<Indent, IEnumerable<Tuple<Indent, IndexedObject>>> item in orderedEnumerable)
					{
						List<IndexedObject> list = new List<IndexedObject>();
						using (IEnumerator<Tuple<Indent, IndexedObject>> enumerator2 = item.g.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								Tuple<Indent, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_017f;
								}
								continue;
								end_IL_017f:
								break;
							}
						}
						ObservableCollection<IndentItem> indents = Indents;
						indents.Add(new IndentItem(item.Item1, list, template, NavItemDataTemplate, intTotal, num));
						IndentItem indentItem = indents[indents.Count - 1];
						IndentOptions.Add(new IndentOption(indentItem.Indent, ((BaseItem)indentItem).Label));
						indentItem = null;
						_ = null;
						list = null;
						num++;
					}
					orderedEnumerable = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				template = null;
				return;
			}
		}
	}

	private void L()
	{
		BulletStyles = new ObservableCollection<BulletStyleItem>();
		BulletStyleOptions = new ObservableCollection<BulletStyleOption>();
		if (!Conventions.UsedBulletStyles.Any())
		{
			return;
		}
		DataTemplate template = A(AH.A(52862));
		int num = 0;
		checked
		{
			try
			{
				List<Tuple<BulletStyle, IndexedObject>> usedBulletStyles = Conventions.UsedBulletStyles;
				Func<Tuple<BulletStyle, IndexedObject>, BulletStyle> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (Tuple<BulletStyle, IndexedObject> A) => A.Item1);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					keySelector = _Closure_0024__.A;
				}
				Func<Tuple<BulletStyle, IndexedObject>, Tuple<BulletStyle, IndexedObject>> elementSelector;
				if (_Closure_0024__.A == null)
				{
					elementSelector = (_Closure_0024__.A = [SpecialName] (Tuple<BulletStyle, IndexedObject> A) => A);
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
					elementSelector = _Closure_0024__.A;
				}
				Func<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>, W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>> resultSelector;
				if (_Closure_0024__.A == null)
				{
					resultSelector = (_Closure_0024__.A = [SpecialName] (BulletStyle A, IEnumerable<Tuple<BulletStyle, IndexedObject>> B) => new W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>(A, B));
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
					resultSelector = _Closure_0024__.A;
				}
				IEnumerable<W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>> source = usedBulletStyles.GroupBy(keySelector, elementSelector, resultSelector);
				Func<W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>, int> keySelector2;
				if (_Closure_0024__.A == null)
				{
					keySelector2 = (_Closure_0024__.A = [SpecialName] (W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>> A) => A.g.Count());
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
					keySelector2 = _Closure_0024__.A;
				}
				IOrderedEnumerable<W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
				int intTotal = orderedEnumerable.Count();
				using (IEnumerator<W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>>> enumerator = orderedEnumerable.GetEnumerator())
				{
					IEnumerator<Tuple<BulletStyle, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<BulletStyle, IndexedObject>>);
					while (enumerator.MoveNext())
					{
						W<BulletStyle, IEnumerable<Tuple<BulletStyle, IndexedObject>>> current = enumerator.Current;
						List<IndexedObject> list = new List<IndexedObject>();
						try
						{
							enumerator2 = current.g.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Tuple<BulletStyle, IndexedObject> current2 = enumerator2.Current;
								list.Add(current2.Item2);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_017b;
								}
								continue;
								end_IL_017b:
								break;
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									enumerator2.Dispose();
									break;
								}
							}
						}
						ObservableCollection<BulletStyleItem> bulletStyles = BulletStyles;
						bulletStyles.Add(new BulletStyleItem(current.Item1, list, template, NavItemDataTemplate, intTotal, num));
						BulletStyleItem bulletStyleItem = bulletStyles[bulletStyles.Count - 1];
						BulletStyleOptions.Add(new BulletStyleOption(bulletStyleItem.Style, ((BaseItem)bulletStyleItem).Label));
						bulletStyleItem = null;
						_ = null;
						list = null;
						num++;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0211;
						}
						continue;
						end_IL_0211:
						break;
					}
				}
				orderedEnumerable = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			template = null;
		}
	}

	private void M()
	{
		TextBoxMargins = new ObservableCollection<TextBoxMarginsItem>();
		TextBoxMarginsOptions = new ObservableCollection<MarginsOption>();
		if (!Conventions.UsedTextBoxMargins.Any())
		{
			return;
		}
		checked
		{
			IEnumerator<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> enumerator = default(IEnumerator<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>>);
			IEnumerator<Tuple<Margins, IndexedObject>> enumerator2 = default(IEnumerator<Tuple<Margins, IndexedObject>>);
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
				DataTemplate template = A(AH.A(52901));
				int num = 0;
				try
				{
					List<Tuple<Margins, IndexedObject>> usedTextBoxMargins = Conventions.UsedTextBoxMargins;
					Func<Tuple<Margins, IndexedObject>, Margins> keySelector = [SpecialName] (Tuple<Margins, IndexedObject> A) => A.Item1;
					Func<Tuple<Margins, IndexedObject>, Tuple<Margins, IndexedObject>> elementSelector = [SpecialName] (Tuple<Margins, IndexedObject> A) => A;
					Func<Margins, IEnumerable<Tuple<Margins, IndexedObject>>, W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> resultSelector;
					if (_Closure_0024__.A == null)
					{
						resultSelector = (_Closure_0024__.A = [SpecialName] (Margins A, IEnumerable<Tuple<Margins, IndexedObject>> B) => new W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>(A, B));
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
						resultSelector = _Closure_0024__.A;
					}
					IOrderedEnumerable<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> orderedEnumerable = from A in usedTextBoxMargins.GroupBy(keySelector, elementSelector, resultSelector)
						orderby A.g.Count() descending
						select A;
					int intTotal = orderedEnumerable.Count();
					try
					{
						enumerator = orderedEnumerable.GetEnumerator();
						while (enumerator.MoveNext())
						{
							W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> current = enumerator.Current;
							List<IndexedObject> list = new List<IndexedObject>();
							try
							{
								enumerator2 = current.g.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									Tuple<Margins, IndexedObject> current2 = enumerator2.Current;
									list.Add(current2.Item2);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_016b;
									}
									continue;
									end_IL_016b:
									break;
								}
							}
							finally
							{
								if (enumerator2 != null)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										enumerator2.Dispose();
										break;
									}
								}
							}
							ObservableCollection<TextBoxMarginsItem> textBoxMargins = TextBoxMargins;
							textBoxMargins.Add(new TextBoxMarginsItem(current.Item1, list, template, NavItemDataTemplate, intTotal, num, AH.A(52946)));
							TextBoxMarginsItem textBoxMarginsItem = textBoxMargins[textBoxMargins.Count - 1];
							TextBoxMarginsOptions.Add(new MarginsOption(textBoxMarginsItem.Margins, ((BaseItem)textBoxMarginsItem).Label));
							textBoxMarginsItem = null;
							_ = null;
							list = null;
							num++;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_020f;
							}
							continue;
							end_IL_020f:
							break;
						}
					}
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								enumerator.Dispose();
								break;
							}
						}
					}
					orderedEnumerable = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Interaction.MsgBox(ex2.Message);
					ProjectData.ClearProjectError();
				}
				template = null;
				return;
			}
		}
	}

	private void N()
	{
		CellMargins = new ObservableCollection<CellMarginsItem>();
		CellMarginsOptions = new ObservableCollection<MarginsOption>();
		if (!Conventions.UsedCellMargins.Any())
		{
			return;
		}
		checked
		{
			IEnumerator<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> enumerator = default(IEnumerator<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>>);
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
				DataTemplate template = A(AH.A(52979));
				int num = 0;
				try
				{
					List<Tuple<Margins, IndexedObject>> usedCellMargins = Conventions.UsedCellMargins;
					Func<Tuple<Margins, IndexedObject>, Margins> keySelector = [SpecialName] (Tuple<Margins, IndexedObject> A) => A.Item1;
					Func<Tuple<Margins, IndexedObject>, Tuple<Margins, IndexedObject>> elementSelector;
					if (_Closure_0024__.B == null)
					{
						elementSelector = (_Closure_0024__.B = [SpecialName] (Tuple<Margins, IndexedObject> A) => A);
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
						elementSelector = _Closure_0024__.B;
					}
					Func<Margins, IEnumerable<Tuple<Margins, IndexedObject>>, W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> resultSelector;
					if (_Closure_0024__.B == null)
					{
						resultSelector = (_Closure_0024__.B = [SpecialName] (Margins A, IEnumerable<Tuple<Margins, IndexedObject>> B) => new W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>(A, B));
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
						resultSelector = _Closure_0024__.B;
					}
					IEnumerable<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> source = usedCellMargins.GroupBy(keySelector, elementSelector, resultSelector);
					Func<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>, int> keySelector2;
					if (_Closure_0024__.B == null)
					{
						keySelector2 = (_Closure_0024__.B = [SpecialName] (W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> A) => A.g.Count());
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
					IOrderedEnumerable<W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>>> orderedEnumerable = source.OrderByDescending(keySelector2);
					int intTotal = orderedEnumerable.Count();
					try
					{
						enumerator = orderedEnumerable.GetEnumerator();
						while (enumerator.MoveNext())
						{
							W<Margins, IEnumerable<Tuple<Margins, IndexedObject>>> current = enumerator.Current;
							List<IndexedObject> list = new List<IndexedObject>();
							using (IEnumerator<Tuple<Margins, IndexedObject>> enumerator2 = current.g.GetEnumerator())
							{
								while (enumerator2.MoveNext())
								{
									Tuple<Margins, IndexedObject> current2 = enumerator2.Current;
									list.Add(current2.Item2);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0189;
									}
									continue;
									end_IL_0189:
									break;
								}
							}
							ObservableCollection<CellMarginsItem> cellMargins = CellMargins;
							cellMargins.Add(new CellMarginsItem(current.Item1, list, template, NavItemDataTemplate, intTotal, num, AH.A(53018)));
							CellMarginsItem cellMarginsItem = cellMargins[cellMargins.Count - 1];
							CellMarginsOptions.Add(new MarginsOption(cellMarginsItem.Margins, ((BaseItem)cellMarginsItem).Label));
							cellMarginsItem = null;
							_ = null;
							list = null;
							num++;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0221;
							}
							continue;
							end_IL_0221:
							break;
						}
					}
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								enumerator.Dispose();
								break;
							}
						}
					}
					orderedEnumerable = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				template = null;
				return;
			}
		}
	}

	private int A(int A)
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0096: Expected O, but got Unknown
		int num = ColorOptions.IndexOf(ColorOptions.FirstOrDefault([SpecialName] (PaletteColor val) => val.OLE == A));
		if (num == -1)
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
			System.Drawing.Color color = ColorTranslator.FromOle(A);
			ColorOptions.Add(new PaletteColor(System.Windows.Media.Color.FromRgb(color.R, color.G, color.B), true, false, color.Name));
			num = checked(ColorOptions.Count - 1);
		}
		return num;
	}

	private int A(MsoLineDashStyle A)
	{
		ObservableCollection<LineDashStyle> lineDashStyleOptions = LineDashStyleOptions;
		return lineDashStyleOptions.IndexOf(lineDashStyleOptions.First([SpecialName] (LineDashStyle val) => val.Style == A));
	}

	private int A(float A)
	{
		ObservableCollection<LineWeight> lineWeightOptions = LineWeightOptions;
		return lineWeightOptions.IndexOf(lineWeightOptions.First([SpecialName] (LineWeight val) => val.Weight == A));
	}

	public void ExpandItems(object sender, RoutedEventArgs e)
	{
		A(RuntimeHelpers.GetObjectValue(sender), B: true);
	}

	public void CollapseItems(object sender, RoutedEventArgs e)
	{
		A(RuntimeHelpers.GetObjectValue(sender), B: false);
	}

	private void A(object A, bool B)
	{
		if (!((A as ToggleButton)?.DataContext is BaseItem baseItem))
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
			foreach (NavigationItem @object in baseItem.Objects)
			{
				((BaseItem)@object).IsVisible = B;
			}
			BaseItem baseItem2 = null;
			return;
		}
	}

	private void lbxResults_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Escape)
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
					if (key != Key.Space)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								if ((uint)(key - 23) <= 3u)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											if (!e.IsRepeat)
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														break;
													default:
														lbxResults.KeyUp += NavKeyUp;
														return;
													}
												}
											}
											return;
										}
									}
								}
								return;
							}
						}
					}
					if (lbxResults.SelectedIndex > -1)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								_ = lbxResults.SelectedItem is NavigationItem;
								return;
							}
						}
					}
					return;
				}
			}
		}
		P();
		e.Handled = true;
	}

	private void NavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		lbxResults.KeyUp -= NavKeyUp;
		O();
		e.Handled = true;
	}

	private void NavigateObjects(object sender, SelectionChangedEventArgs e)
	{
		O();
	}

	private void O()
	{
		//IL_0235: Unknown result type (might be due to invalid IL or missing references)
		//IL_023f: Expected O, but got Unknown
		P();
		int D = 0;
		if (lbxResults.SelectedIndex > -1)
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
			if (!Keyboard.IsKeyDown(Key.Down) && !Keyboard.IsKeyDown(Key.Up))
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
				if (!Keyboard.IsKeyDown(Key.Right))
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
					if (!Keyboard.IsKeyDown(Key.Left))
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
						if (!Keyboard.IsKeyDown(Key.Next))
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
							if (!Keyboard.IsKeyDown(Key.Prior))
							{
								List<Rect> A = new List<Rect>();
								DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
								this.m_A = true;
								if (lbxResults.SelectedItem is NavigationItem)
								{
									NavigationItem navigationItem = (NavigationItem)lbxResults.SelectedItem;
									if (navigationItem.IndexedObject.SlideOrLayout is Slide)
									{
										activeWindow.View.GotoSlide(((Slide)navigationItem.IndexedObject.SlideOrLayout).SlideIndex);
										System.Windows.Forms.Application.DoEvents();
										this.A(navigationItem, ref A, ref D);
									}
									navigationItem = null;
								}
								else
								{
									List<NavigationItem> objects = ((BaseItem)lbxResults.SelectedItem).Objects;
									this.A(ref A, objects, activeWindow.Selection.SlideRange[1].SlideIndex, ref D);
									if (!A.Any())
									{
										if (objects == null)
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
										else if (objects.Count > 0)
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
											IndexedObject indexedObject = objects[0].IndexedObject;
											if (indexedObject.SlideOrLayout is Slide)
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
												int slideIndex = ((Slide)indexedObject.SlideOrLayout).SlideIndex;
												activeWindow.View.GotoSlide(slideIndex);
												System.Windows.Forms.Application.DoEvents();
												this.A(ref A, objects, slideIndex, ref D);
											}
											indexedObject = null;
										}
									}
									JG.A(objects);
									objects = null;
								}
								if (A.Any())
								{
									MarchingAnts = new wpfMarchingAnts(A);
									wpfMarchingAnts marchingAnts = MarchingAnts;
									((Window)(object)marchingAnts).Owner = Window.GetWindow(this);
									((Window)(object)marchingAnts).Left = A.OrderBy([SpecialName] (Rect rect) => rect.Left).ToList()[0].Left;
									((Window)(object)marchingAnts).Top = A.OrderBy([SpecialName] (Rect rect) => rect.Top).ToList()[0].Top;
									((Window)(object)marchingAnts).ShowActivated = false;
									((Window)(object)marchingAnts).Show();
									_ = null;
								}
								this.m_A = false;
								A = null;
								activeWindow = null;
							}
						}
					}
				}
			}
		}
		PowerPointAddIn1.DeckCheck.UI.Pane.A(D);
	}

	private void A(ref List<Rect> A, List<NavigationItem> B, int C, ref int D)
	{
		List<NavigationItem> list = this.A(B, C);
		using (List<NavigationItem>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				NavigationItem current = enumerator.Current;
				this.A(current, ref A, ref D);
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
				break;
			}
		}
		JG.A(list);
		list = null;
	}

	private void P()
	{
		if (MarchingAnts == null)
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
			MarchingAnts.CloseByCode = true;
			((Window)(object)MarchingAnts).Close();
			MarchingAnts = null;
			return;
		}
	}

	private List<NavigationItem> A(List<NavigationItem> A, int B)
	{
		return A.Where([SpecialName] (NavigationItem navigationItem) => navigationItem.IndexedObject.SlideOrLayout is Slide && ((Slide)navigationItem.IndexedObject.SlideOrLayout).SlideIndex == B).ToList();
	}

	private void A(NavigationItem A, ref List<Rect> B, ref int C)
	{
		LC a = default(LC);
		LC CS_0024_003C_003E8__locals147 = new LC(a);
		CS_0024_003C_003E8__locals147.A = A.IndexedObject.Shape;
		CS_0024_003C_003E8__locals147.A = RuntimeHelpers.GetObjectValue(A.IndexedObject.Child);
		CS_0024_003C_003E8__locals147.A = 0f;
		CS_0024_003C_003E8__locals147.B = 0f;
		checked
		{
			try
			{
				if (CS_0024_003C_003E8__locals147.A is TextRange2)
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
						if (PowerPointAddIn1.DeckCheck.UI.MarchingAnts.UseRelativePosition(CS_0024_003C_003E8__locals147.A))
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
							CS_0024_003C_003E8__locals147.A = CS_0024_003C_003E8__locals147.A.Left;
							CS_0024_003C_003E8__locals147.B = CS_0024_003C_003E8__locals147.A.Top;
						}
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetTextRangeRectangle((TextRange2)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
						break;
					}
				}
				else if (CS_0024_003C_003E8__locals147.A is Microsoft.Office.Interop.PowerPoint.Shape)
				{
					this.A(ref B, [SpecialName] () =>
					{
						Type typeFromHandle = typeof(MarchingAnts);
						string memberName = AH.A(170589);
						object[] obj = new object[1] { CS_0024_003C_003E8__locals147.A };
						object[] array = obj;
						bool[] obj2 = new bool[1] { true };
						bool[] array2 = obj2;
						object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							CS_0024_003C_003E8__locals147.A = RuntimeHelpers.GetObjectValue(array[0]);
						}
						if (obj3 == null)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									return default(Rect);
								}
							}
						}
						return (Rect)obj3;
					}, ref C);
				}
				else if (CS_0024_003C_003E8__locals147.A is Cell)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(((Cell)CS_0024_003C_003E8__locals147.A).Shape), ref C);
						break;
					}
				}
				else if (CS_0024_003C_003E8__locals147.A is BulletFormat2)
				{
					NC nC = default(NC);
					MC a2 = default(MC);
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						nC = new NC(nC);
						nC.A = CS_0024_003C_003E8__locals147;
						if (PowerPointAddIn1.DeckCheck.UI.MarchingAnts.UseRelativePosition(nC.A.A))
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
							nC.A.A = nC.A.A.Left;
							nC.A.B = nC.A.A.Top;
						}
						ParagraphFormat2 paragraphFormat = (ParagraphFormat2)((BulletFormat2)nC.A.A).Parent;
						nC.A = paragraphFormat.LeftIndent;
						MC CS_0024_003C_003E8__locals107 = new MC(a2);
						CS_0024_003C_003E8__locals107.A = nC;
						CS_0024_003C_003E8__locals107.A = (TextRange2)paragraphFormat.Parent;
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetObjectRectangle(CS_0024_003C_003E8__locals107.A.A.A + CS_0024_003C_003E8__locals107.A.BoundLeft - CS_0024_003C_003E8__locals107.A.A, CS_0024_003C_003E8__locals107.A.A.B + CS_0024_003C_003E8__locals107.A.BoundTop, CS_0024_003C_003E8__locals107.A.A, CS_0024_003C_003E8__locals107.A.BoundHeight), ref C);
						paragraphFormat = null;
						break;
					}
				}
				else if (CS_0024_003C_003E8__locals147.A is ChartArea)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(CS_0024_003C_003E8__locals147.A), ref C);
						break;
					}
				}
				else if (CS_0024_003C_003E8__locals147.A is PlotArea)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
						CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle((PlotArea)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
						break;
					}
				}
				else
				{
					if (CS_0024_003C_003E8__locals147.A is Gridlines)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is HiLoLines)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is DropLines)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is UpBars)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is DownBars)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is IMsoErrorBars)
					{
						goto IL_0338;
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
					if (CS_0024_003C_003E8__locals147.A is IMsoLeaderLines)
					{
						goto IL_0338;
					}
					if (CS_0024_003C_003E8__locals147.A is IMsoTrendline)
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
						goto IL_0338;
					}
					if (CS_0024_003C_003E8__locals147.A is IMsoDataLabel)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLabelRectangle((IMsoDataLabel)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is IMsoDataLabels)
					{
						IEnumerator enumerator = default(IEnumerator);
						OC oC = default(OC);
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							enumerator = ((IEnumerable)CS_0024_003C_003E8__locals147.A).GetEnumerator();
							try
							{
								while (enumerator.MoveNext())
								{
									oC = new OC(oC);
									oC.A = CS_0024_003C_003E8__locals147;
									oC.A = (IMsoDataLabel)enumerator.Current;
									this.A(ref B, oC.A, ref C);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0469;
									}
									continue;
									end_IL_0469:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is IMsoSeries)
					{
						CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
						CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle(CS_0024_003C_003E8__locals147.A.Chart.PlotArea, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
					}
					else if (CS_0024_003C_003E8__locals147.A is ChartTitle)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetChartTitleRectangle((ChartTitle)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is Axis)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisRectangle((Axis)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is AxisTitle)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisTitleRectangle((AxisTitle)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is TickLabels)
					{
						PC a3 = default(PC);
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							PC CS_0024_003C_003E8__locals131 = new PC(a3);
							CS_0024_003C_003E8__locals131.A = CS_0024_003C_003E8__locals147;
							CS_0024_003C_003E8__locals131.A = ((TickLabels)CS_0024_003C_003E8__locals131.A.A).Parent as Axis;
							if (CS_0024_003C_003E8__locals131.A != null)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									CS_0024_003C_003E8__locals131.A.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals131.A.A);
									CS_0024_003C_003E8__locals131.A.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals131.A.A);
									this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetAxisRectangle(CS_0024_003C_003E8__locals131.A, CS_0024_003C_003E8__locals131.A.A, CS_0024_003C_003E8__locals131.A.B), ref C);
									break;
								}
							}
							else if (A.IndexedObject.AreRadarLabels)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaOuterRectangle(CS_0024_003C_003E8__locals131.A.A), ref C);
									break;
								}
							}
							else
							{
								this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(CS_0024_003C_003E8__locals131.A.A), ref C);
							}
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is Legend)
					{
						CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
						CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
						this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendRectangle((Legend)CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
					}
					else if (CS_0024_003C_003E8__locals147.A is Microsoft.Office.Core.LegendEntry)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
							CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendRectangle(CS_0024_003C_003E8__locals147.A.Chart.Legend, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is DataTable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetShapeRectangle(CS_0024_003C_003E8__locals147.A), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is Microsoft.Office.Core.Shape)
					{
						this.A(ref B, [SpecialName] () =>
						{
							Type typeFromHandle = typeof(MarchingAnts);
							string memberName = AH.A(170589);
							object[] obj = new object[1] { CS_0024_003C_003E8__locals147.A };
							object[] array = obj;
							bool[] obj2 = new bool[1] { true };
							bool[] array2 = obj2;
							object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
							if (array2[0])
							{
								CS_0024_003C_003E8__locals147.A = RuntimeHelpers.GetObjectValue(array[0]);
							}
							if (obj3 == null)
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
										return default(Rect);
									}
								}
							}
							return (Rect)obj3;
						}, ref C);
					}
					else if (CS_0024_003C_003E8__locals147.A is IMsoLegendKey)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetLegendKeyRectangle(CS_0024_003C_003E8__locals147.A, (IMsoLegendKey)CS_0024_003C_003E8__locals147.A), ref C);
							break;
						}
					}
					else if (CS_0024_003C_003E8__locals147.A is ChartPoint)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetChartPointRectangle(CS_0024_003C_003E8__locals147.A, (ChartPoint)CS_0024_003C_003E8__locals147.A), ref C);
							break;
						}
					}
				}
				goto end_IL_004c;
				IL_0338:
				CS_0024_003C_003E8__locals147.A = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals147.A);
				CS_0024_003C_003E8__locals147.B = PowerPointAddIn1.DeckCheck.UI.MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals147.A);
				this.A(ref B, [SpecialName] () => PowerPointAddIn1.DeckCheck.UI.MarchingAnts.GetPlotAreaRectangle(CS_0024_003C_003E8__locals147.A.Chart.PlotArea, CS_0024_003C_003E8__locals147.A, CS_0024_003C_003E8__locals147.B), ref C);
				end_IL_004c:;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				C++;
				ProjectData.ClearProjectError();
			}
			CS_0024_003C_003E8__locals147.A = null;
			CS_0024_003C_003E8__locals147.A = null;
		}
	}

	private void A(ref List<Rect> A, Func<Rect> B, ref int C)
	{
		checked
		{
			try
			{
				A.Add(B());
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				C++;
				ProjectData.ClearProjectError();
			}
		}
	}

	private void ListBoxLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
	{
		if (!this.m_A)
		{
			lbxResults.SelectedIndex = -1;
		}
	}

	private void GrowBorder(object sender, System.Windows.Input.MouseEventArgs e)
	{
		((Border)sender).BorderBrush = ((Border)sender).Background;
	}

	private void ShrinkBorder(object sender, System.Windows.Input.MouseEventArgs e)
	{
		((Border)sender).BorderBrush = new SolidColorBrush(System.Windows.Media.Colors.White);
	}

	private void ColorChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
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
			System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
			if (comboBox.IsLoaded)
			{
				int oLE = ColorOptions[comboBox.SelectedIndex].OLE;
				NG.A.Application.StartNewUndoEntry();
				System.Windows.Controls.ComboBox comboBox2 = comboBox;
				if (comboBox2.DataContext is FillColorItem)
				{
					A((FillColorItem)comboBox2.DataContext, oLE);
				}
				else if (comboBox2.DataContext is FontColorItem)
				{
					A((FontColorItem)comboBox2.DataContext, oLE);
				}
				else
				{
					A((BorderColorItem)comboBox2.DataContext, oLE);
				}
				comboBox2 = null;
			}
			comboBox = null;
			return;
		}
	}

	private void A(int A, int B, ref List<Tuple<int, IndexedObject>> C, Action D)
	{
		List<Tuple<int, IndexedObject>> list = C;
		checked
		{
			int num = list.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				if (list[i].Item1 == A)
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
					List<Tuple<int, IndexedObject>> list2 = C;
					list2[i] = new Tuple<int, IndexedObject>(B, list2[i].Item2);
					list2 = null;
				}
				_ = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				list = null;
				this.m_B = true;
				D();
				Q();
				this.m_B = false;
				return;
			}
		}
	}

	private void Q()
	{
		Conventions.DeterminePaletteUsage();
		checked
		{
			int num = ((Conventions)Conventions).ColorPalette.Count - 1;
			QC qC = default(QC);
			for (int i = 0; i <= num; i++)
			{
				qC = new QC(qC);
				qC.A = this;
				qC.A = i;
				ColorOptions[i].IsUsed = ((Conventions)Conventions).ColorPalette.First(qC.A).IsUsed;
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
	}

	private void A(FontColorItem A, int B)
	{
		int num = ColorTranslator.ToOle(A.Color);
		List<string> listErrors = new List<string>();
		S();
		A.Reformat(B, num, ref listErrors);
		if (!listErrors.Any())
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
			if (this.A())
			{
				Conventions conventions;
				List<Tuple<int, IndexedObject>> C = (conventions = Conventions).UsedFontColors;
				this.A(num, B, ref C, this.A);
				conventions.UsedFontColors = C;
				this.A(FontColors);
			}
		}
		else
		{
			this.A(listErrors);
		}
		T();
		listErrors = null;
	}

	private void A(FillColorItem A, int B)
	{
		int num = ColorTranslator.ToOle(A.Color);
		List<string> listErrors = new List<string>();
		S();
		A.Reformat(B, num, ref listErrors);
		if (!listErrors.Any())
		{
			if (this.A())
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
				Conventions conventions;
				List<Tuple<int, IndexedObject>> C = (conventions = Conventions).UsedFillColors;
				this.A(num, B, ref C, this.B);
				conventions.UsedFillColors = C;
				this.A(FillColors);
			}
		}
		else
		{
			this.A(listErrors);
		}
		T();
		listErrors = null;
	}

	private void FillTransparencyChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					FillTransparencyItem obj = (FillTransparencyItem)comboBox.DataContext;
					float transparency = obj.Transparency;
					float num = FillTransparencyOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(num, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
						{
							List<Tuple<float, IndexedObject>> usedFillTransparencies = Conventions.UsedFillTransparencies;
							int num2 = usedFillTransparencies.Count - 1;
							for (int i = 0; i <= num2; i++)
							{
								Tuple<float, IndexedObject> tuple = usedFillTransparencies[i];
								if (tuple.Item1 == transparency)
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
									Conventions.UsedFillTransparencies[i] = new Tuple<float, IndexedObject>(num, tuple.Item2);
								}
								tuple = null;
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
							usedFillTransparencies = null;
							this.m_B = true;
							D();
							this.m_B = false;
							A(FillTransparencies);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void A(BorderColorItem A, int B)
	{
		int num = ColorTranslator.ToOle(A.Color);
		List<string> listErrors = new List<string>();
		S();
		A.Reformat(B, num, ref listErrors);
		if (!listErrors.Any())
		{
			if (this.A())
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
				Conventions conventions;
				List<Tuple<int, IndexedObject>> C = (conventions = Conventions).UsedBorderColors;
				this.A(num, B, ref C, this.C);
				conventions.UsedBorderColors = C;
				this.A(BorderColors);
			}
		}
		else
		{
			this.A(listErrors);
		}
		T();
		listErrors = null;
	}

	private void BorderDashStyleChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
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
			System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
			if (comboBox.IsLoaded)
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
				BorderDashItem obj = (BorderDashItem)comboBox.DataContext;
				LineDashStyle val = LineDashStyleOptions[comboBox.SelectedIndex];
				MsoLineDashStyle style = obj.Style;
				MsoLineDashStyle style2 = val.Style;
				List<string> listErrors = new List<string>();
				S();
				obj.Reformat(style2, ref listErrors);
				if (!listErrors.Any())
				{
					if (A())
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
						List<Tuple<int, IndexedObject>> usedBorderDashStyles = Conventions.UsedBorderDashStyles;
						int num = checked(usedBorderDashStyles.Count - 1);
						for (int i = 0; i <= num; i = checked(i + 1))
						{
							Tuple<int, IndexedObject> tuple = usedBorderDashStyles[i];
							if (tuple.Item1 == (int)style)
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
								Conventions.UsedBorderDashStyles[i] = new Tuple<int, IndexedObject>((int)style2, tuple.Item2);
							}
							tuple = null;
						}
						usedBorderDashStyles = null;
						this.m_B = true;
						H();
						this.m_B = false;
						A(BorderDashStyles);
					}
				}
				else
				{
					A(listErrors);
				}
				T();
				listErrors = null;
				val = null;
			}
			comboBox = null;
			return;
		}
	}

	private void BorderWeightChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					BorderWeightItem obj = (BorderWeightItem)comboBox.DataContext;
					LineWeight val = LineWeightOptions[comboBox.SelectedIndex];
					float weight = obj.Weight;
					float weight2 = val.Weight;
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(weight2, ref listErrors);
					if (!listErrors.Any())
					{
						if (A())
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
							List<Tuple<float, IndexedObject>> usedBorderWeights = Conventions.UsedBorderWeights;
							int num = usedBorderWeights.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<float, IndexedObject> tuple = usedBorderWeights[i];
								if (tuple.Item1 == weight)
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
									Conventions.UsedBorderWeights[i] = new Tuple<float, IndexedObject>(weight2, tuple.Item2);
								}
								tuple = null;
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
							usedBorderWeights = null;
							this.m_B = true;
							I();
							this.m_B = false;
							A(BorderWeights);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
					val = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void FontFamilyChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					FontFamilyItem obj = (FontFamilyItem)comboBox.DataContext;
					string family = obj.Family;
					string text = FontFamilyOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(text, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
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
							R(family, text);
							A(FontFamilies);
							List<Tuple<FontStyle, IndexedObject>> usedFontStyles = Conventions.UsedFontStyles;
							int num = usedFontStyles.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<FontStyle, IndexedObject> tuple = usedFontStyles[i];
								if (Operators.CompareString(tuple.Item1.Family, family, TextCompare: false) == 0)
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
									FontStyle item = new FontStyle
									{
										Family = text,
										Size = tuple.Item1.Size
									};
									Conventions.UsedFontStyles[i] = new Tuple<FontStyle, IndexedObject>(item, tuple.Item2);
								}
								tuple = null;
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
							usedFontStyles = null;
							this.m_B = true;
							F();
							this.m_B = false;
							A(FontStyles);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void FontStyleChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					FontStyleItem obj = (FontStyleItem)comboBox.DataContext;
					int selectedIndex = comboBox.SelectedIndex;
					string family = obj.Style.Family;
					float size = obj.Style.Size;
					string family2 = FontStyleOptions[selectedIndex].Style.Family;
					float size2 = FontStyleOptions[selectedIndex].Style.Size;
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(family2, size2, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
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
							R(family, family2);
							A(FontFamilies);
							List<Tuple<FontStyle, IndexedObject>> usedFontStyles = Conventions.UsedFontStyles;
							int num = usedFontStyles.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<FontStyle, IndexedObject> tuple = usedFontStyles[i];
								if (Operators.CompareString(tuple.Item1.Family, family, TextCompare: false) == 0 && tuple.Item1.Size == size)
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
									FontStyle item = new FontStyle
									{
										Family = family2,
										Size = size2
									};
									Conventions.UsedFontStyles[i] = new Tuple<FontStyle, IndexedObject>(item, tuple.Item2);
								}
								tuple = null;
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
							usedFontStyles = null;
							this.m_B = true;
							F();
							this.m_B = false;
							A(FontStyles);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void R(string A, string B)
	{
		List<Tuple<string, IndexedObject>> usedFontFamilies = Conventions.UsedFontFamilies;
		checked
		{
			int num = usedFontFamilies.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				Tuple<string, IndexedObject> tuple = usedFontFamilies[i];
				if (Operators.CompareString(tuple.Item1, A, TextCompare: false) == 0)
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
					Conventions.UsedFontFamilies[i] = new Tuple<string, IndexedObject>(B, tuple.Item2);
				}
				tuple = null;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				usedFontFamilies = null;
				this.m_B = true;
				E();
				this.m_B = false;
				return;
			}
		}
	}

	private void TextDecorationChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		checked
		{
			if (comboBox.IsLoaded)
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
				TextDecorationItem obj = (TextDecorationItem)comboBox.DataContext;
				int selectedIndex = comboBox.SelectedIndex;
				TextDecoration decoration = obj.Decoration;
				TextDecorationOption textDecorationOption = TextDecorationOptions[selectedIndex];
				List<string> listErrors = new List<string>();
				S();
				obj.Reformat(textDecorationOption, ref listErrors);
				if (!listErrors.Any())
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
					if (A())
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
						List<Tuple<TextDecoration, IndexedObject>> usedTextDecorations = Conventions.UsedTextDecorations;
						int num = usedTextDecorations.Count - 1;
						for (int i = 0; i <= num; i++)
						{
							Tuple<TextDecoration, IndexedObject> tuple = usedTextDecorations[i];
							if (tuple.Item1.Bold == decoration.Bold && tuple.Item1.Italic == decoration.Italic)
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
								if (tuple.Item1.UnderlineStyle == decoration.UnderlineStyle)
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
									TextDecoration item = default(TextDecoration);
									TextDecoration decoration2 = textDecorationOption.Decoration;
									item.Bold = decoration2.Bold;
									item.Italic = decoration2.Italic;
									item.UnderlineStyle = decoration2.UnderlineStyle;
									Conventions.UsedTextDecorations[i] = new Tuple<TextDecoration, IndexedObject>(item, tuple.Item2);
								}
							}
							tuple = null;
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
						usedTextDecorations = null;
						this.m_B = true;
						G();
						this.m_B = false;
						A(TextDecorations);
					}
				}
				else
				{
					A(listErrors);
				}
				T();
				listErrors = null;
			}
			comboBox = null;
		}
	}

	private void ParagraphSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					ParagraphSpacingItem obj = (ParagraphSpacingItem)comboBox.DataContext;
					LineSpacing spacing = obj.Spacing;
					ParagraphSpacingOption paragraphSpacingOption = ParagraphSpacingOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(paragraphSpacingOption, ref listErrors);
					if (!listErrors.Any())
					{
						if (A())
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
							List<Tuple<LineSpacing, IndexedObject>> usedParagraphSpacing = Conventions.UsedParagraphSpacing;
							int num = usedParagraphSpacing.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<LineSpacing, IndexedObject> tuple = usedParagraphSpacing[i];
								if (tuple.Item1.Before == spacing.Before && tuple.Item1.After == spacing.After)
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
									if (tuple.Item1.Within == spacing.Within)
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
										if (tuple.Item1.LineRuleWithin == spacing.LineRuleWithin)
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
											LineSpacing item = default(LineSpacing);
											LineSpacing spacing2 = paragraphSpacingOption.Spacing;
											item.Before = spacing2.Before;
											item.After = spacing2.After;
											item.Within = spacing2.Within;
											item.LineRuleWithin = spacing2.LineRuleWithin;
											Conventions.UsedParagraphSpacing[i] = new Tuple<LineSpacing, IndexedObject>(item, tuple.Item2);
										}
									}
								}
								tuple = null;
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
							usedParagraphSpacing = null;
							this.m_B = true;
							J();
							this.m_B = false;
							A(ParagraphSpacing);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void IndentChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
				{
					IndentItem obj = (IndentItem)comboBox.DataContext;
					Indent indent = obj.Indent;
					IndentOption indentOption = IndentOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(indentOption, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
						{
							List<Tuple<Indent, IndexedObject>> usedIndents = Conventions.UsedIndents;
							int num = usedIndents.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<Indent, IndexedObject> tuple = usedIndents[i];
								if (tuple.Item1.LeftIndent == indent.LeftIndent)
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
									if (tuple.Item1.FirstLineIndent == indent.FirstLineIndent)
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
										Indent item = default(Indent);
										Indent indent2 = indentOption.Indent;
										item.LeftIndent = indent2.LeftIndent;
										item.FirstLineIndent = indent2.FirstLineIndent;
										item.IndentLevel = indent2.IndentLevel;
										Conventions.UsedIndents[i] = new Tuple<Indent, IndexedObject>(item, tuple.Item2);
									}
								}
								tuple = null;
							}
							usedIndents = null;
							this.m_B = true;
							K();
							this.m_B = false;
							A(Indents);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void BulletStyleChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
				{
					BulletStyleItem obj = (BulletStyleItem)comboBox.DataContext;
					MsoNumberedBulletStyle style = obj.Style.Style;
					BulletStyleOption bulletStyleOption = BulletStyleOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(bulletStyleOption, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
						{
							List<Tuple<BulletStyle, IndexedObject>> usedBulletStyles = Conventions.UsedBulletStyles;
							int num = usedBulletStyles.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<BulletStyle, IndexedObject> tuple = usedBulletStyles[i];
								if (tuple.Item1.Style == style)
								{
									BulletStyle item = default(BulletStyle);
									BulletStyle style2 = bulletStyleOption.Style;
									item.Style = style2.Style;
									item.IndentLevel = style2.IndentLevel;
									Conventions.UsedBulletStyles[i] = new Tuple<BulletStyle, IndexedObject>(item, tuple.Item2);
								}
								tuple = null;
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
							usedBulletStyles = null;
							this.m_B = true;
							L();
							this.m_B = false;
							A(BulletStyles);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void TextBoxMarginsChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					TextBoxMarginsItem obj = (TextBoxMarginsItem)comboBox.DataContext;
					Margins margins = obj.Margins;
					MarginsOption marginsOption = TextBoxMarginsOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(marginsOption, ref listErrors);
					if (!listErrors.Any())
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
						if (A())
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
							List<Tuple<Margins, IndexedObject>> usedTextBoxMargins = Conventions.UsedTextBoxMargins;
							int num = usedTextBoxMargins.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<Margins, IndexedObject> tuple = usedTextBoxMargins[i];
								if (tuple.Item1.Top == margins.Top && tuple.Item1.Right == margins.Right && tuple.Item1.Bottom == margins.Bottom && tuple.Item1.Left == margins.Left)
								{
									Margins item = default(Margins);
									Margins margins2 = marginsOption.Margins;
									item.Top = margins2.Top;
									item.Right = margins2.Right;
									item.Bottom = margins2.Bottom;
									item.Left = margins2.Left;
									Conventions.UsedTextBoxMargins[i] = new Tuple<Margins, IndexedObject>(item, tuple.Item2);
								}
								tuple = null;
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
							usedTextBoxMargins = null;
							this.m_B = true;
							M();
							this.m_B = false;
							A(TextBoxMargins);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void CellMarginsChanged(object sender, SelectionChangedEventArgs e)
	{
		if (this.m_B)
		{
			return;
		}
		checked
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
				System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
				if (comboBox.IsLoaded)
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
					CellMarginsItem obj = (CellMarginsItem)comboBox.DataContext;
					Margins margins = obj.Margins;
					MarginsOption marginsOption = CellMarginsOptions[comboBox.SelectedIndex];
					List<string> listErrors = new List<string>();
					S();
					obj.Reformat(marginsOption, ref listErrors);
					if (!listErrors.Any())
					{
						if (A())
						{
							List<Tuple<Margins, IndexedObject>> usedCellMargins = Conventions.UsedCellMargins;
							int num = usedCellMargins.Count - 1;
							for (int i = 0; i <= num; i++)
							{
								Tuple<Margins, IndexedObject> tuple = usedCellMargins[i];
								if (tuple.Item1.Top == margins.Top)
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
									if (tuple.Item1.Right == margins.Right)
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
										if (tuple.Item1.Bottom == margins.Bottom && tuple.Item1.Left == margins.Left)
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
											Margins item = default(Margins);
											Margins margins2 = marginsOption.Margins;
											item.Top = margins2.Top;
											item.Right = margins2.Right;
											item.Bottom = margins2.Bottom;
											item.Left = margins2.Left;
											Conventions.UsedCellMargins[i] = new Tuple<Margins, IndexedObject>(item, tuple.Item2);
										}
									}
								}
								tuple = null;
							}
							usedCellMargins = null;
							this.m_B = true;
							N();
							this.m_B = false;
							A(CellMargins);
						}
					}
					else
					{
						A(listErrors);
					}
					T();
					listErrors = null;
				}
				comboBox = null;
				return;
			}
		}
	}

	private void A<A>(ObservableCollection<A> A)
	{
		checked
		{
			int num3 = default(int);
			if ((object)typeof(A) != typeof(FontColorItem))
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
				int num = AllItems.Count - 1;
				int num2 = 0;
				while (true)
				{
					if (num2 <= num)
					{
						if ((object)((object)AllItems[num2]).GetType() == typeof(A))
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
							num3 = num2;
							break;
						}
						num2 += AllItems[num2].Objects.Count;
						num2++;
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
					break;
				}
			}
			else
			{
				num3 = 0;
			}
			lbxResults.SelectionChanged -= NavigateObjects;
			this.m_B = true;
			List<BaseItem> allItems = AllItems;
			try
			{
				do
				{
					allItems.RemoveRange(num3, allItems[num3].Objects.Count + 1);
				}
				while ((object)((object)allItems[num3]).GetType() == typeof(A));
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0107;
					}
					continue;
					end_IL_0107:
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			List<A> list = A.ToList();
			list.Reverse();
			using (List<A>.Enumerator enumerator = list.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					BaseItem baseItem = enumerator.Current as BaseItem;
					allItems.Insert(num3, baseItem);
					allItems.InsertRange(num3 + 1, baseItem.Objects);
					baseItem = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_017c;
					}
					continue;
					end_IL_017c:
					break;
				}
			}
			list = null;
			allItems = null;
			try
			{
				SourceCollection.Refresh();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				IDisposable disposable = SourceCollection.DeferRefresh();
				try
				{
					SourceCollection.GroupDescriptions.Clear();
					SourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(AH.A(52438)));
				}
				finally
				{
					if (disposable != null)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							disposable.Dispose();
							break;
						}
					}
				}
				ProjectData.ClearProjectError();
			}
			this.m_B = false;
			lbxResults.SelectionChanged += NavigateObjects;
		}
	}

	private void S()
	{
		bdrWorking.Visibility = Visibility.Visible;
		System.Windows.Forms.Application.DoEvents();
	}

	private void T()
	{
		bdrWorking.Visibility = Visibility.Collapsed;
	}

	private bool A()
	{
		return true;
	}

	private void A(List<string> A)
	{
		_ = A.Count;
		string text = AH.A(53055);
		A = A.Distinct().ToList();
		using (List<string>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				text = text + AH.A(7894) + current;
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
				break;
			}
		}
		A = null;
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(AH.A(53363), UriKind.Relative);
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
		if (connectionId == 3)
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
					lbxResults = (System.Windows.Controls.ListBox)target;
					lbxResults.LostKeyboardFocus += ListBoxLostKeyboardFocus;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					bdrWorking = (Border)target;
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(GrowBorder);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(ShrinkBorder);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 2)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.CheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(ExpandItems);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = ToggleButton.UncheckedEvent;
			eventSetter.Handler = new RoutedEventHandler(CollapseItems);
			((Style)target).Setters.Add(eventSetter);
		}
		if (connectionId == 4)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += ColorChanged;
		}
		if (connectionId == 5)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += ColorChanged;
		}
		if (connectionId == 6)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FontFamilyChanged;
		}
		if (connectionId == 7)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FontStyleChanged;
		}
		if (connectionId == 8)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += TextDecorationChanged;
		}
		if (connectionId == 9)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += BorderDashStyleChanged;
		}
		if (connectionId == 10)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += BorderWeightChanged;
		}
		if (connectionId == 11)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FillTransparencyChanged;
		}
		if (connectionId == 12)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += TextBoxMarginsChanged;
		}
		if (connectionId == 13)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += CellMarginsChanged;
		}
		if (connectionId == 14)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += ParagraphSpacingChanged;
		}
		if (connectionId == 15)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += IndentChanged;
		}
		if (connectionId != 16)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += BulletStyleChanged;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
