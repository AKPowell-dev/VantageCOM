using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

[DesignerGenerated]
public sealed class wpfPrecedents : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	private sealed class SC
	{
		internal BaseItem A;

		internal BaseItem B;
	}

	public sealed class Precedent
	{
		[CompilerGenerated]
		private string A;

		[CompilerGenerated]
		private string B;

		[CompilerGenerated]
		private Precedent A;

		[CompilerGenerated]
		private List<Precedent> A;

		[CompilerGenerated]
		private bool A;

		public string Address
		{
			[CompilerGenerated]
			get
			{
				return this.A;
			}
			[CompilerGenerated]
			set
			{
				this.A = value;
			}
		}

		public string Value
		{
			[CompilerGenerated]
			get
			{
				return B;
			}
			[CompilerGenerated]
			set
			{
				B = value;
			}
		}

		public Precedent ProvisionalParent
		{
			[CompilerGenerated]
			get
			{
				return this.A;
			}
			[CompilerGenerated]
			set
			{
				this.A = value;
			}
		}

		public List<Precedent> ProvisionalChildren
		{
			[CompilerGenerated]
			get
			{
				return this.A;
			}
			[CompilerGenerated]
			set
			{
				this.A = value;
			}
		}

		public bool ProvisionalParentMismatched
		{
			[CompilerGenerated]
			get
			{
				return A;
			}
			[CompilerGenerated]
			set
			{
				A = value;
			}
		}

		public static void AddProvisionalPrec(string addr, Precedent provParent, ref List<string> curPrecAddrs, ref List<Precedent> provPrecs)
		{
			if (curPrecAddrs.Contains(addr))
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
				Precedent item = new Precedent
				{
					Address = addr,
					ProvisionalParent = provParent,
					ProvisionalParentMismatched = false
				};
				if (provParent.ProvisionalChildren == null)
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
					provParent.ProvisionalChildren = new List<Precedent>();
				}
				provParent.ProvisionalChildren.Add(item);
				curPrecAddrs.Add(addr);
				provPrecs.Add(item);
				return;
			}
		}

		public bool IsProvAndParentMatched()
		{
			if (ProvisionalParent != null)
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
						return !ProvisionalParentMismatched;
					}
				}
			}
			return false;
		}
	}

	public struct Argument
	{
		public string Name;

		public string Label;

		public string Value;

		public int Index;

		public int Length;
	}

	private enum TC
	{
		A,
		B,
		C
	}

	private sealed class VC
	{
		[Serializable]
		[CompilerGenerated]
		internal sealed class _Closure$__
		{
			public static readonly _Closure$__ A;

			public static Func<Match, int, string> A;

			public static Func<Match, int, string> B;

			public static Func<Match, int, string> C;

			public static Comparison<Precedent> A;

			static _Closure$__()
			{
				_Closure$__.A = new _Closure$__();
			}

			[SpecialName]
			internal string A(Match A, int B)
			{
				return string.Format(VH.A(212221), A.Groups[checked(1 + B)].Value);
			}

			[SpecialName]
			internal string B(Match A, int B)
			{
				return string.Format(VH.A(212242), A.Groups[checked(1 + B)].Value);
			}

			[SpecialName]
			internal string C(Match A, int B)
			{
				return checked(string.Format(VH.A(211932), A.Groups[1 + B].Value, A.Groups[2 + B].Value));
			}

			[SpecialName]
			internal int A(Precedent A, Precedent B)
			{
				return B.Address.Length.CompareTo(A.Address.Length);
			}
		}

		[CompilerGenerated]
		internal sealed class UC
		{
			public string A;

			public string B;

			[SpecialName]
			internal bool A(Precedent A)
			{
				return object.Equals(A.Address, this.A);
			}

			[SpecialName]
			internal bool B(Precedent A)
			{
				return object.Equals(A.Address, this.B);
			}
		}

		private const string m_A = "[A-Z]{1,3}";

		private const string B = "\\d{1,7}";

		[CompilerGenerated]
		private string C;

		private readonly string D;

		private readonly List<Precedent> m_A;

		private readonly int m_A;

		private Worksheet m_A;

		private const int B = 6;

		private const int C = 2;

		private string CellPatt
		{
			[CompilerGenerated]
			get
			{
				return this.C;
			}
		}

		internal VC(string A, ref List<Precedent> B, Worksheet C)
		{
			this.C = string.Format(VH.A(211932), VH.A(211289), VH.A(211274));
			D = A;
			this.m_A = B;
			this.m_A = C;
			this.m_A = this.m_A.Count;
		}

		internal void A()
		{
			if (!D.Contains(VH.A(2826)))
			{
				this.m_A = null;
				return;
			}
			checked
			{
				try
				{
					string a = string.Format(VH.A(211213), VH.A(211274), VH.A(211274));
					Func<Match, int, string> b;
					if (_Closure$__.A == null)
					{
						b = (_Closure$__.A = [SpecialName] (Match A, int B) => string.Format(VH.A(212221), A.Groups[1 + B].Value));
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
						b = _Closure$__.A;
					}
					A(a, b);
					A(string.Format(VH.A(211213), VH.A(211289), VH.A(211289)), [SpecialName] (Match A, int B) => string.Format(VH.A(212242), A.Groups[1 + B].Value));
					A(string.Format(VH.A(211310), VH.A(211289), VH.A(211274), CellPatt), [SpecialName] (Match A, int B) => string.Format(VH.A(211932), A.Groups[1 + B].Value, A.Groups[2 + B].Value));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(new Exception(string.Format(VH.A(211381), ex2.Message), ex2));
					ProjectData.ClearProjectError();
				}
				finally
				{
					if (this.m_A.Count > this.m_A)
					{
						List<Precedent> a2 = this.m_A;
						Comparison<Precedent> comparison;
						if (_Closure$__.A == null)
						{
							comparison = (_Closure$__.A = [SpecialName] (Precedent A, Precedent B) => B.Address.Length.CompareTo(A.Address.Length));
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
							comparison = _Closure$__.A;
						}
						a2.Sort(comparison);
					}
					this.m_A = null;
				}
			}
		}

		private void A(string A, Func<Match, int, string> B)
		{
			A = this.A(A);
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = Regex.Matches(D, A).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match match = (Match)enumerator.Current;
					if (match.Index > 0)
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
						if (Operators.CompareString(Conversions.ToString(D[checked(match.Index - 1)]), VH.A(49303), TextCompare: false) == 0)
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
					}
					string text = Regex.Escape(match.Groups[2].Value);
					string text2 = B(match, 6);
					bool flag = !string.IsNullOrEmpty(text);
					string text3 = VH.A(211454);
					string arg = string.Format(VH.A(211499), text3, text, text2, CellPatt);
					string arg2 = string.Format(VH.A(211530), text3, text, text2);
					this.A(D, string.Format(VH.A(211553), arg, arg2), flag);
					this.A(match, flag);
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}

		private void A(string A, string B, bool C)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = Regex.Matches(A, B).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Match a = (Match)enumerator.Current;
					this.A(a, C);
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
					return;
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
		}

		private void A(Match A, bool B)
		{
			int num = A.Value.LastIndexOf('!');
			string A2 = checked(string.Format(VH.A(49936), A.Value.Substring(0, num + 1), A.Value.Substring(num + 1).Replace(VH.A(41262), "")));
			if (this.m_A.FirstOrDefault([SpecialName] (Precedent precedent) => object.Equals(precedent.Address, A2)) != null)
			{
				return;
			}
			string B2 = "";
			if (B && VC.A(A2, ref B2))
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
				if (this.m_A.FirstOrDefault([SpecialName] (Precedent precedent) => object.Equals(precedent.Address, B2)) != null)
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
			}
			try
			{
				Range range;
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
					range = ((_Worksheet)this.m_A).get_Range((object)A2, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				else
				{
					range = ((_Application)this.m_A.Application).get_Range((object)A2, RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Range range2 = range;
				List<Precedent> a = this.m_A;
				Precedent obj = new Precedent
				{
					Address = A2
				};
				string format = VH.A(48282);
				object obj2;
				if (!Operators.ConditionalCompareObjectLess(range2.CountLarge, 2, TextCompare: false))
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
					obj2 = VH.A(41885);
				}
				else
				{
					obj2 = range2.Text;
				}
				obj.Value = string.Format(format, RuntimeHelpers.GetObjectValue(obj2));
				a.Add(obj);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (!A2.StartsWith(VH.A(43335)))
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
					clsReporting.LogException(new Exception(string.Format(VH.A(211576), A.Value, ex2.Message), ex2));
				}
				ProjectData.ClearProjectError();
			}
			finally
			{
				Range range2 = null;
			}
		}

		private static bool A(string A, ref string B)
		{
			if (!A.StartsWith(VH.A(39851)))
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
						return false;
					}
				}
			}
			int num = A.LastIndexOf(VH.A(43343));
			if (num < 1)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
			B = checked(A.Substring(1, num - 1).Replace(VH.A(39854), VH.A(39851)) + A.Substring(num + 1));
			return true;
		}

		private string A(string A)
		{
			return string.Format(VH.A(211669), VH.A(211686), A);
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure$__
	{
		public static readonly _Closure$__ A;

		public static Func<SC, BaseItem> A;

		public static Func<SC, BaseItem> B;

		public static Func<BaseItem, int> A;

		public static Func<BaseItem, int> B;

		public static Func<Precedent, string> A;

		public static Action<Precedent> A;

		public static Comparison<Precedent> A;

		public static Action<Precedent> B;

		public static Comparison<Precedent> B;

		public static Action<Precedent> C;

		public static Func<Precedent, bool> A;

		public static Func<char, bool> A;

		public static Func<Match, string> A;

		public static Comparison<string> A;

		public static Func<Precedent, string> B;

		public static Func<BaseItem, int> C;

		public static Func<BaseItem, int> D;

		public static Func<BaseItem, int> E;

		public static Func<BaseItem, int> F;

		public static Comparison<BaseItem> A;

		static _Closure$__()
		{
			_Closure$__.A = new _Closure$__();
		}

		[SpecialName]
		internal BaseItem A(SC A)
		{
			return A.B;
		}

		[SpecialName]
		internal BaseItem B(SC A)
		{
			return A.A;
		}

		[SpecialName]
		internal int A(BaseItem A)
		{
			return A.SelectionIndex;
		}

		[SpecialName]
		internal int B(BaseItem A)
		{
			return A.SelectionLength;
		}

		[SpecialName]
		internal string A(Precedent A)
		{
			return A.Address;
		}

		[SpecialName]
		internal void A(Precedent A)
		{
			if (A.Address.Contains(VH.A(7827)))
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
				A.Address = VH.A(211957) + A.Address;
				return;
			}
		}

		[SpecialName]
		internal int A(Precedent A, Precedent B)
		{
			return A.Address.CompareTo(B.Address);
		}

		[SpecialName]
		internal void B(Precedent A)
		{
			A.Address = A.Address.Replace(VH.A(211957), "");
		}

		[SpecialName]
		internal int B(Precedent A, Precedent B)
		{
			return B.Address.Length.CompareTo(A.Address.Length);
		}

		[SpecialName]
		internal void C(Precedent A)
		{
			A.ProvisionalParentMismatched = true;
		}

		[SpecialName]
		internal bool A(Precedent A)
		{
			return A.IsProvAndParentMatched();
		}

		[SpecialName]
		internal bool A(char A)
		{
			return A == '\'';
		}

		[SpecialName]
		internal string A(Match A)
		{
			return A.Value;
		}

		[SpecialName]
		internal int C(string A, string B)
		{
			return B.Length.CompareTo(A.Length);
		}

		[SpecialName]
		internal string B(Precedent A)
		{
			return A.Address;
		}

		[SpecialName]
		internal int C(BaseItem A)
		{
			return A.SelectionIndex;
		}

		[SpecialName]
		internal int D(BaseItem A)
		{
			return A.SelectionEnd;
		}

		[SpecialName]
		internal int E(BaseItem A)
		{
			return A.SelectionIndex;
		}

		[SpecialName]
		internal int F(BaseItem A)
		{
			return A.SelectionLength;
		}

		[SpecialName]
		internal int D(BaseItem A, BaseItem B)
		{
			return A.SelectionIndex.CompareTo(B.SelectionIndex);
		}
	}

	[CompilerGenerated]
	internal sealed class WC
	{
		public Range A;

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
		internal bool A(BaseItem A)
		{
			return Operators.CompareString(A.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), this.A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class XC
	{
		public List<string> A;

		public List<Precedent> A;

		public string A;

		public XC(XC A)
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
				this.A = A.A;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A(Precedent A)
		{
			int num = A.Address.IndexOf(VH.A(2826));
			if (num < 0)
			{
				return;
			}
			string text = A.Address.Substring(0, num);
			checked
			{
				string text2 = A.Address.Substring(num + 1);
				if (text.Contains(VH.A(7827)) && !text2.Contains(VH.A(7827)))
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
					text2 = string.Format(VH.A(49936), text.Substring(0, text.LastIndexOf('!') + 1), text2);
				}
				Precedent.AddProvisionalPrec(text, A, ref this.A, ref this.A);
				Precedent.AddProvisionalPrec(text2, A, ref this.A, ref this.A);
			}
		}

		[SpecialName]
		internal bool A(Match A)
		{
			string source = this.A.Substring(0, A.Index);
			Func<char, bool> predicate;
			if (_Closure$__.A == null)
			{
				predicate = (_Closure$__.A = _Closure$__.A.A);
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
				predicate = _Closure$__.A;
			}
			return source.Count(predicate) % 2 == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class YC
	{
		public Argument A;

		public YC(YC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class ZC
	{
		public int A;

		public YC A;

		public ZC(ZC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseItem A)
		{
			if (A.SelectionIndex >= this.A.A.Index)
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
						return A.SelectionEnd <= this.A;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class AD
	{
		public BaseItem A;

		public AD(AD A)
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
		internal SC A(BaseItem A)
		{
			return new SC
			{
				A = A,
				B = this.A
			};
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private static FieldInfo m_A;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private Range m_A;

	private List<Worksheet> m_A;

	private List<Microsoft.Office.Interop.Excel.Workbook> m_A;

	private Microsoft.Office.Interop.Excel.Workbook m_A;

	private Microsoft.Office.Interop.Excel.Workbook m_B;

	private Worksheet m_A;

	private Range m_B;

	private RoutedPropertyChangedEventHandler<object> m_A;

	private bool m_A;

	private bool m_B;

	private ScrollViewer m_A;

	private Visibility m_A;

	private bool m_C;

	private readonly int m_A;

	private readonly int m_B;

	private ObservableCollection<BaseItem> m_A;

	private double m_A;

	private Visibility m_B;

	[CompilerGenerated]
	private frmPrecedentsHost m_A;

	[CompilerGenerated]
	private object m_A;

	[CompilerGenerated]
	private BaseItem m_A;

	private double m_B;

	private bool m_D;

	private ScrollViewer m_B;

	private bool m_E;

	private static readonly string m_A = VH.A(49901);

	[CompilerGenerated]
	[AccessedThroughProperty("ThisWindow")]
	private wpfPrecedents m_A;

	[AccessedThroughProperty("grdMain")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("grdSplit")]
	[CompilerGenerated]
	private Grid m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("tbFormula")]
	private TextBlock m_A;

	[AccessedThroughProperty("tbDummy")]
	[CompilerGenerated]
	private TextBlock m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkWrap")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("masterC")]
	[CompilerGenerated]
	private ColumnDefinition m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("trvTrace")]
	private System.Windows.Controls.TreeView m_A;

	[AccessedThroughProperty("chkSettings")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("popSettings")]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbEvaluate")]
	private TextBlock m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkEvaluate")]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("tbArguments")]
	[CompilerGenerated]
	private TextBlock m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkArguments")]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkArrows")]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkUnhide")]
	private System.Windows.Controls.CheckBox m_F;

	[AccessedThroughProperty("chkOpenLinks")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("tbHighlight")]
	[CompilerGenerated]
	private TextBlock m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHighlight")]
	private System.Windows.Controls.CheckBox m_H;

	[AccessedThroughProperty("chkMove")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_I;

	private bool m_F;

	public ObservableCollection<BaseItem> RootItems
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(42688));
		}
	}

	public double TreeViewViewportWidth
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(VH.A(42707));
		}
	}

	public Visibility ArgumentsColumnVisibility
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(VH.A(46358));
		}
	}

	public frmPrecedentsHost Host
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

	public object HighlightedInline
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public BaseItem AuditedItem
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

	internal virtual wpfPrecedents ThisWindow
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

	internal virtual Grid grdMain
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

	internal virtual Grid grdSplit
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

	internal virtual ScrollViewer scroller
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

	internal virtual TextBlock tbFormula
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

	internal virtual TextBlock tbDummy
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

	internal virtual System.Windows.Controls.CheckBox chkWrap
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

	internal virtual ColumnDefinition masterC
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

	internal virtual System.Windows.Controls.TreeView trvTrace
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
			System.Windows.Input.KeyEventHandler value2 = TreeViewPreviewKeyDown;
			System.Windows.Controls.TreeView treeView = this.m_A;
			if (treeView != null)
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
				treeView.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			treeView = this.m_A;
			if (treeView != null)
			{
				treeView.PreviewKeyDown += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSettings
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

	internal virtual System.Windows.Controls.Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
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
				switch (1)
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

	internal virtual System.Windows.Controls.Button btnCancel
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
			RoutedEventHandler value2 = btnCancel_Click;
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual Popup popSettings
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
			EventHandler value2 = SettingsPopupOpened;
			EventHandler value3 = SettingsPopupClosed;
			System.Windows.Input.KeyEventHandler value4 = CloseSettingsPopup;
			Popup popup = this.m_A;
			if (popup != null)
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
				popup.Opened -= value2;
				popup.Closed -= value3;
				popup.PreviewKeyDown -= value4;
			}
			this.m_A = value;
			popup = this.m_A;
			if (popup == null)
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
				popup.Opened += value2;
				popup.Closed += value3;
				popup.PreviewKeyDown += value4;
				return;
			}
		}
	}

	internal virtual TextBlock tbEvaluate
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

	internal virtual System.Windows.Controls.CheckBox chkEvaluate
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

	internal virtual TextBlock tbArguments
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

	internal virtual System.Windows.Controls.CheckBox chkArguments
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

	internal virtual System.Windows.Controls.CheckBox chkArrows
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

	internal virtual System.Windows.Controls.CheckBox chkUnhide
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

	internal virtual System.Windows.Controls.CheckBox chkOpenLinks
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

	internal virtual TextBlock tbHighlight
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

	internal virtual System.Windows.Controls.CheckBox chkHighlight
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

	internal virtual System.Windows.Controls.CheckBox chkMove
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
				switch (7)
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

	public wpfPrecedents(frmPrecedentsHost frm)
	{
		base.Loaded += wpfPrecedents_Loaded;
		base.PreviewKeyDown += wpfChild_KeyUp;
		this.m_A = [SpecialName] (object a0, RoutedPropertyChangedEventArgs<object> a1) =>
		{
			TreeViewSelectionChanged((System.Windows.Controls.TreeView)a0, a1);
		};
		this.m_B = false;
		this.m_C = false;
		this.m_A = 40;
		this.m_B = 7;
		this.m_A = null;
		this.m_B = 0.0;
		this.m_D = false;
		this.m_B = null;
		InitializeComponent();
		Host = frm;
		A();
		this.m_A = MH.A.Application;
		this.m_A = new List<Worksheet>();
		try
		{
			this.m_A = this.m_A.ActiveWindow.ActivePane.VisibleRange;
			this.m_B = (Range)this.m_A.Selection;
			this.m_A = this.m_B.Worksheet;
			this.m_B = (Microsoft.Office.Interop.Excel.Workbook)this.m_A.Parent;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Host.Close();
			ProjectData.ClearProjectError();
		}
		if (!Access.IsLegacyPlan())
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
			if (!Access.AllowExcelOperation((PlanType)5, (Restriction)0, true))
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
				K.Settings.AuditEvaluateFormulas = false;
				chkEvaluate.IsEnabled = false;
				tbEvaluate.IsEnabled = false;
				K.Settings.AuditHighlightCells = false;
				chkHighlight.IsEnabled = false;
				tbHighlight.IsEnabled = false;
			}
		}
		MySettings settings = K.Settings;
		chkWrap.IsChecked = settings.AuditFormulaWrap;
		chkEvaluate.IsChecked = settings.AuditEvaluateFormulas;
		chkArguments.IsChecked = settings.AuditEvaluateArguments;
		chkUnhide.IsChecked = settings.AuditUnhideRowsColumns;
		chkOpenLinks.IsChecked = settings.AuditOpenWorkbookLinks;
		chkHighlight.IsChecked = settings.AuditHighlightCells;
		settings = null;
		if (chkEvaluate.IsChecked == true)
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
			tbArguments.IsEnabled = true;
			chkArguments.IsEnabled = true;
		}
		else
		{
			tbArguments.IsEnabled = false;
			chkArguments.IsEnabled = false;
		}
		chkWrap.Checked += EnableFormulaWrap;
		chkWrap.Unchecked += DisableFormulaWrap;
		if (chkWrap.IsChecked == true)
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
			tbFormula.TextWrapping = TextWrapping.Wrap;
		}
		chkEvaluate.Checked += EvaluateCheckedChanged;
		chkEvaluate.Unchecked += EvaluateCheckedChanged;
		chkArguments.Checked += ArgumentsCheckedChanged;
		chkArguments.Unchecked += ArgumentsCheckedChanged;
		chkHighlight.Checked += HighlightCheckedChanged;
		chkHighlight.Unchecked += HighlightCheckedChanged;
		chkOpenLinks.Checked += OpenLinksCheckedChanged;
		chkOpenLinks.Unchecked += OpenLinksCheckedChanged;
		chkUnhide.Checked += UnhideCheckedChanged;
		chkUnhide.Unchecked += UnhideCheckedChanged;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(42971)).AddEventHandler(this.m_A, new AppEvents_WorkbookBeforeCloseEventHandler(A));
		RootItems = new ObservableCollection<BaseItem>();
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

	private static void A()
	{
		wpfPrecedents.m_A = typeof(SystemParameters).GetField(VH.A(42651), BindingFlags.Static | BindingFlags.NonPublic);
		B();
		SystemParameters.StaticPropertyChanged += SystemParameters_StaticPropertyChanged;
	}

	private static void SystemParameters_StaticPropertyChanged(object sender, PropertyChangedEventArgs e)
	{
		B();
	}

	private static void B()
	{
		if (!SystemParameters.MenuDropAlignment)
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
			if ((object)wpfPrecedents.m_A == null)
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
				wpfPrecedents.m_A.SetValue(null, false);
				return;
			}
		}
	}

	private void wpfPrecedents_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A = Visibility.Collapsed;
		this.m_A = (ScrollViewer)Forms.GetScrollViewer((DependencyObject)trvTrace);
		TreeViewViewportWidth = this.m_A.ViewportWidth;
		this.m_A.ScrollChanged += TreeViewScrollChanged;
		this.m_A.SizeChanged += TreeViewSizeChanged;
		ArgumentsColumnVisibility = Visibility.Collapsed;
		masterC.Width = new GridLength(100.0);
		C();
		Base.CheckForWorkshare();
	}

	private void C()
	{
		bool flag = false;
		Range range = null;
		Range b = this.m_B;
		if (Operators.ConditionalCompareObjectGreater(b.Cells.CountLarge, 1, TextCompare: false))
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
			bool flag2 = Information.IsDBNull(RuntimeHelpers.GetObjectValue(b.MergeCells));
			if (!flag2)
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
				if (Conversions.ToBoolean(b.MergeCells))
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
					range = (Range)this.m_B.Cells[1, 1];
					goto IL_05b9;
				}
			}
			if (flag2)
			{
				Range range2 = null;
				try
				{
					range = b.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = range.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range range3 = (Range)enumerator.Current;
							if (range3.MergeArea != null)
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
								range2 = ((range2 == null) ? ((Range)range3.MergeArea.Cells[1, 1]) : this.m_A.Union(range2, (Range)range3.MergeArea.Cells[1, 1], RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
							}
							else if (range2 != null)
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
								range2 = this.m_A.Union(range2, range3, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							else
							{
								range2 = range3;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_042e;
							}
							continue;
							end_IL_042e:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					range = range2;
					range2 = null;
					if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 5, TextCompare: false))
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								range = null;
								range2 = null;
								D(VH.A(46409));
								Close();
								return;
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
				finally
				{
					range2 = null;
				}
			}
			else
			{
				try
				{
					range = b.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
					if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 5, TextCompare: false))
					{
						range = null;
						D(VH.A(46409));
						Close();
						return;
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
		}
		else if (Conversions.ToBoolean(b.HasFormula))
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
			range = this.m_B;
		}
		else
		{
			try
			{
				if (Conversions.ToBoolean(NewLateBinding.LateGet(b, null, VH.A(46494), new object[0], null, null, null)))
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						range = (Range)NewLateBinding.LateGet(b, null, VH.A(46511), new object[0], null, null, null);
						break;
					}
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		goto IL_05b9;
		IL_0eac:
		if (flag)
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
			if (Operators.ConditionalCompareObjectEqual(this.m_B.Cells.CountLarge, 1, TextCompare: false))
			{
				this.m_A.ScreenUpdating = false;
				try
				{
					Range rng = (Range)this.m_A.Selection;
					List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
					if (A(this.m_B))
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
						D(VH.A(47546));
					}
					else
					{
						C(VH.A(47625));
					}
					Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
					Base.ReturnToPreviousRange(rng);
				}
				catch (Exception ex7)
				{
					ProjectData.SetProjectError(ex7);
					Exception ex8 = ex7;
					C(VH.A(47625));
					ProjectData.ClearProjectError();
				}
				finally
				{
				}
				this.m_A.ScreenUpdating = true;
			}
			else
			{
				C(VH.A(47694));
			}
			Close();
			return;
		}
		IL_0d84:
		RootItems.Clear();
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = range.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Range rng2 = (Range)enumerator2.Current;
				RootItems.Add(new RootItem(rng2));
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0dc8;
				}
				continue;
				end_IL_0dc8:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		checked
		{
			int num = RootItems.Count - 1;
			for (int i = 0; i <= num; i++)
			{
				AuditedItem = RootItems[i];
				C(AuditedItem);
				if (AuditedItem.Items.Count <= 0)
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
				A(AuditedItem);
				RemoveTreeViewSelectionChangedHandler();
				AuditedItem.IsSelected = true;
				E();
				AuditedItem.IsExpanded = true;
				trvTrace.Focus();
				flag = true;
				break;
			}
			range = null;
			goto IL_0eac;
		}
		IL_05b9:
		b = null;
		if (range != null)
		{
			bool? isChecked = chkArguments.IsChecked;
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
					goto IL_0d84;
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
			if (Dialog.ArgumentNames == null)
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
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					Dialog.ArgumentNames = new Dictionary<string, List<string>>();
					Dictionary<string, List<string>> argumentNames = Dialog.ArgumentNames;
					argumentNames.Add(VH.A(3794), new List<string>(new string[3]
					{
						VH.A(46534),
						VH.A(46559),
						VH.A(46590)
					}));
					argumentNames.Add(VH.A(2056), new List<string>(new string[6]
					{
						VH.A(46623),
						VH.A(46648),
						VH.A(46673),
						VH.A(46698),
						VH.A(46727),
						VH.A(46752)
					}));
					argumentNames.Add(VH.A(2015), new List<string>(new string[4]
					{
						VH.A(46623),
						VH.A(46779),
						VH.A(46802),
						VH.A(46829)
					}));
					argumentNames.Add(VH.A(2030), new List<string>(new string[4]
					{
						VH.A(46623),
						VH.A(46779),
						VH.A(46858),
						VH.A(46829)
					}));
					argumentNames.Add(VH.A(4444), new List<string>(new string[5]
					{
						VH.A(46885),
						VH.A(2877),
						VH.A(46904),
						VH.A(46913),
						VH.A(46930)
					}));
					argumentNames.Add(VH.A(2045), new List<string>(new string[3]
					{
						VH.A(46623),
						VH.A(46648),
						VH.A(46945)
					}));
					argumentNames.Add(VH.A(4533), new List<string>(new string[3]
					{
						VH.A(46970),
						VH.A(46981),
						VH.A(46998)
					}));
					argumentNames.Add(VH.A(47021), new List<string>(new string[2]
					{
						VH.A(46970),
						VH.A(46981)
					}));
					argumentNames.Add(VH.A(47036), new List<string>(new string[2]
					{
						VH.A(47051),
						VH.A(47072)
					}));
					argumentNames.Add(VH.A(47085), new List<string>(new string[3]
					{
						VH.A(47094),
						VH.A(47103),
						VH.A(47116)
					}));
					argumentNames.Add(VH.A(47127), new List<string>(new string[5]
					{
						VH.A(47094),
						VH.A(47132),
						VH.A(47141),
						VH.A(47148),
						VH.A(47157)
					}));
					argumentNames.Add(VH.A(47170), new List<string>(new string[5]
					{
						VH.A(47094),
						VH.A(47132),
						VH.A(47141),
						VH.A(47175),
						VH.A(47157)
					}));
					argumentNames.Add(VH.A(47184), new List<string>(new string[3]
					{
						VH.A(47103),
						VH.A(47116),
						VH.A(47193)
					}));
					argumentNames.Add(VH.A(47208), new List<string>(new string[2]
					{
						VH.A(47103),
						VH.A(47193)
					}));
					argumentNames.Add(VH.A(47215), new List<string>(new string[5]
					{
						VH.A(47094),
						VH.A(47141),
						VH.A(47224),
						VH.A(47175),
						VH.A(47157)
					}));
					argumentNames.Add(VH.A(47229), new List<string>(new string[7]
					{
						VH.A(47240),
						VH.A(47261),
						VH.A(47094),
						VH.A(47278),
						VH.A(47283),
						VH.A(47304),
						VH.A(47323)
					}));
					argumentNames.Add(VH.A(47338), new List<string>(new string[6]
					{
						VH.A(47240),
						VH.A(47261),
						VH.A(47355),
						VH.A(47368),
						VH.A(47304),
						VH.A(47323)
					}));
					argumentNames.Add(VH.A(47375), new List<string>(new string[1] { VH.A(47390) }));
					argumentNames.Add(VH.A(47401), new List<string>(new string[1] { VH.A(47390) }));
					argumentNames.Add(VH.A(47410), new List<string>(new string[1] { VH.A(47423) }));
					argumentNames.Add(VH.A(47446), new List<string>(new string[1] { VH.A(47423) }));
					argumentNames.Add(VH.A(47453), new List<string>(new string[2]
					{
						VH.A(47462),
						VH.A(47051)
					}));
					argumentNames.Add(VH.A(47479), new List<string>(new string[3]
					{
						VH.A(47051),
						VH.A(47462),
						VH.A(47494)
					}));
					argumentNames.Add(VH.A(47511), new List<string>(new string[2]
					{
						VH.A(47518),
						VH.A(47531)
					}));
					argumentNames = null;
				}
			}
			goto IL_0d84;
		}
		goto IL_0eac;
	}

	private void wpfChild_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		System.Windows.Input.KeyEventArgs e2 = e;
		if (e2.Key == Key.Escape)
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
			if (!popSettings.IsOpen)
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
				this.m_A = true;
				D();
				e2.Handled = true;
			}
		}
		else if (e2.Key == Key.Return)
		{
			this.m_A = false;
			Close();
			e2.Handled = true;
		}
		else if (Base.ProcessShortcut(Host, e))
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
			e2.Handled = true;
		}
		else
		{
			if (e2.Key == Key.F2)
			{
				goto IL_00d4;
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
			if (e2.Key == Key.E)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Alt)
				{
					goto IL_00d4;
				}
			}
			if (e2.Key == Key.E)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
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
					System.Windows.Controls.CheckBox checkBox = chkEvaluate;
					bool? isChecked = chkEvaluate.IsChecked;
					bool? isChecked2;
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
						isChecked2 = isChecked;
					}
					else
					{
						isChecked2 = isChecked != true;
					}
					checkBox.IsChecked = isChecked2;
					e2.Handled = true;
					goto IL_0238;
				}
			}
			if (e2.Key == Key.H)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
				{
					System.Windows.Controls.CheckBox checkBox2 = chkHighlight;
					bool? isChecked = chkHighlight.IsChecked;
					bool? isChecked3;
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
						isChecked3 = isChecked;
					}
					else
					{
						isChecked3 = isChecked != true;
					}
					checkBox2.IsChecked = isChecked3;
					e2.Handled = true;
					goto IL_0238;
				}
			}
			if (e2.Key == Key.W)
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
				if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
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
					System.Windows.Controls.CheckBox checkBox3 = chkWrap;
					bool? isChecked = chkWrap.IsChecked;
					checkBox3.IsChecked = (!isChecked) ?? isChecked;
					e2.Handled = true;
				}
			}
		}
		goto IL_0238;
		IL_00d4:
		I();
		e2.Handled = true;
		goto IL_0238;
		IL_0238:
		e2 = null;
	}

	public void Close()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		int num5 = default(int);
		int num6 = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 980:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0014;
						case 4:
							goto IL_0030;
						case 5:
							goto IL_0037;
						case 6:
							goto IL_0044;
						case 7:
							goto IL_005b;
						case 8:
							goto IL_0091;
						case 9:
							goto IL_00b1;
						case 10:
							goto IL_00bb;
						case 11:
							goto IL_00c4;
						case 12:
							goto IL_00de;
						case 13:
							goto IL_00f8;
						case 14:
							goto IL_0114;
						case 15:
							goto IL_012e;
						case 16:
							goto IL_014a;
						case 17:
							goto IL_0166;
						case 18:
							goto IL_0182;
						case 19:
							goto IL_019c;
						case 20:
							goto IL_01b8;
						case 21:
							goto IL_01d4;
						case 22:
							goto IL_01f0;
						case 23:
							goto IL_020a;
						case 24:
							goto IL_0226;
						case 25:
							goto IL_0242;
						case 26:
							goto IL_0256;
						case 27:
							goto IL_028b;
						case 28:
							goto IL_02a1;
						case 29:
							goto IL_02ab;
						case 30:
							goto IL_02b5;
						case 31:
							goto IL_02bf;
						case 32:
							goto IL_02c9;
						case 33:
							goto IL_02d3;
						case 34:
							goto IL_02dd;
						case 35:
							goto IL_02e7;
						case 36:
							goto IL_02f1;
						case 37:
							goto IL_02fb;
						case 38:
							goto IL_0305;
						case 39:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 40:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_02e7:
					num2 = 35;
					HighlightedInline = null;
					goto IL_02f1;
					IL_0007:
					num2 = 2;
					this.m_A.ClearArrows();
					goto IL_0014;
					IL_0014:
					num2 = 3;
					Base.HideSheets(this.m_A, this.m_A, !this.m_A);
					goto IL_0030;
					IL_0030:
					num2 = 4;
					Base.RemoveHighlight();
					goto IL_0037;
					IL_0037:
					num2 = 5;
					if (this.m_A != null)
					{
						goto IL_0044;
					}
					goto IL_00bb;
					IL_0044:
					num2 = 6;
					num5 = checked(this.m_A.Count - 1);
					num6 = 0;
					goto IL_0099;
					IL_0099:
					if (num6 <= num5)
					{
						goto IL_005b;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_00b1;
					IL_02f1:
					num2 = 36;
					AuditedItem = null;
					goto IL_02fb;
					IL_02fb:
					num2 = 37;
					RootItems = null;
					goto IL_0305;
					IL_0305:
					num2 = 38;
					this.m_A = null;
					break;
					IL_00b1:
					num2 = 9;
					this.m_A = null;
					goto IL_00bb;
					IL_005b:
					num2 = 7;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(47767)).RemoveEventHandler(this.m_A, new AppEvents_WorkbookActivateEventHandler(this.B));
					goto IL_0091;
					IL_0091:
					num2 = 8;
					num6 = checked(num6 + 1);
					goto IL_0099;
					IL_00bb:
					num2 = 10;
					RemoveTreeViewSelectionChangedHandler();
					goto IL_00c4;
					IL_00c4:
					num2 = 11;
					this.m_A.ScrollChanged -= TreeViewScrollChanged;
					goto IL_00de;
					IL_00de:
					num2 = 12;
					this.m_A.SizeChanged -= TreeViewSizeChanged;
					goto IL_00f8;
					IL_00f8:
					num2 = 13;
					chkWrap.Checked -= EnableFormulaWrap;
					goto IL_0114;
					IL_0114:
					num2 = 14;
					chkWrap.Unchecked -= DisableFormulaWrap;
					goto IL_012e;
					IL_012e:
					num2 = 15;
					chkEvaluate.Checked -= EvaluateCheckedChanged;
					goto IL_014a;
					IL_014a:
					num2 = 16;
					chkEvaluate.Unchecked -= EvaluateCheckedChanged;
					goto IL_0166;
					IL_0166:
					num2 = 17;
					chkArguments.Checked -= ArgumentsCheckedChanged;
					goto IL_0182;
					IL_0182:
					num2 = 18;
					chkArguments.Unchecked -= ArgumentsCheckedChanged;
					goto IL_019c;
					IL_019c:
					num2 = 19;
					chkHighlight.Checked -= HighlightCheckedChanged;
					goto IL_01b8;
					IL_01b8:
					num2 = 20;
					chkHighlight.Unchecked -= HighlightCheckedChanged;
					goto IL_01d4;
					IL_01d4:
					num2 = 21;
					chkOpenLinks.Checked -= OpenLinksCheckedChanged;
					goto IL_01f0;
					IL_01f0:
					num2 = 22;
					chkOpenLinks.Unchecked -= OpenLinksCheckedChanged;
					goto IL_020a;
					IL_020a:
					num2 = 23;
					chkUnhide.Checked -= UnhideCheckedChanged;
					goto IL_0226;
					IL_0226:
					num2 = 24;
					chkUnhide.Unchecked -= UnhideCheckedChanged;
					goto IL_0242;
					IL_0242:
					num2 = 25;
					SystemParameters.StaticPropertyChanged -= SystemParameters_StaticPropertyChanged;
					goto IL_0256;
					IL_0256:
					num2 = 26;
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(42971)).RemoveEventHandler(this.m_A, new AppEvents_WorkbookBeforeCloseEventHandler(A));
					goto IL_028b;
					IL_028b:
					num2 = 27;
					A(this.m_A.ActiveWorkbook);
					goto IL_02a1;
					IL_02a1:
					num2 = 28;
					this.m_A = null;
					goto IL_02ab;
					IL_02ab:
					num2 = 29;
					this.m_A = null;
					goto IL_02b5;
					IL_02b5:
					num2 = 30;
					this.m_A = null;
					goto IL_02bf;
					IL_02bf:
					num2 = 31;
					this.m_A = null;
					goto IL_02c9;
					IL_02c9:
					num2 = 32;
					this.m_B = null;
					goto IL_02d3;
					IL_02d3:
					num2 = 33;
					this.m_A = null;
					goto IL_02dd;
					IL_02dd:
					num2 = 34;
					this.m_B = null;
					goto IL_02e7;
					end_IL_0000_2:
					break;
				}
				num2 = 39;
				Host.Close();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 980;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}

	private void D()
	{
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		Base.CollapseExpandedCells();
		try
		{
			this.m_B.Activate();
			this.m_A.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			this.m_A.Goto(this.m_B, false);
			this.m_A.Goto(this.m_B, false);
			Base.ScrollTo(this.m_A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		this.m_A.EnableEvents = true;
		this.m_A.ScreenUpdating = true;
		Close();
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		this.m_A = false;
		Close();
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		this.m_A = true;
		D();
	}

	private void TreeViewScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_A.ComputedVerticalScrollBarVisibility != this.m_A)
		{
			this.m_A = this.m_A.ComputedVerticalScrollBarVisibility;
			TreeViewViewportWidth = this.m_A.ViewportWidth;
		}
	}

	private void TreeViewSizeChanged(object sender, SizeChangedEventArgs e)
	{
		TreeViewViewportWidth = Base.TreeViewSizeChanged(e, this.m_A);
	}

	private void TreeViewPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
	{
		Key key = e.Key;
		if (key != Key.Up)
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
			if (key != Key.Down)
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
		}
		if (!e.IsRepeat)
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
			RemoveTreeViewSelectionChangedHandler();
			trvTrace.KeyUp += NavKeyUp;
		}
		if (e.Key == Key.Up && ((BaseItem)trvTrace.Items[0]).IsSelected)
		{
			e.Handled = true;
		}
	}

	private void NavKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
	{
		trvTrace.KeyUp -= NavKeyUp;
		E();
		F();
		e.Handled = true;
	}

	private void E()
	{
		trvTrace.SelectedItemChanged += this.m_A;
	}

	public void RemoveTreeViewSelectionChangedHandler()
	{
		trvTrace.SelectedItemChanged -= this.m_A;
	}

	private void TreeViewSelectionChanged(System.Windows.Controls.TreeView sender, RoutedPropertyChangedEventArgs<object> e)
	{
		F();
	}

	private void F()
	{
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		if (baseItem == null)
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
			Host.IgnoreDeactivate = true;
			try
			{
				A(baseItem.Range);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Host.IgnoreDeactivate = false;
			if (!Host.Focused)
			{
				trvTrace.Focus();
			}
			BaseItem baseItem2;
			try
			{
				baseItem2 = B(baseItem);
				if (AuditedItem != baseItem2)
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
					A(baseItem2);
					AuditedItem = baseItem2;
				}
				if (baseItem.Parent is MultiCellItem)
				{
					goto IL_00e8;
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
				if (baseItem.Parent is ThreeDRangeItem)
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
					goto IL_00e8;
				}
				B(baseItem);
				goto end_IL_0086;
				IL_0144:
				B((BaseItem)baseItem.Parent);
				goto end_IL_0086;
				IL_00e8:
				if (!(baseItem.Parent is MultiCellItem))
				{
					goto IL_0144;
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
				if (baseItem.Level <= 1 || !(baseItem.Parent.Parent is ThreeDRangeItem))
				{
					goto IL_0144;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					B((BaseItem)baseItem.Parent.Parent);
					break;
				}
				end_IL_0086:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			baseItem2 = null;
			baseItem = null;
			return;
		}
	}

	private void A(Range A)
	{
		Base.GoToCell(A, this.m_A, trvTrace, chkUnhide.IsChecked.Value, ref this.m_A);
		try
		{
			Base.MoveFormAsNeeded(Host, A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void OnSelected(object sender, RoutedEventArgs e)
	{
		if (!object.ReferenceEquals(RuntimeHelpers.GetObjectValue(sender), RuntimeHelpers.GetObjectValue(e.OriginalSource)))
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
			if (e.OriginalSource is TreeViewItem treeViewItem)
			{
				treeViewItem.BringIntoView();
				TreeViewItem treeViewItem2 = null;
			}
			return;
		}
	}

	private void OnExpanded(object sender, RoutedEventArgs e)
	{
		if (!this.m_C)
		{
			this.m_C = true;
			e.Handled = true;
			return;
		}
		BaseItem baseItem = (BaseItem)((TreeViewItem)sender).DataContext;
		if (!(baseItem is SingleCellItem))
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
			if (!(baseItem is RootItem))
			{
				if (baseItem is MultiCellItem)
				{
					A((MultiCellItem)baseItem);
				}
				else if (baseItem is FunctionItem)
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
					A((FunctionItem)baseItem);
				}
				else if (baseItem is GroupItem)
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
					A((GroupItem)baseItem);
				}
				else if (baseItem is ThreeDRangeItem)
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
					A((ThreeDRangeItem)baseItem);
				}
				goto IL_00de;
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
		C(baseItem);
		A(baseItem);
		goto IL_00de;
		IL_00de:
		e.Handled = true;
		baseItem = null;
	}

	private void OnCollapsed(object sender, RoutedEventArgs e)
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
			BaseItem baseItem = (BaseItem)((TreeViewItem)sender).DataContext;
			BaseItem baseItem2 = baseItem;
			if (baseItem2.IsSelected)
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
				if (baseItem is SingleCellItem)
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
					if (!(baseItem2.Parent is MultiCellItem))
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
						if (!(baseItem2.Parent is ThreeDRangeItem))
						{
							if (!(baseItem2.Parent is FunctionItem))
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
								if (!(baseItem2.Parent is GroupItem))
								{
									A((BaseItem)baseItem2.Parent);
									B(baseItem);
									goto IL_0179;
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
							A(A((BaseItem)baseItem2.Parent.Parent));
							B(baseItem);
							goto IL_0179;
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
					if (chkEvaluate.IsChecked == true)
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
						A(A((BaseItem)baseItem2.Parent.Parent));
					}
					else
					{
						A((BaseItem)baseItem2.Parent.Parent);
					}
					B((BaseItem)baseItem2.Parent);
				}
			}
			goto IL_0179;
			IL_0179:
			baseItem2 = null;
			e.Handled = true;
			baseItem = null;
			return;
		}
	}

	private BaseItem A(BaseItem A)
	{
		while (true)
		{
			if (!(A is FunctionItem))
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
				if (!(A is GroupItem))
				{
					break;
				}
			}
			A = (BaseItem)A.Parent;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			return A;
		}
	}

	private BaseItem B(BaseItem A)
	{
		while (true)
		{
			if (A is SingleCellItem)
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
				if (A.IsExpanded)
				{
					break;
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
			if (!(A is RootItem))
			{
				A = (BaseItem)A.Parent;
				continue;
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
			break;
		}
		return A;
	}

	private void A(BaseItem A)
	{
		try
		{
			Range range = A.Range;
			string strFormula;
			try
			{
				strFormula = NewLateBinding.LateGet(range, null, VH.A(8714), new object[0], null, null, null).ToString();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				strFormula = range.FormulaLocal.ToString();
				ProjectData.ClearProjectError();
			}
			strFormula = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(strFormula, range.Worksheet.Name);
			range = null;
			bool? isChecked = chkWrap.IsChecked;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				strFormula = Base.RemoveNewlines(strFormula);
				scroller.ScrollToHorizontalOffset(0.0);
			}
			try
			{
				this.A(A, strFormula);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				tbFormula.Text = strFormula;
				ProjectData.ClearProjectError();
			}
			try
			{
				NewLateBinding.LateCall(this.m_A.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
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
			clsReporting.LogException(ex8);
			ProjectData.ClearProjectError();
		}
	}

	private void B(BaseItem A)
	{
		bool value = chkWrap.IsChecked.Value;
		int num = 0;
		try
		{
			num = Conversions.ToInteger(A.Range.Cells.CountLarge);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		G();
		if (A.Level != 0)
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
			if (num == 1)
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
				if (A.IsExpanded)
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
					if (A is SingleCellItem)
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
						goto IL_009b;
					}
				}
			}
			if (chkEvaluate.IsChecked == true)
			{
				this.A(A, tbFormula.Inlines, value);
				return;
			}
			IEnumerator<Inline> enumerator = default(IEnumerator<Inline>);
			try
			{
				enumerator = tbFormula.Inlines.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Run run = (Run)enumerator.Current;
					if (run.Tag != A)
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
						this.A(run, value);
						return;
					}
				}
				return;
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
		}
		goto IL_009b;
		IL_009b:
		if (value)
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
			scroller.ScrollToHorizontalOffset(0.0);
			return;
		}
	}

	private void G()
	{
		if (HighlightedInline == null)
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
			if (HighlightedInline is Run)
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
				((Run)HighlightedInline).Background = null;
			}
			else
			{
				((Span)HighlightedInline).Background = null;
			}
			HighlightedInline = null;
			return;
		}
	}

	private void A(BaseItem A, InlineCollection B, bool C)
	{
		Span span;
		Run run;
		using (IEnumerator<Inline> enumerator = B.GetEnumerator())
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					object current = enumerator.Current;
					if (current is Span)
					{
						span = (Span)current;
						if (span.Tag == A)
						{
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
								this.A(span, C);
								break;
							}
							break;
						}
						if (span.Inlines.Count > 0)
						{
							this.A(A, span.Inlines, C);
						}
						continue;
					}
					run = (Run)current;
					if (run.Tag != A)
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
						this.A(run, C);
						break;
					}
					break;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_00a3;
					}
					continue;
					end_IL_00a3:
					break;
				}
				break;
			}
		}
		span = null;
		run = null;
	}

	private void A(Span A, bool B)
	{
		A.Background = A.Foreground.Clone();
		A.Background.Opacity = 0.15;
		HighlightedInline = A;
		if (B)
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
			H(A.ElementStart, A.ElementEnd);
			return;
		}
	}

	private void A(Run A, bool B)
	{
		A.Background = A.Foreground.Clone();
		A.Background.Opacity = 0.15;
		HighlightedInline = A;
		if (!B)
		{
			H(A.ElementStart, A.ElementEnd);
		}
	}

	private void H(TextPointer A, TextPointer B)
	{
		if (this.A(B))
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
					this.A(A);
					return;
				}
			}
		}
		if (this.B(A))
		{
			this.A(A);
		}
	}

	private bool A(TextPointer A)
	{
		return A.GetCharacterRect(LogicalDirection.Forward).Location.X > scroller.ActualWidth + scroller.HorizontalOffset;
	}

	private bool B(TextPointer A)
	{
		return A.GetCharacterRect(LogicalDirection.Forward).Location.X < scroller.HorizontalOffset;
	}

	private void A(TextPointer A)
	{
		scroller.ScrollToHorizontalOffset(A.GetPositionAtOffset(-2).GetCharacterRect(LogicalDirection.Forward).Location.X);
	}

	private void A(BaseItem A, string B)
	{
		List<Range> G = new List<Range>();
		List<string> H = new List<string>();
		List<int> I = new List<int>();
		List<Inline> F = new List<Inline>();
		Worksheet worksheet = A.Range.Worksheet;
		int E = 0;
		List<BaseItem> B2 = new List<BaseItem>();
		bool? isChecked = chkEvaluate.IsChecked;
		bool? flag;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		checked
		{
			int J = default(int);
			Range range;
			if (isChecked == true)
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
				this.A(A, ref B2);
				using List<BaseItem>.Enumerator enumerator = B2.GetEnumerator();
				while (enumerator.MoveNext())
				{
					BaseItem current = enumerator.Current;
					string right;
					Run A2;
					if (!(current is ThreeDRangeItem))
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
						range = current.Range;
						if (range == null)
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
						F.Add(new Run(B.Substring(E, current.SelectionIndex - E)));
						A2 = new Run(B.Substring(current.SelectionIndex, current.SelectionLength));
						E = current.SelectionIndex + current.SelectionLength;
						this.A(ref A2, current);
						if (range.Worksheet == worksheet)
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
							bool flag2 = false;
							int num = 0;
							right = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							foreach (Range item in G)
							{
								if (Operators.CompareString(item.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										flag2 = true;
										break;
									}
									break;
								}
								num++;
							}
							int b;
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
								b = I[num];
							}
							else
							{
								J++;
								b = J;
								G.Add(range);
								I.Add(J);
							}
							this.A(A2, b);
						}
						F.Add(A2);
						A2 = null;
						continue;
					}
					right = ((ThreeDRangeItem)current).Label;
					F.Add(new Run(B.Substring(E, current.SelectionIndex - E)));
					A2 = new Run(B.Substring(current.SelectionIndex, current.SelectionLength));
					E = current.SelectionIndex + current.SelectionLength;
					this.A(ref A2, current);
					if (this.A(worksheet, (ThreeDRangeItem)current))
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
						bool flag2 = false;
						int num = 0;
						using (List<string>.Enumerator enumerator3 = H.GetEnumerator())
						{
							while (true)
							{
								if (enumerator3.MoveNext())
								{
									if (Operators.CompareString(enumerator3.Current, right, TextCompare: false) == 0)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											flag2 = true;
											break;
										}
										break;
									}
									num++;
									continue;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0336;
									}
									continue;
									end_IL_0336:
									break;
								}
								break;
							}
						}
						int b;
						if (flag2)
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
							b = I[num];
						}
						else
						{
							J++;
							b = J;
							H.Add(right);
							I.Add(J);
						}
						this.A(A2, b);
					}
					F.Add(A2);
					A2 = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_03aa;
					}
					continue;
					end_IL_03aa:
					break;
				}
			}
			else
			{
				Worksheet a = worksheet;
				Dictionary<BaseItem, List<BaseItem>> K = null;
				this.A(a, A, B, null, ref E, ref F, ref G, ref H, ref I, ref J, ref K);
			}
			if (E < B.Length)
			{
				F.Add(new Run(B.Substring(E, B.Length - E)));
			}
			InlineCollection inlines = tbFormula.Inlines;
			inlines.Clear();
			inlines.AddRange(F.ToArray());
			if (Operators.CompareString(tbFormula.Text, B, TextCompare: false) != 0)
			{
				tbFormula.Text = B;
			}
			_ = null;
			range = null;
			worksheet = null;
			G = null;
			H = null;
			I = null;
			B2 = null;
			F = null;
		}
	}

	private void A(ref Run A, BaseItem B)
	{
		Run obj = A;
		obj.Tag = B;
		obj.Cursor = System.Windows.Input.Cursors.Hand;
		obj.MouseEnter += ShowHyperlink;
		obj.MouseLeave += HideHyperlink;
		obj.MouseLeftButtonDown += HyperlinkClick;
		_ = null;
	}

	private void ShowHyperlink(object sender, System.Windows.Input.MouseEventArgs e)
	{
		Inline inline = (Inline)sender;
		if (inline.TextDecorations == null)
		{
			inline.TextDecorations = new TextDecorationCollection();
		}
		inline.TextDecorations.Add(TextDecorations.Underline);
		inline = null;
	}

	private void HideHyperlink(object sender, System.Windows.Input.MouseEventArgs e)
	{
		((Inline)sender).TextDecorations = null;
	}

	private void HyperlinkClick(object sender, MouseButtonEventArgs e)
	{
		BaseItem baseItem = (BaseItem)((Inline)sender).Tag;
		if (chkEvaluate.IsChecked == true)
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
			BaseItem baseItem2;
			for (baseItem2 = (BaseItem)baseItem.Parent; baseItem2 != null; baseItem2 = (BaseItem)baseItem2.Parent)
			{
				baseItem2.IsExpanded = true;
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
			baseItem2 = null;
		}
		baseItem.IsSelected = true;
		baseItem = null;
		e.Handled = true;
		trvTrace.Focus();
	}

	private void A(BaseItem A, ref List<BaseItem> B)
	{
		if (A.Items == null)
		{
			return;
		}
		IEnumerator<TraceItem> enumerator = default(IEnumerator<TraceItem>);
		try
		{
			enumerator = A.Items.GetEnumerator();
			while (enumerator.MoveNext())
			{
				BaseItem baseItem = (BaseItem)enumerator.Current;
				if (!(baseItem is FunctionItem))
				{
					if (!(baseItem is GroupItem))
					{
						B.Add(baseItem);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
				}
				this.A(baseItem, ref B);
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
	}

	private void A(Worksheet A, BaseItem B, string C, Span D, ref int E, ref List<Inline> F, ref List<Range> G, ref List<string> H, ref List<int> I, ref int J, [Optional][DefaultParameterValue(null)] ref Dictionary<BaseItem, List<BaseItem>> K)
	{
		if (K == null)
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
			K = this.A(B);
		}
		object obj;
		if (!K.ContainsKey(B))
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
			obj = B.Items?.Cast<BaseItem>().ToList();
		}
		else
		{
			obj = K[B];
		}
		List<BaseItem> list = (List<BaseItem>)obj;
		if (list == null)
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
				using List<BaseItem>.Enumerator enumerator = list.GetEnumerator();
				while (enumerator.MoveNext())
				{
					BaseItem current = enumerator.Current;
					Run A2;
					Range range;
					if (!(current is FunctionItem))
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
						if (!(current is GroupItem))
						{
							if (current is ThreeDRangeItem)
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
								if (!(current.Parent is FunctionItem))
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
									if (!(current.Parent is GroupItem))
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
										D = null;
									}
								}
								if (D == null)
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
									F.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
								}
								else
								{
									D.Inlines.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
								}
								A2 = new Run(C.Substring(current.SelectionIndex, current.SelectionLength));
								E = current.SelectionIndex + current.SelectionLength;
								this.A(ref A2, current);
								string label = ((ThreeDRangeItem)current).Label;
								if (this.A(A, (ThreeDRangeItem)current))
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
									bool flag = false;
									int num = 0;
									using (List<string>.Enumerator enumerator2 = H.GetEnumerator())
									{
										while (true)
										{
											if (enumerator2.MoveNext())
											{
												if (Operators.CompareString(enumerator2.Current, label, TextCompare: false) == 0)
												{
													while (true)
													{
														switch (6)
														{
														case 0:
															continue;
														}
														flag = true;
														break;
													}
													break;
												}
												num++;
												continue;
											}
											while (true)
											{
												switch (2)
												{
												case 0:
													break;
												default:
													goto end_IL_0337;
												}
												continue;
												end_IL_0337:
												break;
											}
											break;
										}
									}
									int b;
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
										b = I[num];
									}
									else
									{
										J++;
										b = J;
										H.Add(label);
										I.Add(J);
									}
									this.A(A2, b);
								}
								if (D == null)
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
									F.Add(A2);
								}
								else
								{
									D.Inlines.Add(A2);
								}
								A2 = null;
							}
							else
							{
								if (D == null)
								{
									F.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
								}
								else
								{
									D.Inlines.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
								}
								A2 = new Run(C.Substring(current.SelectionIndex, current.SelectionLength));
								E = current.SelectionIndex + current.SelectionLength;
								this.A(ref A2, current);
								range = current.Range;
								if (range != null && range.Worksheet == A)
								{
									bool flag = false;
									int num = 0;
									string label = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
									using (List<Range>.Enumerator enumerator3 = G.GetEnumerator())
									{
										while (true)
										{
											if (enumerator3.MoveNext())
											{
												if (Operators.CompareString(enumerator3.Current.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), label, TextCompare: false) == 0)
												{
													while (true)
													{
														switch (2)
														{
														case 0:
															continue;
														}
														flag = true;
														break;
													}
													break;
												}
												num++;
												continue;
											}
											while (true)
											{
												switch (3)
												{
												case 0:
													break;
												default:
													goto end_IL_0518;
												}
												continue;
												end_IL_0518:
												break;
											}
											break;
										}
									}
									int b;
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
										b = I[num];
									}
									else
									{
										J++;
										b = J;
										G.Add(range);
										I.Add(J);
									}
									this.A(A2, b);
								}
								if (D == null)
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
									F.Add(A2);
								}
								else
								{
									D.Inlines.Add(A2);
								}
								A2 = null;
							}
							goto IL_05a4;
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
					if (D == null)
					{
						F.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
					}
					else
					{
						if (current.SelectionIndex < E)
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
						D.Inlines.Add(new Run(C.Substring(E, current.SelectionIndex - E)));
					}
					E = current.SelectionIndex;
					Span span = new Span();
					span.Tag = current;
					if (D == null)
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
						F.Add(span);
					}
					else
					{
						D.Inlines.Add(span);
					}
					this.A(A, current, C, span, ref E, ref F, ref G, ref H, ref I, ref J, ref K);
					span.Inlines.Add(new Run(C.Substring(E, current.SelectionIndex + current.SelectionLength - E)));
					E = current.SelectionIndex + current.SelectionLength;
					span = null;
					goto IL_05a4;
					IL_05a4:
					A2 = null;
					range = null;
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

	private Dictionary<BaseItem, List<BaseItem>> A(BaseItem A)
	{
		AD a = default(AD);
		AD CS$<>8__locals4 = new AD(a);
		CS$<>8__locals4.A = A;
		List<BaseItem> B = new List<BaseItem>();
		this.B(CS$<>8__locals4.A, ref B);
		List<SC> list = (from BaseItem a3 in B
			select new SC
			{
				A = a3,
				B = CS$<>8__locals4.A
			}).ToList();
		using (List<SC>.Enumerator enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				SC current = enumerator.Current;
				int selectionIndex = current.A.SelectionIndex;
				int selectionEnd = current.A.SelectionEnd;
				foreach (SC item in list)
				{
					if (object.Equals(current, item))
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
					BaseItem a2 = item.A;
					BaseItem b = current.B;
					int selectionIndex2 = a2.SelectionIndex;
					int selectionEnd2 = a2.SelectionEnd;
					if (selectionIndex2 > selectionIndex)
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
					if (selectionEnd2 < selectionEnd)
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
					if (!object.Equals(b, CS$<>8__locals4.A))
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
						if (a2.SelectionLength >= b.SelectionLength)
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
							if (a2.SelectionLength > b.SelectionLength)
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
							if (a2.SelectionIndex < b.SelectionIndex)
							{
								continue;
							}
						}
					}
					current.B = a2;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0198;
				}
				continue;
				end_IL_0198:
				break;
			}
		}
		Dictionary<BaseItem, List<BaseItem>> dictionary = new Dictionary<BaseItem, List<BaseItem>>();
		IEnumerator<IGrouping<BaseItem, SC>> enumerator3 = default(IEnumerator<IGrouping<BaseItem, SC>>);
		try
		{
			Func<SC, BaseItem> keySelector;
			if (_Closure$__.A == null)
			{
				keySelector = (_Closure$__.A = [SpecialName] (SC sC) => sC.B);
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
				keySelector = _Closure$__.A;
			}
			enumerator3 = list.GroupBy(keySelector).GetEnumerator();
			while (enumerator3.MoveNext())
			{
				IGrouping<BaseItem, SC> current3 = enumerator3.Current;
				BaseItem key = current3.Key;
				Func<SC, BaseItem> selector;
				if (_Closure$__.B == null)
				{
					selector = (_Closure$__.B = [SpecialName] (SC sC) => sC.A);
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
					selector = _Closure$__.B;
				}
				IEnumerable<BaseItem> source = current3.Select(selector);
				Func<BaseItem, int> keySelector2;
				if (_Closure$__.A != null)
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
					keySelector2 = _Closure$__.A;
				}
				else
				{
					keySelector2 = (_Closure$__.A = [SpecialName] (BaseItem baseItem) => baseItem.SelectionIndex);
				}
				dictionary[key] = source.OrderBy(keySelector2).ThenByDescending((_Closure$__.B == null) ? (_Closure$__.B = [SpecialName] (BaseItem baseItem) => baseItem.SelectionLength) : _Closure$__.B).ToList();
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				return dictionary;
			}
		}
		finally
		{
			if (enumerator3 != null)
			{
				while (true)
				{
					switch (7)
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

	private void B(BaseItem A, ref List<BaseItem> B)
	{
		IEnumerator<BaseItem> enumerator = default(IEnumerator<BaseItem>);
		try
		{
			enumerator = A.Items.Cast<BaseItem>().GetEnumerator();
			while (enumerator.MoveNext())
			{
				BaseItem current = enumerator.Current;
				B.Add(current);
				this.B(current, ref B);
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
	}

	private void A(Run A, int B)
	{
		int num = B % this.m_B;
		A.Foreground = checked((num > 0) ? Dialog.FormulaBrushes[num - 1] : Dialog.FormulaBrushes[this.m_B - 1]);
	}

	private bool A(Worksheet A, ThreeDRangeItem B)
	{
		if (A.Index >= B.Ranges[0].Worksheet.Index)
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
					return A.Index <= B.Ranges[checked(B.Ranges.Count - 1)].Worksheet.Index;
				}
			}
		}
		return false;
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A, ref bool B)
	{
		Close();
	}

	private void EnableFormulaWrap(object sender, RoutedEventArgs e)
	{
		tbFormula.TextWrapping = TextWrapping.Wrap;
		K.Settings.AuditFormulaWrap = true;
	}

	private void DisableFormulaWrap(object sender, RoutedEventArgs e)
	{
		tbFormula.TextWrapping = TextWrapping.NoWrap;
		K.Settings.AuditFormulaWrap = false;
	}

	private void SettingsPopupOpened(object sender, EventArgs e)
	{
		chkSettings.IsHitTestVisible = false;
		chkEvaluate.Focus();
	}

	private void SettingsPopupClosed(object sender, EventArgs e)
	{
		chkSettings.IsChecked = false;
		chkSettings.IsHitTestVisible = true;
	}

	private void CloseSettingsPopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			chkSettings.IsChecked = false;
			e.Handled = true;
			return;
		}
	}

	private void EvaluateCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditEvaluateFormulas = chkEvaluate.IsChecked.Value;
		bool? isChecked = chkEvaluate.IsChecked;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
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
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			ArgumentsColumnVisibility = Visibility.Collapsed;
			tbArguments.IsEnabled = false;
			System.Windows.Controls.CheckBox checkBox = chkArguments;
			checkBox.IsEnabled = false;
			checkBox.Unchecked -= ArgumentsCheckedChanged;
			checkBox.IsChecked = false;
			checkBox.Unchecked += ArgumentsCheckedChanged;
			_ = null;
			K.Settings.AuditEvaluateArguments = false;
		}
		else
		{
			tbArguments.IsEnabled = true;
			chkArguments.IsEnabled = true;
		}
		C();
	}

	private void ArgumentsCheckedChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = chkArguments.IsChecked;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
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
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			ArgumentsColumnVisibility = Visibility.Collapsed;
		}
		C();
		K.Settings.AuditEvaluateArguments = chkArguments.IsChecked.Value;
	}

	private void HighlightCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditHighlightCells = chkHighlight.IsChecked.Value;
		bool? isChecked = chkHighlight.IsChecked;
		bool? flag;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		if (isChecked != true)
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
			Base.RemoveHighlight();
			return;
		}
	}

	private void OpenLinksCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditOpenWorkbookLinks = chkOpenLinks.IsChecked.Value;
	}

	private void UnhideCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (chkUnhide.IsChecked == true)
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
			Range range = (Range)this.m_A.Selection;
			if (Ranges.HasHiddenCells(range))
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
				Base.D(range);
			}
			range.Select();
			range = null;
		}
		trvTrace.Focus();
		K.Settings.AuditUnhideRowsColumns = chkUnhide.IsChecked.Value;
	}

	private void ArrowsCheckedChanged(object sender, RoutedEventArgs e)
	{
	}

	private void MoveCheckedChanged(object sender, RoutedEventArgs e)
	{
		K.Settings.AuditFormMoveOnNavigate = chkMove.IsChecked.Value;
	}

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		if (this.m_B == null)
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
			this.m_B = (ScrollViewer)trvTrace.Template.FindName(VH.A(43250), trvTrace);
			if (this.m_B != null)
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
				this.m_B.ScrollChanged -= TreeViewScrollViewerScrollChanged;
				this.m_B.ScrollChanged += TreeViewScrollViewerScrollChanged;
			}
		}
		this.m_D = true;
		this.m_B = this.m_B.HorizontalOffset;
	}

	private void TreeViewScrollViewerScrollChanged(object sender, ScrollChangedEventArgs e)
	{
		if (this.m_D)
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
			this.m_B.ScrollToHorizontalOffset(this.m_B);
		}
		this.m_D = false;
	}

	private DependencyObject C(DependencyObject A)
	{
		DependencyObject parent = VisualTreeHelper.GetParent(A);
		if (!(parent is System.Windows.Controls.TreeView))
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
					return C(parent);
				}
			}
		}
		return parent;
	}

	private void A(MultiCellItem A)
	{
		List<BaseItem> list = new List<BaseItem>();
		A.Items.Clear();
		Range rng = (Range)this.m_A.Selection;
		this.m_A.ScreenUpdating = false;
		List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
		checked
		{
			try
			{
				string B = string.Empty;
				string C = string.Empty;
				if (A.IsName)
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
					BaseItem baseItem = A.FirstRangeParent();
					object a;
					if (baseItem == null)
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
						a = null;
					}
					else
					{
						a = baseItem.Range;
					}
					CD.A((Range)a, ref B, ref C);
				}
				int num = ((A.ExtraAreas != null) ? A.ExtraAreas.Count : 0);
				IEnumerator enumerator = default(IEnumerator);
				WC wC = default(WC);
				IEnumerator enumerator2 = default(IEnumerator);
				for (int i = 0; i <= num; i++)
				{
					Range range;
					if (i != 0)
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
						range = A.ExtraAreas[i - 1];
					}
					else
					{
						range = A.Range;
					}
					Range range2 = range;
					if (ExcelAddIn1.Formulas.Helpers.ContainsMergedCells(range2))
					{
						try
						{
							enumerator = range2.GetEnumerator();
							while (enumerator.MoveNext())
							{
								wC = new WC(wC);
								wC.A = (Range)enumerator.Current;
								if (Conversions.ToBoolean(wC.A.MergeCells))
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
									wC.A = (Range)wC.A.MergeArea.Cells[1, 1];
									if (list.Where(wC.A).Count() != 0)
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
									list.Add(this.A(A, wC.A, B, C));
								}
								else
								{
									list.Add(this.A(A, wC.A, B, C));
								}
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_01d2;
								}
								continue;
								end_IL_01d2:
								break;
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					else
					{
						try
						{
							enumerator2 = range2.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Range b = (Range)enumerator2.Current;
								list.Add(this.A(A, b, B, C));
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
					range2 = null;
				}
				A.Items.AddRange(list);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
			Base.ReturnToPreviousRange(rng);
			this.m_A.ScreenUpdating = true;
			list = null;
			rng = null;
			hiddenSheetNames = null;
		}
	}

	private SingleCellItem A(MultiCellItem A, Range B, string C, string D)
	{
		string strLabel = CD.A(B, C, D);
		BaseItem baseItem = new SingleCellItem(A, B, strLabel, this.B(B));
		if (this.A(B))
		{
			baseItem.Items.Add(new DummyItem(baseItem));
		}
		return (SingleCellItem)baseItem;
	}

	private void A(ThreeDRangeItem A)
	{
		A.Items.Clear();
		Range rng;
		List<string> hiddenSheetNames;
		try
		{
			rng = (Range)this.m_A.Selection;
			hiddenSheetNames = Base.UnhideHiddenSheets();
			try
			{
				using List<Range>.Enumerator enumerator = A.Ranges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range current = enumerator.Current;
					string strLabel = CD.A("", current.Worksheet.Name, current.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
					bool flag = false;
					long num = Conversions.ToLong(current.Cells.CountLarge);
					BaseItem baseItem;
					if (num == 1)
					{
						baseItem = new SingleCellItem(A, current, strLabel, B(current));
						try
						{
							this.m_A.ScreenUpdating = false;
							flag = this.A(current);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
					}
					else
					{
						baseItem = new MultiCellItem(A, current, strLabel);
						flag = num <= this.m_A;
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						baseItem.Items.Add(new DummyItem(baseItem));
					}
					A.Items.Add(baseItem);
					baseItem = null;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_014b;
					}
					continue;
					end_IL_014b:
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				E(ex4.Message);
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
			if (!this.m_A.ScreenUpdating)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					Base.ReturnToPreviousRange(rng);
					this.m_A.ScreenUpdating = true;
					break;
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		rng = null;
		hiddenSheetNames = null;
	}

	private void A(FunctionItem A)
	{
	}

	private void A(GroupItem A)
	{
	}

	private void C(BaseItem A)
	{
		//IL_04b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_04be: Invalid comparison between Unknown and I4
		XC a = default(XC);
		XC CS$<>8__locals38 = new XC(a);
		Microsoft.Office.Interop.Excel.Application application = A.Range.Application;
		List<Precedent> B = new List<Precedent>();
		List<ParenthesesPair> list = new List<ParenthesesPair>();
		bool flag = false;
		bool F = false;
		List<BaseItem> E = new List<BaseItem>();
		List<Argument> list2 = null;
		A.Items.Clear();
		Range range = A.Range;
		Range rng = (Range)application.Selection;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		application2.DisplayAlerts = false;
		XlCalculation calculation = application2.Calculation;
		bool iteration = application2.Iteration;
		application2.Calculation = XlCalculation.xlCalculationManual;
		application2.Iteration = false;
		application2.ScreenUpdating = false;
		_ = null;
		List<string> hiddenSheetNames = Base.UnhideHiddenSheets();
		Range range2 = range;
		if (K.Settings.AuditOpenWorkbookLinks)
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
			Host.IgnoreDeactivate = true;
			try
			{
				this.B(range2.Formula.ToString());
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Host.IgnoreDeactivate = false;
		}
		string strFormula;
		try
		{
			strFormula = NewLateBinding.LateGet(range2, null, VH.A(8714), new object[0], null, null, null).ToString();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			strFormula = range2.FormulaLocal.ToString();
			ProjectData.ClearProjectError();
		}
		bool? isChecked = chkWrap.IsChecked;
		if (((!isChecked) ?? isChecked) == true)
		{
			strFormula = Base.RemoveNewlines(strFormula);
		}
		Worksheet worksheet = range2.Worksheet;
		range2 = null;
		CS$<>8__locals38.A = ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(strFormula, worksheet.Name);
		strFormula = CS$<>8__locals38.A;
		try
		{
			if (CS$<>8__locals38.A.StartsWith(VH.A(47800)))
			{
				Match match = Regex.Matches(CS$<>8__locals38.A, VH.A(47813))[0];
				Group obj = match.Groups[1];
				if (obj.ToString().Length > 0)
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
					Precedent item = new Precedent
					{
						Address = obj.ToString(),
						Value = Conversions.ToString(((_Worksheet)worksheet).get_Range((object)obj.ToString(), RuntimeHelpers.GetObjectValue(Missing.Value)).Text)
					};
					B.Add(item);
					item = null;
				}
				obj = null;
				Group obj2 = match.Groups[2];
				if (obj2.ToString().Length > 0)
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
					Precedent item = new Precedent();
					item.Address = obj2.ToString();
					item.Value = Conversions.ToString(((_Worksheet)worksheet).get_Range((object)obj2.ToString(), RuntimeHelpers.GetObjectValue(Missing.Value)).Text);
					B.Add(item);
					item = null;
				}
				obj2 = null;
				_ = null;
			}
			else
			{
				try
				{
					B = this.A(range);
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					clsReporting.LogException(ex6);
					ProjectData.ClearProjectError();
				}
			}
			application.Goto(range, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		CS$<>8__locals38.A = new List<Precedent>();
		XC xC = CS$<>8__locals38;
		List<Precedent> source = B;
		Func<Precedent, string> selector;
		if (_Closure$__.A == null)
		{
			selector = (_Closure$__.A = [SpecialName] (Precedent precedent) => precedent.Address);
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
			selector = _Closure$__.A;
		}
		xC.A = source.Select(selector).ToList();
		B.ForEach(checked([SpecialName] (Precedent precedent) =>
		{
			int num16 = precedent.Address.IndexOf(VH.A(2826));
			if (num16 >= 0)
			{
				string text4 = precedent.Address.Substring(0, num16);
				string text5 = precedent.Address.Substring(num16 + 1);
				if (text4.Contains(VH.A(7827)) && !text5.Contains(VH.A(7827)))
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
					text5 = string.Format(VH.A(49936), text4.Substring(0, text4.LastIndexOf('!') + 1), text5);
				}
				Precedent.AddProvisionalPrec(text4, precedent, ref CS$<>8__locals38.A, ref CS$<>8__locals38.A);
				Precedent.AddProvisionalPrec(text5, precedent, ref CS$<>8__locals38.A, ref CS$<>8__locals38.A);
			}
		}));
		B.AddRange(CS$<>8__locals38.A);
		try
		{
			List<Precedent> list3 = B;
			Action<Precedent> action;
			if (_Closure$__.A == null)
			{
				action = (_Closure$__.A = [SpecialName] (Precedent precedent) =>
				{
					if (!precedent.Address.Contains(VH.A(7827)))
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
								precedent.Address = VH.A(211957) + precedent.Address;
								return;
							}
						}
					}
				});
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
				action = _Closure$__.A;
			}
			list3.ForEach(action);
			List<Precedent> list4 = B;
			Comparison<Precedent> comparison;
			if (_Closure$__.A == null)
			{
				comparison = (_Closure$__.A = [SpecialName] (Precedent precedent, Precedent precedent2) => precedent.Address.CompareTo(precedent2.Address));
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
				comparison = _Closure$__.A;
			}
			list4.Sort(comparison);
			List<Precedent> list5 = B;
			Action<Precedent> action2;
			if (_Closure$__.B == null)
			{
				action2 = (_Closure$__.B = [SpecialName] (Precedent precedent) =>
				{
					precedent.Address = precedent.Address.Replace(VH.A(211957), "");
				});
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
				action2 = _Closure$__.B;
			}
			list5.ForEach(action2);
			B = B.Distinct(new PrecedentComparer()).ToList();
			List<Precedent> list6 = B;
			Comparison<Precedent> comparison2;
			if (_Closure$__.B == null)
			{
				comparison2 = (_Closure$__.B = [SpecialName] (Precedent precedent2, Precedent precedent) => precedent.Address.Length.CompareTo(precedent2.Address.Length));
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
				comparison2 = _Closure$__.B;
			}
			list6.Sort(comparison2);
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			ProjectData.ClearProjectError();
		}
		Helpers.MaskQuotedText(ref CS$<>8__locals38.A);
		int num;
		if ((int)clsEnvironment.ApplicationLanguage == 1)
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
			num = ((Operators.CompareString(CultureInfo.CurrentCulture.TextInfo.ListSeparator, VH.A(2378), TextCompare: false) != 0) ? 1 : 0);
		}
		else
		{
			num = 1;
		}
		flag = (byte)num != 0;
		bool value = chkEvaluate.IsChecked.Value;
		checked
		{
			List<int> list7 = default(List<int>);
			if (value)
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
				string strFormula2 = CS$<>8__locals38.A;
				int length = strFormula2.Length;
				list7 = new List<int>(length);
				length--;
				int num2 = length;
				for (int num3 = 0; num3 <= num2; num3++)
				{
					list7.Add(0);
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
				if (strFormula2.Contains(VH.A(7827)))
				{
					foreach (Precedent item2 in B)
					{
						Helpers.MaskSheetAndWorkbookNames(ref strFormula2, item2.Address);
					}
				}
				list = Helpers.IdentifyParenthesesPairs(strFormula2);
				list2 = new List<Argument>();
				bool value2 = chkArguments.IsChecked.Value;
				foreach (ParenthesesPair item3 in list)
				{
					try
					{
						if (item3.FunctionName.Length > 0)
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
							E.Add(this.A(A, item3, range, worksheet, strFormula, flag));
							if (value2)
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
								list2.AddRange(this.A(A, item3, worksheet, strFormula, flag));
							}
						}
						else
						{
							E.Add(this.A(A, item3, worksheet, strFormula, flag));
						}
						length = item3.StartIndex + item3.Length - 1;
						int startIndex = item3.StartIndex;
						int num4 = length;
						for (int num5 = startIndex; num5 <= num4; num5++)
						{
							list7[num5]++;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_06d4;
							}
							continue;
							end_IL_06d4:
							break;
						}
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception ex12 = ex11;
						ProjectData.ClearProjectError();
					}
				}
				list = null;
			}
			MatchCollection matchCollection;
			if (CS$<>8__locals38.A.Contains(VH.A(7827)) && CS$<>8__locals38.A.Contains(VH.A(2826)) && CS$<>8__locals38.A.Contains(VH.A(39848)))
			{
				matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(47850) + Base.CELL_REF_PATTERN + VH.A(48128));
				IEnumerator enumerator3 = default(IEnumerator);
				try
				{
					enumerator3 = matchCollection.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						Match match2 = (Match)enumerator3.Current;
						try
						{
							Microsoft.Office.Interop.Excel.Sheets worksheets = this.m_B.Worksheets;
							Conversions.ToInteger(Operators.AddObject(NewLateBinding.LateGet(worksheets.get_Item((object)match2.Groups[3].Value), null, VH.A(48135), new object[0], null, null, null), NewLateBinding.LateGet(worksheets.get_Item((object)match2.Groups[4].Value), null, VH.A(48135), new object[0], null, null, null)));
							worksheets = null;
						}
						catch (Exception ex13)
						{
							ProjectData.SetProjectError(ex13);
							Exception ex14 = ex13;
							ProjectData.ClearProjectError();
							continue;
						}
						Group obj3 = match2.Groups[2];
						ThreeDRangeItem threeDRangeItem = new ThreeDRangeItem(A, obj3.Value, this.m_B);
						threeDRangeItem.SelectionIndex = obj3.Index;
						threeDRangeItem.SelectionLength = obj3.Length;
						obj3 = null;
						threeDRangeItem.Items.Add(new DummyItem(threeDRangeItem));
						E.Add(threeDRangeItem);
						threeDRangeItem = null;
						CS$<>8__locals38.A = Helpers.MaskFormula(match2, CS$<>8__locals38.A);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_08ea;
						}
						continue;
						end_IL_08ea:
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
				matchCollection = null;
			}
			if (CS$<>8__locals38.A.Contains(VH.A(48146)))
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
				matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(48149) + Base.CELL_REF_PATTERN + VH.A(39904));
				this.A(A, matchCollection, worksheet, ref CS$<>8__locals38.A, ref E, ref F);
				matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(48200) + ExcelAddIn1.Formulas.Names.NAME_PATTERN + VH.A(39904), RegexOptions.IgnoreCase);
				this.A(A, matchCollection, worksheet, ref CS$<>8__locals38.A, ref E, ref F);
			}
			new VC(CS$<>8__locals38.A, ref B, worksheet).A();
			_ = null;
			foreach (Precedent item4 in B)
			{
				if (item4.IsProvAndParentMatched() || this.A(item4.Address, item4.Value, range, A, ref E, ref CS$<>8__locals38.A, ref F))
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
				List<Precedent> provisionalChildren = item4.ProvisionalChildren;
				if (provisionalChildren == null)
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
					continue;
				}
				Action<Precedent> action3;
				if (_Closure$__.C == null)
				{
					action3 = (_Closure$__.C = [SpecialName] (Precedent precedent) =>
					{
						precedent.ProvisionalParentMismatched = true;
					});
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
					action3 = _Closure$__.C;
				}
				provisionalChildren.ForEach(action3);
			}
			using (List<Precedent>.Enumerator enumerator5 = B.Where([SpecialName] (Precedent precedent) => precedent.IsProvAndParentMatched()).ToList().GetEnumerator())
			{
				while (enumerator5.MoveNext())
				{
					Precedent item = enumerator5.Current;
					B.Remove(item);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0b05;
					}
					continue;
					end_IL_0b05:
					break;
				}
			}
			matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(48205));
			IEnumerator enumerator6 = default(IEnumerator);
			try
			{
				enumerator6 = matchCollection.GetEnumerator();
				while (enumerator6.MoveNext())
				{
					Match m = (Match)enumerator6.Current;
					CS$<>8__locals38.A = Helpers.MaskFormula(m, CS$<>8__locals38.A);
				}
			}
			finally
			{
				if (enumerator6 is IDisposable)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						(enumerator6 as IDisposable).Dispose();
						break;
					}
				}
			}
			Range range3;
			if (CS$<>8__locals38.A.Contains(VH.A(7120)))
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
				string text = VH.A(48236) + Regex.Escape(VH.A(48247)) + VH.A(48250);
				string text2 = VH.A(48267);
				string arg = string.Format(VH.A(48282), RuntimeHelpers.GetObjectValue(((_Application)application).get_International((object)XlApplicationInternational.xlListSeparator)));
				string[] array = new string[16]
				{
					text2,
					VH.A(48146),
					string.Format(VH.A(48289), text2),
					VH.A(48306),
					VH.A(48315),
					VH.A(48326),
					VH.A(48341),
					string.Format(VH.A(48358), arg, text2),
					string.Format(VH.A(48395), arg, text2),
					string.Format(VH.A(48434), arg, text2),
					string.Format(VH.A(48477), arg, text2),
					string.Format(VH.A(48522), text2, text2),
					string.Format(VH.A(48557), arg, text2, text2),
					string.Format(VH.A(48610), arg, text2, text2),
					string.Format(VH.A(48665), arg, text2, text2),
					string.Format(VH.A(48724), arg, text2, text2)
				};
				IEnumerator enumerator7 = default(IEnumerator);
				foreach (string text3 in array)
				{
					matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(48785) + text + VH.A(48792) + text3 + VH.A(48797), RegexOptions.IgnoreCase);
					try
					{
						try
						{
							enumerator7 = matchCollection.GetEnumerator();
							while (enumerator7.MoveNext())
							{
								Match match3 = (Match)enumerator7.Current;
								object objectValue = RuntimeHelpers.GetObjectValue(worksheet.Evaluate(this.A(string.Format(VH.A(48804), match3.Groups[1].Value), flag)));
								if (objectValue is Range)
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
									range3 = (Range)objectValue;
									int num7 = Conversions.ToInteger(range3.Cells.CountLarge);
									if (num7 == 1)
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
										E.Add(this.A(A, range3, match3.Groups[1].Value, this.A((Range)objectValue), match3.Index, match3.Length));
										F = true;
									}
									else
									{
										E.Add(this.A(A, range3, match3.Groups[1].Value, match3.Index, match3.Length, num7));
									}
									range3 = null;
									CS$<>8__locals38.A = Helpers.MaskFormula(match3, CS$<>8__locals38.A);
								}
								objectValue = null;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0f3c;
								}
								continue;
								end_IL_0f3c:
								break;
							}
						}
						finally
						{
							if (enumerator7 is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									(enumerator7 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					catch (Exception ex15)
					{
						ProjectData.SetProjectError(ex15);
						Exception ex16 = ex15;
						ProjectData.ClearProjectError();
					}
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
			matchCollection = Regex.Matches(CS$<>8__locals38.A, VH.A(39848) + ExcelAddIn1.Formulas.Names.NAME_PATTERN + VH.A(39904), RegexOptions.IgnoreCase);
			List<string> list8;
			if (matchCollection.Count > 0)
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
				IEnumerable<Match> source2 = matchCollection.OfType<Match>().Where([SpecialName] (Match match5) =>
				{
					string source5 = CS$<>8__locals38.A.Substring(0, match5.Index);
					Func<char, bool> predicate;
					if (_Closure$__.A == null)
					{
						predicate = (_Closure$__.A = _Closure$__.A.A);
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
						predicate = _Closure$__.A;
					}
					return unchecked(source5.Count(predicate) % 2) == 0;
				});
				Func<Match, string> selector2;
				if (_Closure$__.A == null)
				{
					selector2 = (_Closure$__.A = [SpecialName] (Match match5) => match5.Value);
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
					selector2 = _Closure$__.A;
				}
				list8 = source2.Select(selector2).Distinct().ToList();
				List<string> list9 = list8;
				Comparison<string> comparison3;
				if (_Closure$__.A == null)
				{
					comparison3 = (_Closure$__.A = [SpecialName] (string text5, string text4) => text4.Length.CompareTo(text5.Length));
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
					comparison3 = _Closure$__.A;
				}
				list9.Sort(comparison3);
				using List<string>.Enumerator enumerator8 = list8.GetEnumerator();
				IEnumerator enumerator9 = default(IEnumerator);
				while (enumerator8.MoveNext())
				{
					string current2 = enumerator8.Current;
					List<Range> D = null;
					try
					{
						if (!BD.A(current2, B.Select([SpecialName] (Precedent precedent) => precedent.Address), application, ref D))
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_10c8;
								}
								continue;
								end_IL_10c8:
								break;
							}
							continue;
						}
						matchCollection = Regex.Matches(CS$<>8__locals38.A, current2, RegexOptions.None);
						try
						{
							enumerator9 = matchCollection.GetEnumerator();
							while (enumerator9.MoveNext())
							{
								Match match4 = (Match)enumerator9.Current;
								try
								{
									range3 = D[0];
									int num7 = Conversions.ToInteger(range3.Cells.CountLarge);
									if (num7 != 1)
									{
										goto IL_118e;
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
									if (D.Count != 1)
									{
										goto IL_118e;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										E.Add(this.A(A, range3, current2, Base.CleanValueText(Conversions.ToString(range3.Text)), match4.Index, match4.Length, G: true));
										F = true;
										break;
									}
									goto end_IL_1109;
									IL_118e:
									E.Add(this.A(A, range3, current2, match4.Index, match4.Length, num7, G: true, D.Skip(1).ToList()));
									end_IL_1109:;
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									ProjectData.ClearProjectError();
								}
								CS$<>8__locals38.A = Helpers.MaskFormula(match4, CS$<>8__locals38.A);
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_11f3;
								}
								continue;
								end_IL_11f3:
								break;
							}
						}
						finally
						{
							if (enumerator9 is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator9 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					catch (Exception projectError2)
					{
						ProjectData.SetProjectError(projectError2);
						ProjectData.ClearProjectError();
					}
					finally
					{
						if (D != null)
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
							D.Clear();
						}
						D = null;
						range3 = null;
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_1257;
					}
					continue;
					end_IL_1257:
					break;
				}
			}
			try
			{
				NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(1630), new object[0], null, null, null, IgnoreReturn: true);
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				ProjectData.ClearProjectError();
			}
			if (E.Any())
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
				if (list2 != null)
				{
					using (List<Argument>.Enumerator enumerator10 = list2.GetEnumerator())
					{
						YC yC = default(YC);
						ZC zC = default(ZC);
						while (enumerator10.MoveNext())
						{
							yC = new YC(yC);
							yC.A = enumerator10.Current;
							bool flag2 = false;
							using (List<BaseItem>.Enumerator enumerator11 = E.GetEnumerator())
							{
								while (true)
								{
									if (enumerator11.MoveNext())
									{
										BaseItem current3 = enumerator11.Current;
										if (current3.SelectionIndex != yC.A.Index)
										{
											continue;
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
										if (current3.SelectionLength == yC.A.Length)
										{
											current3.Info = yC.A.Name;
											flag2 = true;
											break;
										}
										continue;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_1360;
										}
										continue;
										end_IL_1360:
										break;
									}
									break;
								}
							}
							if (flag2)
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
							GroupItem groupItem = new GroupItem(A, yC.A.Label, yC.A.Value);
							groupItem.SelectionIndex = yC.A.Index;
							groupItem.SelectionLength = yC.A.Length;
							groupItem.Info = yC.A.Name;
							E.Add(groupItem);
							groupItem = null;
							if (yC.A.Length > 0)
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
								zC = new ZC(zC);
								zC.A = yC;
								zC.A = zC.A.A.Index - 1 + zC.A.A.Length;
								List<BaseItem> list10 = E.Where(zC.A).ToList();
								if (list10.Count <= 0)
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
								Func<BaseItem, int> selector3;
								if (_Closure$__.C == null)
								{
									selector3 = (_Closure$__.C = [SpecialName] (BaseItem baseItem2) => baseItem2.SelectionIndex);
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
									selector3 = _Closure$__.C;
								}
								int num8 = list10.Min(selector3);
								int num9 = list10.Max([SpecialName] (BaseItem baseItem2) => baseItem2.SelectionEnd);
								int num10 = num9;
								for (int num11 = num8; num11 <= num10; num11++)
								{
									list7[num11]++;
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
							else
							{
								list7[yC.A.Index]++;
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_1554;
							}
							continue;
							end_IL_1554:
							break;
						}
					}
					if (list2.Any())
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
						ArgumentsColumnVisibility = Visibility.Visible;
					}
				}
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
					List<BaseItem> source3 = E;
					Func<BaseItem, int> keySelector;
					if (_Closure$__.E == null)
					{
						keySelector = (_Closure$__.E = [SpecialName] (BaseItem baseItem2) => baseItem2.SelectionIndex);
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
						keySelector = _Closure$__.E;
					}
					IOrderedEnumerable<BaseItem> source4 = source3.OrderBy(keySelector);
					Func<BaseItem, int> keySelector2;
					if (_Closure$__.F == null)
					{
						keySelector2 = (_Closure$__.F = [SpecialName] (BaseItem baseItem2) => baseItem2.SelectionLength);
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
						keySelector2 = _Closure$__.F;
					}
					E = source4.ThenByDescending(keySelector2).ToList();
					Dictionary<int, BaseItem> dictionary = new Dictionary<int, BaseItem>();
					dictionary.Add(0, A);
					BaseItem value4;
					try
					{
						using List<BaseItem>.Enumerator enumerator12 = E.GetEnumerator();
						while (enumerator12.MoveNext())
						{
							BaseItem current4 = enumerator12.Current;
							int num12;
							if (!(current4 is FunctionItem))
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
								num12 = ((current4 is GroupItem) ? 1 : 0);
							}
							else
							{
								num12 = 1;
							}
							bool flag3;
							int length;
							unchecked
							{
								flag3 = (byte)num12 != 0;
								if (flag3)
								{
									BaseItem baseItem = current4;
									if (baseItem.Range == null)
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
										string value3 = baseItem.Value;
										try
										{
											if (Dialog.FormulaErrors.TryGetValue((Dialog.CVErrEnum)Conversions.ToInteger(baseItem.Value), out value3))
											{
												while (true)
												{
													switch (2)
													{
													case 0:
														continue;
													}
													baseItem.Value = value3;
													baseItem.IsError = true;
													break;
												}
											}
										}
										catch (Exception ex19)
										{
											ProjectData.SetProjectError(ex19);
											Exception ex20 = ex19;
											ProjectData.ClearProjectError();
										}
									}
									baseItem = null;
								}
								length = list7[current4.SelectionIndex];
							}
							int num13 = dictionary.Count - 1;
							if (flag3)
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
								for (int num14 = num13; num14 >= 0; num14 += -1)
								{
									value4 = dictionary.ElementAt(num14).Value;
									if (length <= value4.Level - A.Level)
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
										break;
									}
									current4.Parent = value4;
									current4.Level = value4.Level + 1;
									value4.Items.Add(current4);
									break;
								}
							}
							else
							{
								int num15 = num13;
								while (true)
								{
									if (num15 >= 0)
									{
										value4 = dictionary.ElementAt(num15).Value;
										if (length >= value4.Level - A.Level)
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
											current4.Parent = value4;
											current4.Level = value4.Level + 1;
											value4.Items.Add(current4);
											break;
										}
										num15 += -1;
										continue;
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
									break;
								}
							}
							if (!flag3)
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
							if (!dictionary.ContainsKey(length))
							{
								dictionary.Add(length, current4);
							}
							else
							{
								dictionary[length] = current4;
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_183b;
							}
							continue;
							end_IL_183b:
							break;
						}
					}
					catch (Exception ex21)
					{
						ProjectData.SetProjectError(ex21);
						Exception ex22 = ex21;
						this.E(ex22.Message);
						clsReporting.LogException(ex22);
						ProjectData.ClearProjectError();
					}
					dictionary = null;
					value4 = null;
				}
				else
				{
					E.Sort([SpecialName] (BaseItem baseItem2, BaseItem baseItem3) => baseItem2.SelectionIndex.CompareTo(baseItem3.SelectionIndex));
					A.Items.AddRange(E);
				}
				if (F)
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
					Base.ReturnToPreviousRange(rng);
				}
			}
			Base.HidePreviouslyHiddenSheets(hiddenSheetNames);
			hiddenSheetNames = null;
			try
			{
				Microsoft.Office.Interop.Excel.Application application3 = application;
				application3.Iteration = iteration;
				application3.Calculation = calculation;
				application3.DisplayAlerts = true;
				application3.ScreenUpdating = true;
				_ = null;
			}
			catch (Exception ex23)
			{
				ProjectData.SetProjectError(ex23);
				Exception ex24 = ex23;
				ProjectData.ClearProjectError();
			}
			if (range != null)
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
				if (chkArrows.IsChecked == true)
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
					range.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
			application = null;
			B = null;
			list2 = null;
			list8 = null;
			list7 = null;
			E = null;
			worksheet = null;
			range3 = null;
			rng = null;
			range = null;
			matchCollection = null;
		}
	}

	private bool A(string A, string B, Range C, BaseItem D, ref List<BaseItem> E, ref string F, ref bool G)
	{
		bool result = false;
		checked
		{
			try
			{
				string input = A;
				input = Regex.Replace(input, VH.A(48813), VH.A(48824));
				input = Regex.Replace(input, VH.A(48835), VH.A(48824));
				if (input.Contains(VH.A(7827)))
				{
					string expression = Strings.Left(input, Strings.InStrRev(input, VH.A(7827), -1, CompareMethod.Text));
					expression = Strings.Replace(expression, VH.A(43088), null);
					input = expression + Strings.Right(input, input.Length - Strings.InStrRev(input, VH.A(7827)));
				}
				input = Regex.Replace(input, VH.A(48860), VH.A(48917));
				input = Strings.Replace(input, VH.A(48924), VH.A(48931));
				MatchCollection matchCollection = Regex.Matches(F, input);
				if (matchCollection.Count == 0)
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
					if (Strings.InStr(input, VH.A(7827)) > 0)
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
						string expression = Strings.Left(input, Strings.InStrRev(input, VH.A(7827), -1, CompareMethod.Text) - 1);
						expression = VH.A(39851) + Strings.Replace(expression, VH.A(39851), VH.A(39854)) + VH.A(43343);
						input = expression + Strings.Right(input, input.Length - Strings.InStrRev(input, VH.A(7827)));
						matchCollection = Regex.Matches(F, input);
					}
				}
				result = matchCollection.Count > 0;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = matchCollection.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Match match = (Match)enumerator.Current;
						try
						{
							Range range = Base.ResolveAddress(match.Value, C);
							int num = Conversions.ToInteger(range.Cells.CountLarge);
							if (num == 1)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									string text;
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
										text = Base.CleanValueText(B);
									}
									else
									{
										text = this.B(range);
									}
									B = text;
									E.Add(this.A(D, range, match.Value, B, match.Index, match.Length));
									G = true;
									break;
								}
							}
							else
							{
								E.Add(this.A(D, range, match.Value, match.Index, match.Length, num));
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						finally
						{
							Range range = null;
						}
						F = Helpers.MaskFormula(match, F);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0305;
						}
						continue;
						end_IL_0305:
						break;
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
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private SingleCellItem A(BaseItem A, Range B, string C, string D, int E, int F, bool G = false)
	{
		SingleCellItem singleCellItem = new SingleCellItem(A, B, C, D)
		{
			SelectionIndex = E,
			SelectionLength = F,
			IsName = G
		};
		if (this.A(B))
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
			singleCellItem.Items.Add(new DummyItem(singleCellItem));
		}
		return singleCellItem;
	}

	private MultiCellItem A(BaseItem A, Range B, string C, int D, int E, int F, bool G = false, List<Range> H = null)
	{
		MultiCellItem multiCellItem = new MultiCellItem(A, B, C)
		{
			SelectionIndex = D,
			SelectionLength = E,
			IsName = G,
			ExtraAreas = H
		};
		if (F <= this.m_A)
		{
			multiCellItem.Items.Add(new DummyItem(multiCellItem));
		}
		return multiCellItem;
	}

	private FunctionItem A(BaseItem A, ParenthesesPair B, Range C, Worksheet D, string E, bool F)
	{
		bool flag = false;
		ParenthesesPair parenthesesPair = B;
		object A2;
		if (Conversions.ToBoolean(C.HasArray))
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
			A2 = RuntimeHelpers.GetObjectValue(C.Text);
		}
		else
		{
			string a = VH.A(48936) + E.Substring(parenthesesPair.StartIndex, parenthesesPair.Length);
			string A3 = this.A(a, F);
			A2 = RuntimeHelpers.GetObjectValue(this.A(ref A3, D, parenthesesPair.FunctionName));
			try
			{
				this.A(ref A2, A3, parenthesesPair.FunctionName, D);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		string strValue = default(string);
		Range rng;
		if (A2 is Range)
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
			rng = (Range)A2;
			strValue = this.A((Range)A2);
		}
		else
		{
			rng = null;
			if (A2 is double)
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
					DateTime dateTime = DateTime.FromOADate(Conversions.ToDouble(A2));
					string functionName = parenthesesPair.FunctionName;
					uint num = TH.A(functionName);
					if (num <= 1600227931)
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
						if (num != 632323358)
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
							switch (num)
							{
							default:
								goto end_IL_00f1;
							case 1490570733u:
								if (Operators.CompareString(functionName, VH.A(47036), TextCompare: false) == 0)
								{
									break;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_01d7;
									}
									continue;
									end_IL_01d7:
									break;
								}
								goto end_IL_00f1;
							case 1600227931u:
								if (Operators.CompareString(functionName, VH.A(48974), TextCompare: false) == 0)
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
									break;
								}
								goto end_IL_00f1;
							}
							goto IL_02bb;
						}
						if (Operators.CompareString(functionName, VH.A(4504), TextCompare: false) == 0)
						{
							goto IL_02bb;
						}
					}
					else if (num <= 2867742231u)
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
						if (num != 2528864532u)
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
							if (num != 2867742231u)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_018b;
									}
									continue;
									end_IL_018b:
									break;
								}
							}
							else
							{
								if (Operators.CompareString(functionName, VH.A(4497), TextCompare: false) == 0)
								{
									goto IL_02bb;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_024c;
									}
									continue;
									end_IL_024c:
									break;
								}
							}
						}
						else
						{
							if (Operators.CompareString(functionName, VH.A(48939), TextCompare: false) == 0)
							{
								goto IL_02bb;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_01ff;
								}
								continue;
								end_IL_01ff:
								break;
							}
						}
					}
					else if (num != 3221746841u)
					{
						if (num != 3253468938u)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_01af;
								}
								continue;
								end_IL_01af:
								break;
							}
						}
						else
						{
							if (Operators.CompareString(functionName, VH.A(48959), TextCompare: false) == 0)
							{
								goto IL_02bb;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_028d;
								}
								continue;
								end_IL_028d:
								break;
							}
						}
					}
					else
					{
						if (Operators.CompareString(functionName, VH.A(48950), TextCompare: false) == 0)
						{
							goto IL_02bb;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0227;
							}
							continue;
							end_IL_0227:
							break;
						}
					}
					goto end_IL_00f1;
					IL_02bb:
					strValue = dateTime.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
					flag = true;
					end_IL_00f1:;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			if (!flag)
			{
				strValue = this.A(RuntimeHelpers.GetObjectValue(A2));
			}
		}
		FunctionItem result = new FunctionItem(A, rng, parenthesesPair.FunctionName + VH.A(48999), strValue)
		{
			SelectionIndex = parenthesesPair.StartIndex,
			SelectionLength = parenthesesPair.Length
		};
		parenthesesPair = null;
		A2 = null;
		rng = null;
		return result;
	}

	private List<Argument> A(BaseItem A, ParenthesesPair B, Worksheet C, string D, bool E)
	{
		List<Argument> list = new List<Argument>();
		ParenthesesPair parenthesesPair = B;
		checked
		{
			try
			{
				if (Dialog.ArgumentNames.ContainsKey(parenthesesPair.FunctionName))
				{
					string value2 = default(string);
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
						Match match = Regex.Match(VH.A(48936) + D.Substring(parenthesesPair.StartIndex, parenthesesPair.Length), VH.A(4544) + parenthesesPair.FunctionName + VH.A(49010), RegexOptions.IgnoreCase);
						if (match.Success)
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
							string value = match.Groups[1].Value;
							Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
							string text = string.Format(VH.A(48282), RuntimeHelpers.GetObjectValue(((_Application)application).get_International((object)XlApplicationInternational.xlListSeparator)));
							string[] a = BD.A(value, Conversions.ToChar(text)).ToArray();
							Range range = Ranges.FirstBlankCell(C);
							int num = parenthesesPair.StartIndex + parenthesesPair.FunctionName.Length + 1;
							List<string> list2 = this.A(a, value, 3, text);
							int num2 = list2.Count - 1;
							for (int i = 0; i <= num2; i++)
							{
								string text2 = list2[i];
								int length = text2.Length;
								string name = Dialog.ArgumentNames[parenthesesPair.FunctionName][i];
								if (Regex.IsMatch(text2, VH.A(49027)))
								{
									list.Add(new Argument
									{
										Name = name,
										Label = text2,
										Value = text2,
										Index = num,
										Length = length
									});
								}
								else if (text2.Length == 0)
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
									list.Add(new Argument
									{
										Name = name,
										Label = VH.A(49038),
										Value = "",
										Index = num,
										Length = length
									});
								}
								else
								{
									try
									{
										string A2 = this.A(text2, E);
										Range range2 = range;
										range2.Formula = string.Format(VH.A(48804), Strings.Trim(A2));
										object objectValue = RuntimeHelpers.GetObjectValue(range2.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)));
										_ = null;
										object A3 = RuntimeHelpers.GetObjectValue(this.A(ref A2, C, parenthesesPair.FunctionName));
										try
										{
											this.A(ref A3, A2, parenthesesPair.FunctionName, C);
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											clsReporting.LogException(ex2);
											ProjectData.ClearProjectError();
										}
										value2 = ((!(A3 is Range)) ? this.A(RuntimeHelpers.GetObjectValue(objectValue)) : this.A((Range)A3));
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
									finally
									{
										object A3 = null;
										object objectValue = null;
									}
									list.Add(new Argument
									{
										Name = name,
										Label = text2,
										Value = value2,
										Index = num,
										Length = length
									});
								}
								num += list2[i].Length + 1;
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
							range.Value2 = "";
							range = null;
						}
						match = null;
						break;
					}
				}
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			finally
			{
				Microsoft.Office.Interop.Excel.Application application = null;
			}
			parenthesesPair = null;
			return list;
		}
	}

	private object A(ref string A, Worksheet B, string C)
	{
		checked
		{
			object result;
			if (!A.StartsWith(VH.A(49057)))
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
				if (!A.Contains(VH.A(49076)))
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
					result = RuntimeHelpers.GetObjectValue(B.Evaluate(A));
					if (Operators.CompareString(C, VH.A(47446), TextCompare: false) != 0)
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
						if (Operators.CompareString(C, VH.A(47410), TextCompare: false) != 0)
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
						else
						{
							try
							{
								result = ((_Worksheet)B).get_Range((object)A.Substring(A.IndexOf(VH.A(39848)) + 1, A.IndexOf(VH.A(39904)) - A.IndexOf(VH.A(39848)) - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Column;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
					}
					else
					{
						try
						{
							result = ((_Worksheet)B).get_Range((object)A.Substring(A.IndexOf(VH.A(39848)) + 1, A.IndexOf(VH.A(39904)) - A.IndexOf(VH.A(39848)) - 1), RuntimeHelpers.GetObjectValue(Missing.Value)).Row;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
					}
				}
				else
				{
					A = E(A);
					result = RuntimeHelpers.GetObjectValue(this.m_A.Evaluate(A));
				}
			}
			else
			{
				result = RuntimeHelpers.GetObjectValue(this.m_A.Evaluate(A));
			}
			return result;
		}
	}

	private void A(ref object A, string B, string C, Worksheet D)
	{
		bool blnError = false;
		if (Operators.CompareString(C, VH.A(2015), TextCompare: false) != 0)
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
			if (Operators.CompareString(C, VH.A(2030), TextCompare: false) != 0)
			{
				if (Operators.CompareString(C, VH.A(49095), TextCompare: false) != 0)
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
					if (Operators.CompareString(C, VH.A(49102), TextCompare: false) != 0)
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
						if (Operators.CompareString(C, VH.A(4533), TextCompare: false) != 0)
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
							if (Operators.CompareString(C, VH.A(49109), TextCompare: false) != 0)
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
						}
					}
					else
					{
						Evaluate.SimplifyFunction(D, ref B, ref blnError, C, Evaluate.EvaluateMax);
					}
				}
				else
				{
					Evaluate.SimplifyFunction(D, ref B, ref blnError, C, Evaluate.EvaluateMin);
				}
				goto IL_011c;
			}
		}
		Evaluate.SimplifyFunction(D, ref B, ref blnError, C, Evaluate.EvaluateLookup);
		goto IL_011c;
		IL_011c:
		if (blnError)
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
			object objectValue = RuntimeHelpers.GetObjectValue(D.Evaluate(B));
			if (objectValue is Range)
			{
				if (!(A is object))
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
					if (!Operators.ConditionalCompareObjectEqual(((Range)objectValue).Value2, A, TextCompare: false))
					{
						goto IL_0177;
					}
				}
				A = RuntimeHelpers.GetObjectValue(objectValue);
			}
			goto IL_0177;
			IL_0177:
			objectValue = null;
			return;
		}
	}

	private List<string> A(string[] A, string B, int C, string D)
	{
		if (A.Length == C)
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
					return A.ToList();
				}
			}
		}
		int num = 0;
		int num2 = 1;
		int length = B.Length;
		List<string> list = new List<string>();
		int num3 = length;
		checked
		{
			for (int i = 1; i <= num3; i++)
			{
				string left = Strings.Mid(B, i, 1);
				if (Operators.CompareString(left, VH.A(39848), TextCompare: false) == 0)
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
					num++;
				}
				else if (Operators.CompareString(left, VH.A(39904), TextCompare: false) == 0)
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
					num--;
				}
				if (num != 0)
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
				if (Operators.CompareString(left, D, TextCompare: false) != 0)
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
				list.Add(Strings.Mid(B, num2, i - num2));
				num2 = i + 1;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				list.Add(Strings.Mid(B, num2, B.Length - num2 + 1));
				return list;
			}
		}
	}

	private GroupItem A(BaseItem A, ParenthesesPair B, Worksheet C, string D, bool E)
	{
		ParenthesesPair parenthesesPair = B;
		string text = D.Substring(parenthesesPair.StartIndex, parenthesesPair.Length);
		string a = string.Format(VH.A(48804), text);
		string strValue;
		if (!text.Contains(VH.A(49076)))
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
			strValue = this.A(RuntimeHelpers.GetObjectValue(C.Evaluate(this.A(a, E))));
		}
		else
		{
			strValue = this.A(RuntimeHelpers.GetObjectValue(this.m_A.Evaluate(this.A(this.E(a), E))));
		}
		GroupItem result = new GroupItem(A, text, strValue)
		{
			SelectionIndex = parenthesesPair.StartIndex,
			SelectionLength = parenthesesPair.Length
		};
		parenthesesPair = null;
		return result;
	}

	private void A(BaseItem A, MatchCollection B, Worksheet C, ref string D, ref List<BaseItem> E, ref bool F)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Match match = (Match)enumerator.Current;
				try
				{
					object objectValue = RuntimeHelpers.GetObjectValue(this.A(match.Value, C, A));
					if (!(objectValue is Range))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						Range range = (Range)objectValue;
						long num = Conversions.ToLong(range.Cells.CountLarge);
						if (num == 1)
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
							E.Add(this.A(A, range, match.Value, this.A(range), match.Index, match.Length));
							F = true;
						}
						else
						{
							bool g = !string.Format(VH.A(48282), match.Value).Contains(VH.A(2826));
							E.Add(this.A(A, range, match.Value, match.Index, match.Length, checked((int)num), g));
						}
						range = null;
						D = Helpers.MaskFormula(match, D);
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
				finally
				{
					object objectValue = null;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private object A(string A, Worksheet B, BaseItem C)
	{
		try
		{
			if (A.StartsWith(VH.A(48146)))
			{
				Range range = (Range)B.Evaluate(string.Format(VH.A(48804), A.Substring(1)));
				if (range.Areas.Count > 1)
				{
					return range;
				}
				bool C2 = false;
				Range range2 = this.A(C.Range, range, out C2);
				if (!C2)
				{
					object left;
					if (range2 == null)
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
						left = null;
					}
					else
					{
						left = range2.CountLarge;
					}
					if (Operators.ConditionalCompareObjectGreater(left, 1, TextCompare: false))
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								return range;
							}
						}
					}
					Range range3 = range2;
					if (range3 == null)
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
						range3 = range;
					}
					return range3;
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
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range range2 = null;
			Range range = null;
		}
		return B.Evaluate(string.Format(VH.A(48804), A));
	}

	private Range A(Range A, Range B, out bool C)
	{
		C = false;
		try
		{
			Worksheet worksheet = B.Worksheet;
			int B2 = default(int);
			int C2 = default(int);
			int D = default(int);
			int E = default(int);
			JH.A(B, ref B2, ref C2, ref D, ref E);
			int B3 = default(int);
			int C3 = default(int);
			int D2 = default(int);
			int E2 = default(int);
			JH.A(A, ref B3, ref C3, ref D2, ref E2);
			int num = Math.Max(B2, B3);
			int num2 = Math.Min(D, D2);
			int num3 = Math.Max(C2, C3);
			int num4 = Math.Min(E, E2);
			if (num2 < num)
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
				if (num4 < num3)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return null;
						}
					}
				}
			}
			if (num2 >= num)
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
				if (num4 >= num3)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return JH.A(worksheet, num, num3, num2, num4);
						}
					}
				}
			}
			if (num2 >= num)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return JH.A(worksheet, B3, C2, D2, E);
					}
				}
			}
			return JH.A(worksheet, B2, C3, D, E2);
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Worksheet worksheet = null;
		}
		C = true;
		return null;
	}

	private bool A(Range A)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		int num = default(int);
		int num3 = default(int);
		bool flag = default(bool);
		bool flag2 = default(bool);
		int num5 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		bool result = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					application = A.Application;
					goto IL_000c;
				case 462:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_000c;
						case 3:
							goto IL_0011;
						case 4:
							goto IL_0016;
						case 5:
							goto IL_001b;
						case 6:
							goto IL_0022;
						case 7:
							goto IL_0035;
						case 8:
							goto IL_004f;
						case 10:
							goto IL_0059;
						case 11:
							goto IL_006c;
						case 12:
							goto IL_007a;
						case 13:
							goto IL_0085;
						case 14:
							goto IL_008b;
						case 15:
							goto IL_0099;
						case 16:
							goto IL_00a4;
						case 17:
							goto IL_00b6;
						case 18:
							goto IL_00d0;
						case 19:
							goto IL_00da;
						case 20:
							goto IL_00eb;
						case 21:
							goto IL_00f6;
						case 22:
							goto IL_00fc;
						case 23:
							goto IL_0122;
						case 9:
						case 24:
							goto IL_0145;
						case 25:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 26:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00da:
					num2 = 19;
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
						goto IL_00eb;
					}
					goto IL_00f6;
					IL_000c:
					num2 = 2;
					flag2 = false;
					goto IL_0011;
					IL_0011:
					num2 = 3;
					flag = false;
					goto IL_0016;
					IL_0016:
					num2 = 4;
					num5 = 0;
					goto IL_001b;
					IL_001b:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0022;
					IL_0022:
					num2 = 6;
					num5 = A.DirectPrecedents.Count;
					goto IL_0035;
					IL_0035:
					num2 = 7;
					if (num5 > 0)
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
						goto IL_004f;
					}
					goto IL_0059;
					IL_0145:
					num2 = 24;
					application = null;
					break;
					IL_00eb:
					num2 = 20;
					application.ScreenUpdating = true;
					goto IL_00f6;
					IL_00f6:
					num2 = 21;
					workbook = null;
					goto IL_00fc;
					IL_004f:
					num2 = 8;
					flag2 = true;
					goto IL_0145;
					IL_0059:
					num2 = 10;
					if (A != null)
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
						goto IL_006c;
					}
					goto IL_0145;
					IL_00fc:
					num2 = 22;
					if (!flag2 && K.Settings.AuditOpenWorkbookLinks && Conversions.ToBoolean(A.HasFormula))
					{
						goto IL_0122;
					}
					goto IL_0145;
					IL_006c:
					num2 = 11;
					if (application.ScreenUpdating)
					{
						goto IL_007a;
					}
					goto IL_008b;
					IL_007a:
					num2 = 12;
					application.ScreenUpdating = false;
					goto IL_0085;
					IL_0085:
					num2 = 13;
					flag = true;
					goto IL_008b;
					IL_008b:
					num2 = 14;
					workbook = application.ActiveWorkbook;
					goto IL_0099;
					IL_0099:
					num2 = 15;
					flag2 = B(A);
					goto IL_00a4;
					IL_00a4:
					num2 = 16;
					A.ShowPrecedents(true);
					goto IL_00b6;
					IL_00b6:
					num2 = 17;
					if (application.ActiveWorkbook != workbook)
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
						goto IL_00d0;
					}
					goto IL_00da;
					IL_0122:
					num2 = 23;
					flag2 = Regex.IsMatch(A.Formula.ToString(), VH.A(49122));
					goto IL_0145;
					IL_00d0:
					num2 = 18;
					workbook.Activate();
					goto IL_00da;
					end_IL_0000_2:
					break;
				}
				num2 = 25;
				result = flag2;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 462;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool B(Range A)
	{
		bool result;
		try
		{
			A.ShowPrecedents(false);
			Range range = (Range)A.NavigateArrow(true, 1, 1);
			if (Operators.CompareString(range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
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
					result = false;
					break;
				}
			}
			else
			{
				if (Information.IsDBNull(RuntimeHelpers.GetObjectValue(range.MergeCells)))
				{
					goto IL_00f1;
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
				if (!Conversions.ToBoolean(range.MergeCells))
				{
					goto IL_00f1;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					result = false;
					break;
				}
			}
			goto end_IL_0000;
			IL_00f1:
			result = true;
			end_IL_0000:;
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			result = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Range range = null;
		}
		return result;
	}

	private List<Precedent> A(Range A)
	{
		bool flag = false;
		long num = 0L;
		List<Precedent> list = new List<Precedent>();
		Worksheet worksheet = A.Worksheet;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
		XlDisplayDrawingObjects displayDrawingObjects = workbook.DisplayDrawingObjects;
		workbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlHide;
		string name = worksheet.Name;
		string name2 = workbook.Name;
		bool.TryParse(Conversions.ToString(A.MergeCells), out var result);
		string right = ((!result) ? "" : A.MergeArea.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)));
		string right2 = A.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
		string input;
		try
		{
			input = NewLateBinding.LateGet(A, null, VH.A(1998), new object[0], null, null, null).ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			input = A.Formula.ToString();
			ProjectData.ClearProjectError();
		}
		try
		{
			A.ShowPrecedents(RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			if (worksheet.ProtectContents)
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
				System.Windows.Forms.MessageBox.Show(ex4.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			else
			{
				clsReporting.LogException(ex4);
			}
			flag = true;
			ProjectData.ClearProjectError();
		}
		checked
		{
			Range range2;
			MatchCollection matchCollection;
			if (!flag)
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
				bool flag2;
				IEnumerator enumerator = default(IEnumerator);
				do
				{
					flag2 = true;
					num++;
					long num2 = 0L;
					while (true)
					{
						num2++;
						Range range = null;
						try
						{
							range = (Range)A.NavigateArrow(true, num, num2);
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
							break;
						}
						string name3 = ((Microsoft.Office.Interop.Excel.Workbook)range.Worksheet.Parent).Name;
						string name4 = range.Worksheet.Name;
						Precedent precedent = new Precedent();
						if (Operators.CompareString(name2, name3, TextCompare: false) != 0)
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
							precedent.Address = VH.A(7120) + name3 + VH.A(43340) + name4 + VH.A(7827);
						}
						else
						{
							Precedent precedent2 = precedent;
							object address;
							if (Operators.CompareString(name, name4, TextCompare: false) == 0)
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
								address = null;
							}
							else
							{
								address = name4 + VH.A(7827);
							}
							precedent2.Address = (string)address;
						}
						precedent.Address += range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						precedent.Value = this.A(range);
						list.Add(precedent);
						precedent = null;
						string text = range.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
						if (Operators.ConditionalCompareObjectGreater(range.Cells.CountLarge, 1, TextCompare: false))
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
								matchCollection = null;
								matchCollection = Regex.Matches(text, VH.A(49197), RegexOptions.None);
								try
								{
									enumerator = matchCollection.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Match obj = (Match)enumerator.Current;
										string text2 = obj.Groups[1].ToString();
										string text3 = obj.Groups[2].ToString();
										if (!Regex.IsMatch(input, VH.A(49246) + text2 + VH.A(43088) + text3 + VH.A(49257)))
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
										precedent = new Precedent();
										if (Operators.CompareString(name2, name3, TextCompare: false) != 0)
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
											precedent.Address = VH.A(7120) + name3 + VH.A(43340) + name4 + VH.A(7827);
										}
										else
										{
											Precedent precedent3 = precedent;
											object address2;
											if (Operators.CompareString(name, name4, TextCompare: false) == 0)
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
												address2 = null;
											}
											else
											{
												address2 = name4 + VH.A(7827);
											}
											precedent3.Address = (string)address2;
										}
										range2 = Base.ResolveAddress(text2 + text3, A);
										precedent.Address += range2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										precedent.Value = this.A(range2);
										list.Add(precedent);
										precedent = null;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_053b;
										}
										continue;
										end_IL_053b:
										break;
									}
								}
								finally
								{
									if (enumerator is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
						}
						if (num > 1)
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
							if (Operators.CompareString(text, right2, TextCompare: false) == 0)
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
									A.NavigateArrow(true, num + 1, num2);
									if (Operators.CompareString(this.m_A.ActiveCell.get_Address((object)false, (object)false, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right2, TextCompare: false) == 0)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_019d;
											}
										}
									}
									flag2 = false;
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									ProjectData.ClearProjectError();
									break;
								}
								goto IL_0650;
							}
						}
						if (num > 1 && result)
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
							if (Operators.CompareString(text, right, TextCompare: false) == 0)
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
						flag2 = false;
						goto IL_0650;
						IL_0650:
						if (num != 1)
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
						if (num2 <= 100)
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
							break;
						}
						break;
						continue;
						end_IL_019d:
						break;
					}
				}
				while (!flag2);
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
			try
			{
				worksheet.ClearArrows();
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			workbook.DisplayDrawingObjects = displayDrawingObjects;
			workbook.Activate();
			worksheet.Activate();
			A.Select();
			workbook = null;
			worksheet = null;
			range2 = null;
			matchCollection = null;
			return list;
		}
	}

	private string A(object A)
	{
		string result;
		try
		{
			if (Operators.CompareString(A.ToString(), VH.A(49270), TextCompare: false) == 0)
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
					result = VH.A(41885);
					break;
				}
			}
			else
			{
				string thousandsSeparator = this.m_A.ThousandsSeparator;
				string decimalSeparator = this.m_A.DecimalSeparator;
				result = ((double)A).ToString(VH.A(49303) + thousandsSeparator + VH.A(49306) + decimalSeparator + VH.A(49311) + thousandsSeparator + VH.A(49306) + decimalSeparator + VH.A(49324));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if (A.ToString().Contains(VH.A(49337)))
			{
				result = VH.A(49364);
				ProjectData.ClearProjectError();
			}
			else
			{
				result = A.ToString();
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private string D(string A)
	{
		string result = A;
		try
		{
			Range range = (Range)this.A(this.m_A.ActiveWorkbook).Cells[1, 1];
			try
			{
				NewLateBinding.LateSetComplex(range, null, VH.A(8714), new object[1] { A }, null, null, OptimisticSet: false, RValueBase: true);
				result = Conversions.ToString(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				range.FormulaLocal = A;
				result = Conversions.ToString(range.Formula);
				ProjectData.ClearProjectError();
			}
			range = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string A(string A, bool B)
	{
		if (!B)
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
					return A;
				}
			}
		}
		return D(A);
	}

	private string E(string A)
	{
		string text = A;
		checked
		{
			while (text.Contains(VH.A(49076)))
			{
				string text2 = Strings.Mid(text, Strings.InStr(text, VH.A(4457)) + 9);
				text2 = Strings.Left(text2, Strings.InStr(text2, VH.A(39904)) - 1);
				string replacement = Conversions.ToString(this.m_A.Evaluate(text2));
				text = Strings.Replace(text, VH.A(49076) + text2 + VH.A(39904), replacement);
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
				return text;
			}
		}
	}

	private string A(Range A)
	{
		if (!Operators.ConditionalCompareObjectEqual(A.Cells.CountLarge, 1, TextCompare: false))
		{
			return VH.A(41885);
		}
		return B(A);
	}

	private string B(Range A)
	{
		string text = A.Text.ToString();
		if (!string.IsNullOrEmpty(text))
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
					return Base.CleanValueText(text);
				}
			}
		}
		return "";
	}

	private Worksheet A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		Worksheet worksheet = B(A);
		if (worksheet == null)
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
			try
			{
				worksheet = (Worksheet)A.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				worksheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
				worksheet.Name = wpfPrecedents.m_A;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
		}
		return worksheet;
	}

	private void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		Worksheet worksheet = B(A);
		if (worksheet == null)
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
			this.m_A.DisplayAlerts = false;
			try
			{
				worksheet.Visible = XlSheetVisibility.xlSheetHidden;
				worksheet.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			this.m_A.DisplayAlerts = true;
			worksheet = null;
			return;
		}
	}

	private Worksheet B(Microsoft.Office.Interop.Excel.Workbook A)
	{
		Worksheet result;
		try
		{
			result = (Worksheet)A.Worksheets.get_Item((object)wpfPrecedents.m_A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private void B(string A)
	{
		string text = "";
		try
		{
			this.m_A = this.m_A.ActiveWorkbook;
			Array array = (Array)this.m_A.ActiveWorkbook.LinkSources(XlLink.xlExcelLinks);
			int length = array.Length;
			IEnumerator enumerator = default(IEnumerator);
			for (int i = 1; i <= length; i = checked(i + 1))
			{
				string text2 = NewLateBinding.LateIndexGet(array, new object[1] { i }, null).ToString();
				try
				{
					enumerator = this.m_A.Workbooks.GetEnumerator();
					while (true)
					{
						if (!enumerator.MoveNext())
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_00c9;
								}
								continue;
								end_IL_00c9:
								break;
							}
							break;
						}
						Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)enumerator.Current;
						if (Operators.CompareString(workbook.FullName, text2, TextCompare: false) != 0)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							break;
						}
						goto end_IL_006e;
					}
					goto IL_00f5;
					end_IL_006e:;
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				continue;
				IL_00f5:
				if (text2.EndsWith(VH.A(49395)))
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
				if (text2.EndsWith(VH.A(49406)))
				{
					continue;
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
				string text3 = "";
				try
				{
					if (Regex.IsMatch(text2, VH.A(49415)))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							text3 = Path.GetFileName(text2);
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
				if (!A.Contains(text3))
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
				if (text3.Length <= 0)
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
				try
				{
					this.m_A.ScreenUpdating = true;
					Microsoft.Office.Interop.Excel.Workbook workbook = this.m_A.Workbooks.Open(text2, false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					if (workbook == null)
					{
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (this.m_A == null)
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
							this.m_A = new List<Microsoft.Office.Interop.Excel.Workbook>();
						}
						this.m_A.Add(workbook);
						workbook = null;
						new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(47767)).AddEventHandler(this.m_A, new AppEvents_WorkbookActivateEventHandler(this.B));
						break;
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					text = text + VH.A(7803) + text2;
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		if (text.Length <= 0)
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
			E(VH.A(49452) + text);
			return;
		}
	}

	private void B(Microsoft.Office.Interop.Excel.Workbook A)
	{
		try
		{
			if (this.m_A.Contains(A))
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
				this.m_A.Remove(A);
				try
				{
					new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(47767)).RemoveEventHandler(MH.A.Application, new AppEvents_WorkbookActivateEventHandler(this.B));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (this.m_A.Any())
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
				this.m_A.Activate();
				Host.Activate();
				this.m_A = null;
				this.m_A = null;
				return;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void I()
	{
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		TC tC = TC.A;
		BaseItem baseItem = (BaseItem)trvTrace.SelectedItem;
		if (!(baseItem is FunctionItem))
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
			if (!(baseItem is GroupItem))
			{
				if (!(baseItem is MultiCellItem))
				{
					if (!(baseItem is ThreeDRangeItem))
					{
						goto IL_0071;
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
				flag2 = true;
				goto IL_0071;
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
		tC = TC.C;
		flag3 = true;
		goto IL_0071;
		IL_0071:
		BaseItem baseItem2 = (BaseItem)baseItem.Parent;
		BaseItem baseItem3;
		if (!flag2)
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
			if (!flag3)
			{
				if (baseItem.Items.Count == 0)
				{
					try
					{
						if (baseItem2.Range != null && Operators.ConditionalCompareObjectGreater(baseItem2.Range.Cells.CountLarge, 1, TextCompare: false))
						{
							tC = TC.C;
							baseItem3 = baseItem;
						}
						else
						{
							baseItem3 = A(baseItem2);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						baseItem3 = A(baseItem2);
						ProjectData.ClearProjectError();
					}
				}
				else if (baseItem.IsExpanded)
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
					baseItem3 = baseItem;
				}
				else
				{
					baseItem3 = A(baseItem2);
				}
				goto IL_0127;
			}
		}
		baseItem3 = A(baseItem2);
		goto IL_0127;
		IL_0127:
		baseItem = null;
		baseItem2 = null;
		Range range = baseItem3.Range;
		string text;
		try
		{
			if (Conversions.ToBoolean(range.HasArray))
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					D(VH.A(49553));
					baseItem3 = null;
					range = null;
					return;
				}
			}
			try
			{
				text = NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null).ToString();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				text = range.Formula.ToString();
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			E(VH.A(37357));
			baseItem3 = null;
			range = null;
			ProjectData.ClearProjectError();
			return;
		}
		bool flag4 = Operators.CompareString(text, tbFormula.Text, TextCompare: false) == 0;
		Range visibleRange = this.m_A.ActiveWindow.ActivePane.VisibleRange;
		bool moveAfterReturn = this.m_A.MoveAfterReturn;
		this.m_A.MoveAfterReturn = false;
		this.m_A.ScreenUpdating = false;
		this.m_A.EnableEvents = false;
		Range range2 = ((BaseItem)trvTrace.SelectedItem).Range;
		bool flag5 = range.Worksheet == range2.Worksheet;
		Range range3 = null;
		if (!flag5)
		{
			this.m_A.Goto(range2, false);
		}
		try
		{
			Interaction.AppActivate(this.m_A.Caption);
			A(range);
			this.m_A.SendKeys(VH.A(49606), true);
			if (flag4)
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
				if (HighlightedInline != null)
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
					BaseItem baseItem4;
					if (HighlightedInline is Run)
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
						baseItem4 = (BaseItem)((Run)HighlightedInline).Tag;
					}
					else
					{
						baseItem4 = (BaseItem)((Span)HighlightedInline).Tag;
					}
					BaseItem baseItem5 = baseItem4;
					int selectionIndex = baseItem5.SelectionIndex;
					int selectionLength = baseItem5.SelectionLength;
					_ = null;
					baseItem4 = null;
					Microsoft.Office.Interop.Excel.Application a = this.m_A;
					if ((double)selectionIndex < (double)text.Length / 2.0)
					{
						a.SendKeys(VH.A(49615) + selectionIndex + VH.A(19802), true);
					}
					else
					{
						a.SendKeys(VH.A(49642) + checked(text.Length - selectionIndex) + VH.A(19802), true);
					}
					a.SendKeys(VH.A(49655) + selectionLength + VH.A(19802), true);
					if (tC != TC.A)
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
						if (tC != TC.B)
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
						else
						{
							a.SendKeys(VH.A(49606), true);
						}
					}
					else
					{
						a.SendKeys(VH.A(49606), true);
						try
						{
							if (flag5)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									System.Windows.Forms.Application.DoEvents();
									a.Goto(range2, false);
									a.SendKeys(VH.A(49672), true);
									break;
								}
							}
							else
							{
								a.SendKeys(VH.A(49703), true);
								a.SendKeys(VH.A(49712), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
						}
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							E(ex8.Message);
							ProjectData.ClearProjectError();
						}
					}
					a = null;
				}
			}
			do
			{
				System.Windows.Forms.Application.DoEvents();
				Thread.Sleep(10);
			}
			while (Environment.IsEditMode(this.m_A));
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_051e;
				}
				continue;
				end_IL_051e:
				break;
			}
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			E(ex10.Message);
			clsReporting.LogException(ex10);
			flag = true;
			ProjectData.ClearProjectError();
		}
		if (range3 != null)
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
				range3.Worksheet.Activate();
				Base.ScrollTo(range3);
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				ProjectData.ClearProjectError();
			}
			range3 = null;
		}
		this.m_A.ScreenUpdating = true;
		this.m_A.EnableEvents = true;
		this.m_A.MoveAfterReturn = moveAfterReturn;
		Host.Activate();
		bool flag6;
		try
		{
			flag6 = Operators.CompareString(NewLateBinding.LateGet(range, null, VH.A(1998), new object[0], null, null, null).ToString(), text, TextCompare: false) != 0;
		}
		catch (Exception ex13)
		{
			ProjectData.SetProjectError(ex13);
			Exception ex14 = ex13;
			flag6 = Operators.CompareString(range.Formula.ToString(), text, TextCompare: false) != 0;
			ProjectData.ClearProjectError();
		}
		if (flag6 && !flag)
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
			try
			{
				this.m_B = true;
				baseItem3.IsExpanded = false;
				this.m_B = false;
				baseItem3.IsExpanded = true;
				baseItem3.IsSelected = true;
				tbFormula.Text = tbFormula.Text;
				A(baseItem3, tbFormula.Text);
			}
			catch (Exception ex15)
			{
				ProjectData.SetProjectError(ex15);
				Exception ex16 = ex15;
				E(ex16.Message);
				clsReporting.LogException(ex16);
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			this.m_A.ScreenUpdating = false;
			this.m_A.EnableEvents = false;
			try
			{
				range2 = ((BaseItem)trvTrace.SelectedItem).Range;
				Worksheet worksheet = range2.Worksheet;
				((Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent).Activate();
				worksheet.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
				_ = null;
				this.m_A.Goto(range2, false);
				Base.ScrollTo(visibleRange);
			}
			catch (Exception ex17)
			{
				ProjectData.SetProjectError(ex17);
				Exception ex18 = ex17;
				clsReporting.LogException(ex18);
				ProjectData.ClearProjectError();
			}
			this.m_A.EnableEvents = true;
			this.m_A.ScreenUpdating = true;
		}
		baseItem3 = null;
		range = null;
		visibleRange = null;
		range2 = null;
	}

	private void C(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_F)
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
			this.m_F = true;
			Uri resourceLocator = new Uri(VH.A(49749), UriKind.Relative);
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
					ThisWindow = (wpfPrecedents)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					grdMain = (Grid)target;
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
					grdSplit = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			scroller = (ScrollViewer)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					tbFormula = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			tbDummy = (TextBlock)target;
			return;
		}
		if (connectionId == 8)
		{
			chkWrap = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					masterC = (ColumnDefinition)target;
					return;
				}
			}
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
					trvTrace = (System.Windows.Controls.TreeView)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			chkSettings = (System.Windows.Controls.CheckBox)target;
			return;
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
					btnOk = (System.Windows.Controls.Button)target;
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
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			popSettings = (Popup)target;
			return;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					tbEvaluate = (TextBlock)target;
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
					chkEvaluate = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					tbArguments = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkArguments = (System.Windows.Controls.CheckBox)target;
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
					chkArrows = (System.Windows.Controls.CheckBox)target;
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
					chkUnhide = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			chkOpenLinks = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					tbHighlight = (TextBlock)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 23:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				chkHighlight = (System.Windows.Controls.CheckBox)target;
				return;
			}
		case 24:
			chkMove = (System.Windows.Controls.CheckBox)target;
			break;
		default:
			this.m_F = true;
			break;
		}
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
		if (connectionId != 2)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnSelected);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.ExpandedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnExpanded);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = TreeViewItem.CollapsedEvent;
			eventSetter.Handler = new RoutedEventHandler(OnCollapsed);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
