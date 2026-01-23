using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using ExcelAddIn1.Workbook;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Worksheet
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData, string> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData A)
		{
			return A.DataType;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor A)
		{
			return A.OleColor;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor A)
		{
			return A.OleColor;
		}
	}

	[CompilerGenerated]
	internal sealed class W
	{
		public Analysis A;

		public Settings A;

		public bool A;

		public List<int> A;

		public Func<Range, string, Observation> A;

		public Func<Range, Observation> A;

		public Func<Range, Observation> B;

		public Func<Range, Observation> C;

		public Func<Range, Observation> D;

		public Func<Range, Observation> E;

		public Func<Range, int, Observation> A;

		public Func<Range, int, Observation> B;

		[SpecialName]
		internal void A(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.A(this.A.Observations, this.A.HiddenSheets, A);
		}

		[SpecialName]
		internal bool A()
		{
			return this.A.HiddenSheets != Severity.Ignore;
		}

		[SpecialName]
		internal void B(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.B(this.A.Observations, this.A.VeryHiddenSheets, A);
		}

		[SpecialName]
		internal bool B()
		{
			return this.A.VeryHiddenSheets != Severity.Ignore;
		}

		[SpecialName]
		internal void C(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.C(this.A.Observations, this.A.HiddenRowsColumns, A);
		}

		[SpecialName]
		internal bool C()
		{
			return this.A.HiddenRowsColumns != Severity.Ignore;
		}

		[SpecialName]
		internal void D(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.D(this.A.Observations, this.A.CollapsedRowsColumns, A);
		}

		[SpecialName]
		internal bool D()
		{
			return this.A.CollapsedRowsColumns != Severity.Ignore;
		}

		[SpecialName]
		internal void E(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.E(this.A.Observations, this.A.UsedRangeInflation, A);
		}

		[SpecialName]
		internal bool E()
		{
			return this.A.UsedRangeInflation != Severity.Ignore;
		}

		[SpecialName]
		internal void F(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			CircularReferences.A(this.A.Observations, this.A.CircularReferences, A);
		}

		[SpecialName]
		internal bool F()
		{
			if (!this.A)
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
						return this.A.CircularReferences != Severity.Ignore;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal void G(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			FormulaInterruption.A(this.A, this.A.FormulaInterruption, RangeHelpers.A(A));
		}

		[SpecialName]
		internal bool G()
		{
			return this.A.FormulaInterruption != Severity.Ignore;
		}

		[SpecialName]
		internal void H(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Formulas.A(this.A.Observations, this.A.FormulaErrors, RangeHelpers.A(A));
		}

		[SpecialName]
		internal bool H()
		{
			return this.A.FormulaErrors != Severity.Ignore;
		}

		[SpecialName]
		internal void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			PrivacySecurity.A(this.A, ref A, this.A.SensitiveData, B.UsedRange);
		}

		[SpecialName]
		internal void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData> sensitiveData = A.SensitiveData;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData, string> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Func<Range, string, Observation> d;
			if (this.A != null)
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
				d = this.A;
			}
			else
			{
				d = (this.A = [SpecialName] (Range rng, string B) => new ExcelAddIn1.Audit.Check.Observations.SensitiveData(this.A.SensitiveData, rng, B));
			}
			Worksheet.A(a, sensitiveData, c, d);
		}

		[SpecialName]
		internal Observation A(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.SensitiveData(this.A.SensitiveData, A, B);
		}

		[SpecialName]
		internal bool I()
		{
			return this.A.SensitiveData != Severity.Ignore;
		}

		[SpecialName]
		internal void I(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			PrivacySecurity.A(this.A, this.A.CommentsAndNotes, A);
		}

		[SpecialName]
		internal bool J()
		{
			return this.A.CommentsAndNotes != Severity.Ignore;
		}

		[SpecialName]
		internal void J(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.A(this.A, this.A.ShapesOverNonEmptyCells, A);
		}

		[SpecialName]
		internal bool K()
		{
			return this.A.ShapesOverNonEmptyCells != Severity.Ignore;
		}

		[SpecialName]
		internal void K(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Formatting.A(this.A, this.A.ExcessConditionalFormatting, A);
		}

		[SpecialName]
		internal bool L()
		{
			return this.A.ExcessConditionalFormatting != Severity.Ignore;
		}

		[SpecialName]
		internal void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Data.A(ref A, this.A.DataValidationIgnored, B.UsedRange);
		}

		[SpecialName]
		internal void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<Range> dataValidationFailed = A.DataValidationFailed;
			Func<Range, Observation> c;
			if (this.A != null)
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
				c = this.A;
			}
			else
			{
				c = (this.A = [SpecialName] (Range rng) => new ValidationFailed(this.A.DataValidationIgnored, rng));
			}
			Worksheet.A(a, dataValidationFailed, c);
		}

		[SpecialName]
		internal Observation A(Range A)
		{
			return new ValidationFailed(this.A.DataValidationIgnored, A);
		}

		[SpecialName]
		internal bool M()
		{
			return this.A.DataValidationIgnored != Severity.Ignore;
		}

		[SpecialName]
		internal void L(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Worksheet.A(this.A, this.A.UnusedNumericInputs, A.UsedRange);
		}

		[SpecialName]
		internal bool N()
		{
			return this.A.UnusedNumericInputs != Severity.Ignore;
		}

		[SpecialName]
		internal void M(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Formatting.B(this.A, this.A.InputsNotColored, A.UsedRange);
		}

		[SpecialName]
		internal bool O()
		{
			return this.A.InputsNotColored != Severity.Ignore;
		}

		[SpecialName]
		internal void N(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Data.A(this.A, this.A.DataOutliers, A.UsedRange);
		}

		[SpecialName]
		internal bool P()
		{
			return this.A.DataOutliers != Severity.Ignore;
		}

		[SpecialName]
		internal void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Data.A(this.A, ref A, this.A.NumbersStoredAsText, B.UsedRange);
		}

		[SpecialName]
		internal void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Worksheet.A(this.A, A.NumbersStoredAsText, [SpecialName] (Range rng) => new NumberStoredAsText(this.A.NumbersStoredAsText, rng));
		}

		[SpecialName]
		internal Observation B(Range A)
		{
			return new NumberStoredAsText(this.A.NumbersStoredAsText, A);
		}

		[SpecialName]
		internal bool Q()
		{
			return this.A.NumbersStoredAsText != Severity.Ignore;
		}

		[SpecialName]
		internal void O(Microsoft.Office.Interop.Excel.Worksheet A)
		{
			Formatting.A(this.A, this.A.MergedCells, A.UsedRange);
		}

		[SpecialName]
		internal bool R()
		{
			return this.A.MergedCells != Severity.Ignore;
		}

		[SpecialName]
		internal void D(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Worksheet.A(ref A, this.A.EmptyCellCommentsNotes, B);
		}

		[SpecialName]
		internal void D(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<Range> emptyCellComments = A.EmptyCellComments;
			Func<Range, Observation> c;
			if (this.C != null)
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
				c = this.C;
			}
			else
			{
				c = (this.C = [SpecialName] (Range rng) => new EmptyCellComment(this.A.EmptyCellCommentsNotes, rng));
			}
			Worksheet.A(a, emptyCellComments, c);
			Worksheet.A(this.A, A.EmptyCellNotes, [SpecialName] (Range rng) => new EmptyCellNote(this.A.EmptyCellCommentsNotes, rng));
		}

		[SpecialName]
		internal Observation C(Range A)
		{
			return new EmptyCellComment(this.A.EmptyCellCommentsNotes, A);
		}

		[SpecialName]
		internal Observation D(Range A)
		{
			return new EmptyCellNote(this.A.EmptyCellCommentsNotes, A);
		}

		[SpecialName]
		internal bool S()
		{
			return this.A.EmptyCellCommentsNotes != Severity.Ignore;
		}

		[SpecialName]
		internal void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formatting.A(ref A, this.A.TripleSemicolonNumFormat, B);
		}

		[SpecialName]
		internal void E(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<Range> tripleSemicolons = A.TripleSemicolons;
			Func<Range, Observation> c;
			if (this.E != null)
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
				c = this.E;
			}
			else
			{
				c = (this.E = [SpecialName] (Range rng) => new TripleSemicolon(this.A.TripleSemicolonNumFormat, rng));
			}
			Worksheet.A(a, tripleSemicolons, c);
		}

		[SpecialName]
		internal Observation E(Range A)
		{
			return new TripleSemicolon(this.A.TripleSemicolonNumFormat, A);
		}

		[SpecialName]
		internal bool T()
		{
			return this.A.TripleSemicolonNumFormat != Severity.Ignore;
		}

		[SpecialName]
		internal void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formatting.A(ref A, this.A.CellFillColor, B, this.A);
		}

		[SpecialName]
		internal void F(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor> cellFillColors = A.CellFillColors;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				c = _Closure_0024__.A;
			}
			Func<Range, int, Observation> d;
			if (this.A != null)
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
				d = this.A;
			}
			else
			{
				d = (this.A = [SpecialName] (Range rng, int B) => new ExcelAddIn1.Audit.Check.Observations.CellFillColor(this.A.CellFillColor, rng, B));
			}
			Worksheet.A(a, cellFillColors, c, d);
		}

		[SpecialName]
		internal Observation A(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.CellFillColor(this.A.CellFillColor, A, B);
		}

		[SpecialName]
		internal bool U()
		{
			return this.A.CellFillColor != Severity.Ignore;
		}

		[SpecialName]
		internal void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formatting.B(ref A, this.A.CellBorderColor, B, this.A);
		}

		[SpecialName]
		internal void G(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor> cellBorderColors = A.CellBorderColors;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Func<Range, int, Observation> d;
			if (this.B != null)
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
				d = this.B;
			}
			else
			{
				d = (this.B = [SpecialName] (Range rng, int B) => new ExcelAddIn1.Audit.Check.Observations.CellBorderColor(this.A.CellBorderColor, rng, B));
			}
			Worksheet.A(a, cellBorderColors, c, d);
		}

		[SpecialName]
		internal Observation B(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.CellBorderColor(this.A.CellBorderColor, A, B);
		}

		[SpecialName]
		internal bool V()
		{
			return this.A.CellBorderColor != Severity.Ignore;
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__3<_0024CLS0, _0024CLS1> where _0024CLS0 : RawObservation
	{
		public static readonly _Closure_0024__3<_0024CLS0, _0024CLS1> A;

		public static Func<_0024CLS0, Range> A;

		static _Closure_0024__3()
		{
			_Closure_0024__3<_0024CLS0, _0024CLS1>.A = new _Closure_0024__3<_0024CLS0, _0024CLS1>();
		}

		[SpecialName]
		internal Range A(_0024CLS0 A)
		{
			return A.Range;
		}
	}

	[CompilerGenerated]
	internal sealed class X<A, B> where A : RawObservation
	{
		public Func<Range, B, Observation> A;

		public X(X<A, B> A)
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
	}

	[CompilerGenerated]
	internal sealed class Y<C, D> where C : RawObservation
	{
		public IGrouping<D, C> A;

		public X<C, D> A;

		public Y(Y<C, D> A)
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
		internal Observation A(Range A)
		{
			return this.A.A(A, this.A.Key);
		}
	}

	internal static List<RB> A(Analysis A, Settings B, bool C, List<int> D)
	{
		List<RB> list = new List<RB>
		{
			new XB(VH.A(2097), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.A(A.Observations, B.HiddenSheets, c);
			}, [SpecialName] () => B.HiddenSheets != Severity.Ignore),
			new XB(VH.A(2124), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.B(A.Observations, B.VeryHiddenSheets, c);
			}, [SpecialName] () => B.VeryHiddenSheets != Severity.Ignore),
			new XB(VH.A(6153), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.C(A.Observations, B.HiddenRowsColumns, c);
			}, [SpecialName] () => B.HiddenRowsColumns != Severity.Ignore),
			new XB(VH.A(6192), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.D(A.Observations, B.CollapsedRowsColumns, c);
			}, [SpecialName] () => B.CollapsedRowsColumns != Severity.Ignore),
			new XB(VH.A(6237), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.E(A.Observations, B.UsedRangeInflation, c);
			}, [SpecialName] () => B.UsedRangeInflation != Severity.Ignore),
			new XB(VH.A(6105), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				CircularReferences.A(A.Observations, B.CircularReferences, c);
			}, [SpecialName] () =>
			{
				if (!C)
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
							return B.CircularReferences != Severity.Ignore;
						}
					}
				}
				return false;
			}),
			new XB(VH.A(6278), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet a) =>
			{
				FormulaInterruption.A(A, B.FormulaInterruption, RangeHelpers.A(a));
			}, [SpecialName] () => B.FormulaInterruption != Severity.Ignore),
			new XB(VH.A(6321), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet a) =>
			{
				Formulas.A(A.Observations, B.FormulaErrors, RangeHelpers.A(a));
			}, [SpecialName] () => B.FormulaErrors != Severity.Ignore)
		};
		list.AddRange(Formulas.A(A, B));
		Func<Range, Observation> C2 = default(Func<Range, Observation>);
		Func<Range, int, Observation> A2 = default(Func<Range, int, Observation>);
		list.AddRange(new RB[11]
		{
			new XB(VH.A(6350), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations B2, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				PrivacySecurity.A(A, ref B2, B.SensitiveData, worksheet.UsedRange);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData> sensitiveData = observations.SensitiveData;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.SensitiveData, string> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Func<Range, string, Observation> d;
				if (A2 != null)
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
					d = A2;
				}
				else
				{
					d = (A2 = [SpecialName] (Range rng, string strDataType) => new ExcelAddIn1.Audit.Check.Observations.SensitiveData(B.SensitiveData, rng, strDataType));
				}
				Worksheet.A(a, sensitiveData, c, d);
			}, [SpecialName] () => B.SensitiveData != Severity.Ignore),
			new XB(VH.A(6379), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				PrivacySecurity.A(A, B.CommentsAndNotes, c);
			}, [SpecialName] () => B.CommentsAndNotes != Severity.Ignore),
			new XB(VH.A(6416), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.A(A, B.ShapesOverNonEmptyCells, c);
			}, [SpecialName] () => B.ShapesOverNonEmptyCells != Severity.Ignore),
			new XB(VH.A(6451), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Formatting.A(A, B.ExcessConditionalFormatting, c);
			}, [SpecialName] () => B.ExcessConditionalFormatting != Severity.Ignore),
			new XB(VH.A(6496), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations A2, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Data.A(ref A2, B.DataValidationIgnored, worksheet.UsedRange);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<Range> dataValidationFailed = observations.DataValidationFailed;
				Func<Range, Observation> c;
				if (A2 != null)
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
					c = A2;
				}
				else
				{
					c = (A2 = [SpecialName] (Range rng) => new ValidationFailed(B.DataValidationIgnored, rng));
				}
				Worksheet.A(a, dataValidationFailed, c);
			}, [SpecialName] () => B.DataValidationIgnored != Severity.Ignore),
			new XB(VH.A(6527), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Worksheet.A(A, B.UnusedNumericInputs, worksheet.UsedRange);
			}, [SpecialName] () => B.UnusedNumericInputs != Severity.Ignore),
			new XB(VH.A(6570), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Formatting.B(A, B.InputsNotColored, worksheet.UsedRange);
			}, [SpecialName] () => B.InputsNotColored != Severity.Ignore),
			new XB(VH.A(6605), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Data.A(A, B.DataOutliers, worksheet.UsedRange);
			}, [SpecialName] () => B.DataOutliers != Severity.Ignore),
			new XB(VH.A(6632), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations B2, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Data.A(A, ref B2, B.NumbersStoredAsText, worksheet.UsedRange);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Worksheet.A(A, observations.NumbersStoredAsText, [SpecialName] (Range rng) => new NumberStoredAsText(B.NumbersStoredAsText, rng));
			}, [SpecialName] () => B.NumbersStoredAsText != Severity.Ignore),
			new XB(VH.A(6663), [SpecialName] (Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Formatting.A(A, B.MergedCells, worksheet.UsedRange);
			}, [SpecialName] () => B.MergedCells != Severity.Ignore),
			new XB(VH.A(6686), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations A2, Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Worksheet.A(ref A2, B.EmptyCellCommentsNotes, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<Range> emptyCellComments = observations.EmptyCellComments;
				Func<Range, Observation> c;
				if (C2 != null)
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
					c = C2;
				}
				else
				{
					c = (C2 = [SpecialName] (Range rng) => new EmptyCellComment(B.EmptyCellCommentsNotes, rng));
				}
				Worksheet.A(a, emptyCellComments, c);
				Worksheet.A(A, observations.EmptyCellNotes, [SpecialName] (Range rng) => new EmptyCellNote(B.EmptyCellCommentsNotes, rng));
			}, [SpecialName] () => B.EmptyCellCommentsNotes != Severity.Ignore)
		});
		Func<Range, Observation> E = default(Func<Range, Observation>);
		Func<Range, int, Observation> B2 = default(Func<Range, int, Observation>);
		list.AddRange(new RB[3]
		{
			new BC(A, VH.A(6737), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations A2, Range c, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Formatting.A(ref A2, B.TripleSemicolonNumFormat, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<Range> tripleSemicolons = observations.TripleSemicolons;
				Func<Range, Observation> c;
				if (E != null)
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
					c = E;
				}
				else
				{
					c = (E = [SpecialName] (Range rng) => new TripleSemicolon(B.TripleSemicolonNumFormat, rng));
				}
				Worksheet.A(a, tripleSemicolons, c);
			}, [SpecialName] () => B.TripleSemicolonNumFormat != Severity.Ignore),
			new BC(A, VH.A(6770), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations A2, Range c, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Formatting.A(ref A2, B.CellFillColor, c, D);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor> cellFillColors = observations.CellFillColors;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellFillColor, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					c = _Closure_0024__.A;
				}
				Func<Range, int, Observation> d;
				if (A2 != null)
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
					d = A2;
				}
				else
				{
					d = (A2 = [SpecialName] (Range rng, int intOleColor) => new ExcelAddIn1.Audit.Check.Observations.CellFillColor(B.CellFillColor, rng, intOleColor));
				}
				Worksheet.A(a, cellFillColors, c, d);
			}, [SpecialName] () => B.CellFillColor != Severity.Ignore),
			new BC(A, VH.A(6801), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations A2, Range c, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Formatting.B(ref A2, B.CellBorderColor, c, D);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor> cellBorderColors = observations.CellBorderColors;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.CellBorderColor, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Func<Range, int, Observation> d;
				if (B2 != null)
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
					d = B2;
				}
				else
				{
					d = (B2 = [SpecialName] (Range rng, int intOleColor) => new ExcelAddIn1.Audit.Check.Observations.CellBorderColor(B.CellBorderColor, rng, intOleColor));
				}
				Worksheet.A(a, cellBorderColors, c, d);
			}, [SpecialName] () => B.CellBorderColor != Severity.Ignore)
		});
		return list;
	}

	internal static void A(Analysis A, List<Range> B, Func<Range, Observation> C)
	{
		Range range = Worksheet.A(A, B);
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = range.Areas.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range arg = (Range)enumerator.Current;
				A.Observations.Add(C(arg));
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		range = null;
	}

	internal static void A<A, B>(Analysis A, List<A> B, Func<A, B> C, Func<Range, B, Observation> D) where A : RawObservation
	{
		X<A, B> x = new X<A, B>(x);
		x.A = D;
		IEnumerator<IGrouping<B, A>> enumerator = default(IEnumerator<IGrouping<B, A>>);
		try
		{
			enumerator = B.GroupBy(C).GetEnumerator();
			Y<A, B> y = default(Y<A, B>);
			while (enumerator.MoveNext())
			{
				y = new Y<A, B>(y);
				y.A = x;
				y.A = enumerator.Current;
				IGrouping<B, A> source = y.A;
				Func<A, Range> selector;
				if (_Closure_0024__3<A, B>.A != null)
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
					selector = _Closure_0024__3<A, B>.A;
				}
				else
				{
					selector = (_Closure_0024__3<A, B>.A = [SpecialName] (A val) => val.Range);
				}
				Worksheet.A(A, source.Select(selector).ToList(), y.A);
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
	}

	private static Range A(Analysis A, List<Range> B, string C = null)
	{
		if (B != null)
		{
			if (B.Count != 0)
			{
				Range A2 = null;
				int a = A.A(modFunctionsStr.BlankTo(C, VH.A(6836)), B.Count);
				try
				{
					foreach (Range item in B)
					{
						if (A.ItemCancelled())
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0070;
								}
								continue;
								end_IL_0070:
								break;
							}
							break;
						}
						RangeHelpers.A(ref A2, item);
					}
				}
				finally
				{
					A.A(a);
				}
				return A2;
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
		return null;
	}

	private static void A(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
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
			if (C.Visible != XlSheetVisibility.xlSheetHidden)
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
				A.Add(new HiddenSheet(B, C));
				return;
			}
		}
	}

	private static void B(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
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
			if (C.Visible != XlSheetVisibility.xlSheetVeryHidden)
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
				A.Add(new VeryHiddenSheet(B, C));
				return;
			}
		}
	}

	private static void C(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
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
			if (!Ranges.HasHiddenCells(C.Cells))
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
				A.Add(new HiddenRowColumn(B, C, C.Cells));
				return;
			}
		}
	}

	private static void D(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			if (!Ranges.HasHiddenCells(C.Cells))
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
				Range range = RangeHelpers.D(C);
				if (range == null)
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
					try
					{
						enumerator = range.Rows.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Range range2 = (Range)enumerator.Current;
							if (Conversions.ToLong(range2.OutlineLevel) <= 1)
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
							A.Add(new CollapsedRowColumn(B, C, range2));
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					enumerator2 = range.Columns.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							Range range3 = (Range)enumerator2.Current;
							if (Conversions.ToLong(range3.OutlineLevel) > 1)
							{
								A.Add(new CollapsedRowColumn(B, C, range3));
							}
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_010e;
							}
							continue;
							end_IL_010e:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					range = null;
					return;
				}
			}
		}
	}

	private static void E(List<Observation> A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		Range usedRange = C.UsedRange;
		Microsoft.Office.Interop.Excel.Worksheet worksheet = C;
		if (Operators.CompareString(usedRange.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), ((Range)worksheet.Cells[1, 1]).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) != 0)
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
			(long, int) lastRowColumn;
			try
			{
				lastRowColumn = Optimize.GetLastRowColumn(C);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
			if (Conversions.ToInteger(((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[1, lastRowColumn.Item2])).Columns.CountLarge) < Conversions.ToInteger(usedRange.Columns.CountLarge))
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
				A.Add(new UsedRangeInflated(B, usedRange));
			}
			else if (Conversions.ToInteger(((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(worksheet.Cells[1, 1]), RuntimeHelpers.GetObjectValue(worksheet.Cells[lastRowColumn.Item1, 1])).Rows.CountLarge) < Conversions.ToInteger(usedRange.Rows.CountLarge))
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
				A.Add(new UsedRangeInflated(B, usedRange));
			}
		}
		worksheet = null;
		usedRange = null;
	}

	private static void A(Analysis A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = RangeHelpers.H(C.UsedRange);
			if (range == null)
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
				Application application = C.Application;
				A.ActionStarted(VH.A(6879), C.Shapes.Count);
				Range range2;
				try
				{
					enumerator = C.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							Shape shape = (Shape)enumerator.Current;
							if (A.ItemCancelled())
							{
								break;
							}
							if (shape.Visible != MsoTriState.msoTrue)
							{
								continue;
							}
							range2 = ((_Application)application).get_Range((object)shape.TopLeftCell, (object)shape.BottomRightCell);
							if (application.Intersect(range2, range, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
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
							A.Observations.Add(new ShapeOverCells(B, range2, shape));
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_023b;
							}
							continue;
							end_IL_023b:
							break;
						}
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
				A.ActionEnded();
				range2 = null;
				range = null;
				application = null;
				return;
			}
		}
	}

	private static void A(Analysis A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = RangeHelpers.A(C, A, VH.A(2512), 500L);
			if (range == null)
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
				try
				{
					enumerator = range.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range rng = (Range)enumerator.Current;
						A.Observations.Add(new UnusedNumericInput(B, rng));
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_007d;
						}
						continue;
						end_IL_007d:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				range = null;
				return;
			}
		}
	}

	private static void A(ref ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		if (B == Severity.Ignore)
		{
			return;
		}
		IEnumerator enumerator2 = default(IEnumerator);
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
			using (List<object>.Enumerator enumerator = QB.A(C).GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					Range range = QB.A(RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(enumerator.Current)));
					if (range == null)
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
					if (!RangeHelpers.B(range))
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
					A.EmptyCellComments.Add(range);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_0086;
					}
					continue;
					end_IL_0086:
					break;
				}
			}
			Range range2 = RangeHelpers.F(C.Cells);
			if (range2 == null)
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
				try
				{
					enumerator2 = range2.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Range range3 = (Range)enumerator2.Current;
						if (RangeHelpers.B(range3))
						{
							A.EmptyCellNotes.Add(range3);
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_00fb;
						}
						continue;
						end_IL_00fb:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				range2 = null;
				return;
			}
		}
	}
}
