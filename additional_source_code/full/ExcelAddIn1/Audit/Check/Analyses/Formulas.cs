using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using ExcelAddIn1.Audit.Check.Observations;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using ExcelAddIn1.Formulas;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Analyses;

public sealed class Formulas
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity, int> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.DeprecatedFunction, string> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction, string> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch, string> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference, string> A;

		public static Func<ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula, string> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents A)
		{
			return A.NumberOfPrecedents;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators A)
		{
			return A.NumberOfOperators;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions A)
		{
			return A.NumberOfFunctions;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings A)
		{
			return A.NumberOfGroupings;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting A)
		{
			return A.NestingLevel;
		}

		[SpecialName]
		internal int A(ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity A)
		{
			return A.NumberOfConditions;
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.DeprecatedFunction A)
		{
			return A.FunctionName;
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction A)
		{
			return A.FunctionName;
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch A)
		{
			return A.FunctionName;
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference A)
		{
			return A.FunctionName;
		}

		[SpecialName]
		internal string A(ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula A)
		{
			return A.Formula;
		}
	}

	[CompilerGenerated]
	internal sealed class O
	{
		public Analysis A;

		public Settings A;

		public bool A;

		public Regex A;

		public Func<Range, int, Observation> A;

		public Func<Range, Observation> A;

		public Func<Range, int, Observation> B;

		public Func<Range, int, Observation> C;

		public Func<Range, int, Observation> D;

		public Func<Range, int, Observation> E;

		public Func<Range, Observation> B;

		public Func<Range, Observation> C;

		public Func<Range, Observation> D;

		public Func<Range, Observation> E;

		public Func<Range, int, Observation> F;

		public Func<Range, Observation> F;

		public Func<Range, Observation> G;

		public Func<Range, Observation> H;

		public Func<Range, string, Observation> A;

		public Func<Range, string, Observation> B;

		public Func<Range, string, Observation> C;

		public Func<Range, string, Observation> D;

		public Func<Range, Observation> I;

		public Func<Range, string, Observation> E;

		[SpecialName]
		internal void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(this.A, A, this.A, B, ref this.A);
		}

		[SpecialName]
		internal void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents> tooManyPrecedents = A.TooManyPrecedents;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents, int> c = _Closure_0024__.A.A;
			Func<Range, int, Observation> d;
			if (this.A != null)
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
				d = this.A;
			}
			else
			{
				d = (this.A = [SpecialName] (Range rng, int intPrecedents) => new ExcelAddIn1.Audit.Check.Observations.TooManyPrecedents(this.A.TooManyPrecedents, rng, intPrecedents));
			}
			Worksheet.A(a, tooManyPrecedents, c, d);
		}

		[SpecialName]
		internal Observation A(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.TooManyPrecedents(this.A.TooManyPrecedents, A, B);
		}

		[SpecialName]
		internal bool A()
		{
			return this.A.TooManyPrecedents != Severity.Ignore;
		}

		[SpecialName]
		internal void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(this.A, A, this.A.DoubleSums, B);
		}

		[SpecialName]
		internal void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> doubleSums = A.DoubleSums;
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
				c = (this.A = [SpecialName] (Range rng) => new DoubleSum(this.A.DoubleSums, rng));
			}
			Worksheet.A(a, doubleSums, c);
		}

		[SpecialName]
		internal Observation A(Range A)
		{
			return new DoubleSum(this.A.DoubleSums, A);
		}

		[SpecialName]
		internal bool B()
		{
			return false;
		}

		[SpecialName]
		internal void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(A, this.A, B, ref this.A);
		}

		[SpecialName]
		internal void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators> tooManyOperators = A.TooManyOperators;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					switch (4)
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
				d = (this.B = [SpecialName] (Range rng, int intOperators) => new ExcelAddIn1.Audit.Check.Observations.TooManyOperators(this.A.TooManyOperators, rng, intOperators));
			}
			Worksheet.A(a, tooManyOperators, c, d);
		}

		[SpecialName]
		internal Observation B(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.TooManyOperators(this.A.TooManyOperators, A, B);
		}

		[SpecialName]
		internal bool C()
		{
			return this.A.TooManyOperators != Severity.Ignore;
		}

		[SpecialName]
		internal void D(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.B(A, this.A, B, ref this.A);
		}

		[SpecialName]
		internal void D(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions> tooManyFunctions = A.TooManyFunctions;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				c = _Closure_0024__.A;
			}
			Func<Range, int, Observation> d;
			if (this.C != null)
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
				d = this.C;
			}
			else
			{
				d = (this.C = [SpecialName] (Range rng, int intFunctions) => new ExcelAddIn1.Audit.Check.Observations.TooManyFunctions(this.A.TooManyFunctions, rng, intFunctions));
			}
			Worksheet.A(a, tooManyFunctions, c, d);
		}

		[SpecialName]
		internal Observation C(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.TooManyFunctions(this.A.TooManyFunctions, A, B);
		}

		[SpecialName]
		internal bool D()
		{
			return this.A.TooManyFunctions != Severity.Ignore;
		}

		[SpecialName]
		internal void E(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.B(this.A, A, this.A, B, ref this.A);
		}

		[SpecialName]
		internal void E(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings> tooManyGroupings = A.TooManyGroupings;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				c = _Closure_0024__.A;
			}
			Func<Range, int, Observation> d;
			if (this.D != null)
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
				d = this.D;
			}
			else
			{
				d = (this.D = [SpecialName] (Range rng, int intGroupings) => new ExcelAddIn1.Audit.Check.Observations.TooManyGroupings(this.A.TooManyGroupings, rng, intGroupings));
			}
			Worksheet.A(a, tooManyGroupings, c, d);
		}

		[SpecialName]
		internal Observation D(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.TooManyGroupings(this.A.TooManyGroupings, A, B);
		}

		[SpecialName]
		internal bool E()
		{
			return this.A.TooManyGroupings != Severity.Ignore;
		}

		[SpecialName]
		internal void F(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.C(this.A, A, this.A, B, ref this.A);
		}

		[SpecialName]
		internal void F(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting> deepNesting = A.DeepNesting;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Worksheet.A(a, deepNesting, c, [SpecialName] (Range rng, int intLevels) => new ExcelAddIn1.Audit.Check.Observations.DeepNesting(this.A.DeepNesting, rng, intLevels));
		}

		[SpecialName]
		internal Observation E(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.DeepNesting(this.A.DeepNesting, A, B);
		}

		[SpecialName]
		internal bool F()
		{
			return this.A.DeepNesting != Severity.Ignore;
		}

		[SpecialName]
		internal void G(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(A, this.A, B);
		}

		[SpecialName]
		internal void G(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> formulasTooLong = A.FormulasTooLong;
			Func<Range, Observation> c;
			if (this.B != null)
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
				c = this.B;
			}
			else
			{
				c = (this.B = [SpecialName] (Range rng) => new FormulaTooLong(this.A.FormulasTooLong, rng));
			}
			Worksheet.A(a, formulasTooLong, c);
		}

		[SpecialName]
		internal Observation B(Range A)
		{
			return new FormulaTooLong(this.A.FormulasTooLong, A);
		}

		[SpecialName]
		internal bool G()
		{
			if (this.A.FormulasTooLong != Severity.Ignore)
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
						return this.A;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal void H(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(this.A, this.A.DuplicateFormulas, B, C);
		}

		[SpecialName]
		internal void H(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Formulas.A(this.A, A, B);
			Worksheet.A(this.A, A.DuplicateFormulas, [SpecialName] (Range rng) => new DuplicateFormula(this.A.DuplicateFormulas, rng));
		}

		[SpecialName]
		internal Observation C(Range A)
		{
			return new DuplicateFormula(this.A.DuplicateFormulas, A);
		}

		[SpecialName]
		internal bool H()
		{
			return this.A.DuplicateFormulas != Severity.Ignore;
		}

		[SpecialName]
		internal void I(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.B(A, this.A.OmittedReferences, B);
		}

		[SpecialName]
		internal void I(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> omittedReferences = A.OmittedReferences;
			Func<Range, Observation> c;
			if (this.D != null)
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
				c = this.D;
			}
			else
			{
				c = (this.D = [SpecialName] (Range rng) => new OmittedReference(this.A.OmittedReferences, rng));
			}
			Worksheet.A(a, omittedReferences, c);
		}

		[SpecialName]
		internal Observation D(Range A)
		{
			return new OmittedReference(this.A.OmittedReferences, A);
		}

		[SpecialName]
		internal bool I()
		{
			return this.A.OmittedReferences != Severity.Ignore;
		}

		[SpecialName]
		internal void J(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(A, this.A.EmptyCellReferences, B);
		}

		[SpecialName]
		internal void J(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> emptyCellReferences = A.EmptyCellReferences;
			Func<Range, Observation> c;
			if (this.E != null)
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
				c = this.E;
			}
			else
			{
				c = (this.E = [SpecialName] (Range rng) => new EmptyCellReference(this.A.EmptyCellReferences, rng));
			}
			Worksheet.A(a, emptyCellReferences, c);
		}

		[SpecialName]
		internal Observation E(Range A)
		{
			return new EmptyCellReference(this.A.EmptyCellReferences, A);
		}

		[SpecialName]
		internal bool J()
		{
			return this.A.EmptyCellReferences != Severity.Ignore;
		}

		[SpecialName]
		internal void K(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.B(A, this.A, B);
		}

		[SpecialName]
		internal void K(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity> conditionalComplexities = A.ConditionalComplexities;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity, int> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Func<Range, int, Observation> d;
			if (this.F != null)
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
				d = this.F;
			}
			else
			{
				d = (this.F = [SpecialName] (Range rng, int intIfs) => new ExcelAddIn1.Audit.Check.Observations.ConditionalComplexity(this.A.ConditionalComplexity, rng, intIfs));
			}
			Worksheet.A(a, conditionalComplexities, c, d);
		}

		[SpecialName]
		internal Observation F(Range A, int B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.ConditionalComplexity(this.A.ConditionalComplexity, A, B);
		}

		[SpecialName]
		internal bool K()
		{
			return this.A.ConditionalComplexity != Severity.Ignore;
		}

		[SpecialName]
		internal void L(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.C(A, this.A.PartialInputs, B);
		}

		[SpecialName]
		internal void L(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> partialInputs = A.PartialInputs;
			Func<Range, Observation> c;
			if (this.F != null)
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
				c = this.F;
			}
			else
			{
				c = (this.F = [SpecialName] (Range rng) => new PartialInput(this.A.PartialInputs, rng));
			}
			Worksheet.A(a, partialInputs, c);
		}

		[SpecialName]
		internal Observation F(Range A)
		{
			return new PartialInput(this.A.PartialInputs, A);
		}

		[SpecialName]
		internal bool L()
		{
			return this.A.PartialInputs != Severity.Ignore;
		}

		[SpecialName]
		internal void M(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.A(A, this.A.UnnecessaryFormulas, B, this.A);
		}

		[SpecialName]
		internal void M(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Worksheet.A(this.A, A.UnnecessaryFormulas, [SpecialName] (Range rng) => new UnnecessaryFormula(this.A.UnnecessaryFormulas, rng));
		}

		[SpecialName]
		internal Observation G(Range A)
		{
			return new UnnecessaryFormula(this.A.UnnecessaryFormulas, A);
		}

		[SpecialName]
		internal bool M()
		{
			return this.A.UnnecessaryFormulas != Severity.Ignore;
		}

		[SpecialName]
		internal void N(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.D(A, this.A.ExtraneousSheetNames, B);
		}

		[SpecialName]
		internal void N(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Worksheet.A(this.A, A.ExtraneousSheetNames, [SpecialName] (Range rng) => new ExtraneousSheetName(this.A.ExtraneousSheetNames, rng));
		}

		[SpecialName]
		internal Observation H(Range A)
		{
			return new ExtraneousSheetName(this.A.ExtraneousSheetNames, A);
		}

		[SpecialName]
		internal bool N()
		{
			return this.A.ExtraneousSheetNames != Severity.Ignore;
		}

		[SpecialName]
		internal void O(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.G(A, this.A.DeprecatedFunctions, B);
		}

		[SpecialName]
		internal void O(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Worksheet.A(this.A, A.DeprecatedFunctions, _Closure_0024__.A.A, [SpecialName] (Range rng, string strFunction) => new ExcelAddIn1.Audit.Check.Observations.DeprecatedFunction(this.A.DeprecatedFunctions, rng, strFunction));
		}

		[SpecialName]
		internal Observation A(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.DeprecatedFunction(this.A.DeprecatedFunctions, A, B);
		}

		[SpecialName]
		internal bool O()
		{
			return this.A.DeprecatedFunctions != Severity.Ignore;
		}

		[SpecialName]
		internal void P(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.H(A, this.A.VolatileFunctions, B);
		}

		[SpecialName]
		internal void P(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction> volatileFunctions = A.VolatileFunctions;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction, string> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Func<Range, string, Observation> d;
			if (this.B != null)
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
				d = this.B;
			}
			else
			{
				d = (this.B = [SpecialName] (Range rng, string strFunction) => new ExcelAddIn1.Audit.Check.Observations.VolatileFunction(this.A.VolatileFunctions, rng, strFunction));
			}
			Worksheet.A(a, volatileFunctions, c, d);
		}

		[SpecialName]
		internal Observation B(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.VolatileFunction(this.A.VolatileFunctions, A, B);
		}

		[SpecialName]
		internal bool P()
		{
			return this.A.VolatileFunctions != Severity.Ignore;
		}

		[SpecialName]
		internal void Q(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			List<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch> A2 = A.ApproximateMatches;
			AppoximateMatch.A(ref A2, this.A.ApproximateMatch, B);
			A.ApproximateMatches = A2;
		}

		[SpecialName]
		internal void Q(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch> approximateMatches = A.ApproximateMatches;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch, string> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			Func<Range, string, Observation> d;
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
				d = this.C;
			}
			else
			{
				d = (this.C = [SpecialName] (Range rng, string strFunctionName) => new ExcelAddIn1.Audit.Check.Observations.ApproximateMatch(this.A.ApproximateMatch, rng, strFunctionName));
			}
			Worksheet.A(a, approximateMatches, c, d);
		}

		[SpecialName]
		internal Observation C(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.ApproximateMatch(this.A.ApproximateMatch, A, B);
		}

		[SpecialName]
		internal bool Q()
		{
			return this.A.ApproximateMatch != Severity.Ignore;
		}

		[SpecialName]
		internal void R(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A2 = A.NumericIndexReferences;
			NumericIndexReference.A(ref A2, this.A.NumericIndexReference, B);
			A.NumericIndexReferences = A2;
		}

		[SpecialName]
		internal void R(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> numericIndexReferences = A.NumericIndexReferences;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference, string> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				c = _Closure_0024__.A;
			}
			Worksheet.A(a, numericIndexReferences, c, [SpecialName] (Range rng, string strFunctionName) => new ExcelAddIn1.Audit.Check.Observations.NumericIndexReference(this.A.NumericIndexReference, rng, strFunctionName));
		}

		[SpecialName]
		internal Observation D(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.NumericIndexReference(this.A.NumericIndexReference, A, B);
		}

		[SpecialName]
		internal bool R()
		{
			return this.A.NumericIndexReference != Severity.Ignore;
		}

		[SpecialName]
		internal void S(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			Formulas.F(A, this.A.DoubleMinus, B);
		}

		[SpecialName]
		internal void S(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<Range> doubleMinus = A.DoubleMinus;
			Func<Range, Observation> c;
			if (this.I != null)
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
				c = this.I;
			}
			else
			{
				c = (this.I = [SpecialName] (Range rng) => new DoubleMinus(this.A.DoubleMinus, rng));
			}
			Worksheet.A(a, doubleMinus, c);
		}

		[SpecialName]
		internal Observation I(Range A)
		{
			return new DoubleMinus(this.A.DoubleMinus, A);
		}

		[SpecialName]
		internal bool S()
		{
			return this.A.DoubleMinus != Severity.Ignore;
		}

		[SpecialName]
		internal void T(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Range B, Microsoft.Office.Interop.Excel.Worksheet C)
		{
			if (QB.A(B))
			{
				Formulas.E(A, this.A.LegacyArrayFormulas, B);
			}
		}

		[SpecialName]
		internal void T(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Microsoft.Office.Interop.Excel.Worksheet B)
		{
			Analysis a = this.A;
			List<ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula> legacyArrayFormulas = A.LegacyArrayFormulas;
			Func<ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula, string> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
				c = _Closure_0024__.A;
			}
			Func<Range, string, Observation> d;
			if (this.E != null)
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
				d = this.E;
			}
			else
			{
				d = (this.E = [SpecialName] (Range rng, string strFormula) => new ExcelAddIn1.Audit.Check.Observations.LegacyArrayFormula(this.A.LegacyArrayFormulas, rng, strFormula));
			}
			Worksheet.A(a, legacyArrayFormulas, c, d);
		}

		[SpecialName]
		internal Observation E(Range A, string B)
		{
			return new ExcelAddIn1.Audit.Check.Observations.LegacyArrayFormula(this.A.LegacyArrayFormulas, A, B);
		}

		[SpecialName]
		internal bool T()
		{
			return this.A.LegacyArrayFormulas != Severity.Ignore;
		}

		[SpecialName]
		internal void A(Analysis A)
		{
			A.PrecRetriever.A();
			A.DictWsFormulas.Clear();
			A.ParenthesisPairs.A();
			this.A = null;
		}
	}

	internal static List<RB> A(Analysis A, Settings B)
	{
		bool A2 = true;
		Regex A3 = new Regex(VH.A(2974));
		Func<Range, Observation> D = default(Func<Range, Observation>);
		Func<Range, Observation> F = default(Func<Range, Observation>);
		Func<Range, string, Observation> B2 = default(Func<Range, string, Observation>);
		Func<Range, string, Observation> C = default(Func<Range, string, Observation>);
		List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> A4 = default(List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference>);
		Func<Range, Observation> I = default(Func<Range, Observation>);
		Func<Range, string, Observation> E = default(Func<Range, string, Observation>);
		return new List<RB>
		{
			new ZB(A, VH.A(2987), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations b, Range d, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(A, b, B, d, ref A2);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents> tooManyPrecedents = observations.TooManyPrecedents;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents, int> c = _Closure_0024__.A.A;
				Func<Range, int, Observation> d;
				if (A4 != null)
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
					d = A4;
				}
				else
				{
					d = (A4 = [SpecialName] (Range rng, int intPrecedents) => new ExcelAddIn1.Audit.Check.Observations.TooManyPrecedents(B.TooManyPrecedents, rng, intPrecedents));
				}
				Worksheet.A(a, tooManyPrecedents, c, d);
			}, [SpecialName] () => B.TooManyPrecedents != Severity.Ignore),
			new ZB(A, VH.A(3026), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations b, Range d, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(A, b, B.DoubleSums, d);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> doubleSums = observations.DoubleSums;
				Func<Range, Observation> c;
				if (A4 != null)
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
					c = A4;
				}
				else
				{
					c = (A4 = [SpecialName] (Range rng) => new DoubleSum(B.DoubleSums, rng));
				}
				Worksheet.A(a, doubleSums, c);
			}, [SpecialName] () => false),
			new ZB(A, VH.A(3049), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(a, B, c, ref A2);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators> tooManyOperators = observations.TooManyOperators;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
						switch (4)
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
					d = (B2 = [SpecialName] (Range rng, int intOperators) => new ExcelAddIn1.Audit.Check.Observations.TooManyOperators(B.TooManyOperators, rng, intOperators));
				}
				Worksheet.A(a, tooManyOperators, c, d);
			}, [SpecialName] () => B.TooManyOperators != Severity.Ignore),
			new ZB(A, VH.A(3086), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.B(a, B, c, ref A2);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions> tooManyFunctions = observations.TooManyFunctions;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					c = _Closure_0024__.A;
				}
				Func<Range, int, Observation> d;
				if (C != null)
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
					d = C;
				}
				else
				{
					d = (C = [SpecialName] (Range rng, int intFunctions) => new ExcelAddIn1.Audit.Check.Observations.TooManyFunctions(B.TooManyFunctions, rng, intFunctions));
				}
				Worksheet.A(a, tooManyFunctions, c, d);
			}, [SpecialName] () => B.TooManyFunctions != Severity.Ignore),
			new ZB(A, VH.A(3123), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations b, Range d, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.B(A, b, B, d, ref A2);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings> tooManyGroupings = observations.TooManyGroupings;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					c = _Closure_0024__.A;
				}
				Func<Range, int, Observation> d;
				if (D != null)
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
					d = D;
				}
				else
				{
					d = (D = [SpecialName] (Range rng, int intGroupings) => new ExcelAddIn1.Audit.Check.Observations.TooManyGroupings(B.TooManyGroupings, rng, intGroupings));
				}
				Worksheet.A(a, tooManyGroupings, c, d);
			}, [SpecialName] () => B.TooManyGroupings != Severity.Ignore),
			new ZB(A, VH.A(3160), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations b, Range d, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.C(A, b, B, d, ref A2);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting> deepNesting = observations.DeepNesting;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Worksheet.A(a, deepNesting, c, [SpecialName] (Range rng, int intLevels) => new ExcelAddIn1.Audit.Check.Observations.DeepNesting(B.DeepNesting, rng, intLevels));
			}, [SpecialName] () => B.DeepNesting != Severity.Ignore),
			new ZB(A, VH.A(3185), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(a, B, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> formulasTooLong = observations.FormulasTooLong;
				Func<Range, Observation> c;
				if (B2 != null)
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
					c = B2;
				}
				else
				{
					c = (B2 = [SpecialName] (Range rng) => new FormulaTooLong(B.FormulasTooLong, rng));
				}
				Worksheet.A(a, formulasTooLong, c);
			}, [SpecialName] () =>
			{
				if (B.FormulasTooLong != Severity.Ignore)
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
							return A2;
						}
					}
				}
				return false;
			}),
			new ZB(A, VH.A(3212), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(A, B.DuplicateFormulas, c, C);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet c) =>
			{
				Formulas.A(A, observations, c);
				Worksheet.A(A, observations.DuplicateFormulas, [SpecialName] (Range rng) => new DuplicateFormula(B.DuplicateFormulas, rng));
			}, [SpecialName] () => B.DuplicateFormulas != Severity.Ignore),
			new ZB(A, VH.A(3249), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.B(a, B.OmittedReferences, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> omittedReferences = observations.OmittedReferences;
				Func<Range, Observation> c;
				if (D != null)
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
					c = D;
				}
				else
				{
					c = (D = [SpecialName] (Range rng) => new OmittedReference(B.OmittedReferences, rng));
				}
				Worksheet.A(a, omittedReferences, c);
			}, [SpecialName] () => B.OmittedReferences != Severity.Ignore),
			new ZB(A, VH.A(3286), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(a, B.EmptyCellReferences, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> emptyCellReferences = observations.EmptyCellReferences;
				Func<Range, Observation> c;
				if (E != null)
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
					c = E;
				}
				else
				{
					c = (E = [SpecialName] (Range rng) => new EmptyCellReference(B.EmptyCellReferences, rng));
				}
				Worksheet.A(a, emptyCellReferences, c);
			}, [SpecialName] () => B.EmptyCellReferences != Severity.Ignore),
			new ZB(A, VH.A(3329), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.B(a, B, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity> conditionalComplexities = observations.ConditionalComplexities;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity, int> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Func<Range, int, Observation> d;
				if (F != null)
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
					d = F;
				}
				else
				{
					d = (F = [SpecialName] (Range rng, int intIfs) => new ExcelAddIn1.Audit.Check.Observations.ConditionalComplexity(B.ConditionalComplexity, rng, intIfs));
				}
				Worksheet.A(a, conditionalComplexities, c, d);
			}, [SpecialName] () => B.ConditionalComplexity != Severity.Ignore),
			new ZB(A, VH.A(3374), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.C(a, B.PartialInputs, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> partialInputs = observations.PartialInputs;
				Func<Range, Observation> c;
				if (F != null)
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
					c = F;
				}
				else
				{
					c = (F = [SpecialName] (Range rng) => new PartialInput(B.PartialInputs, rng));
				}
				Worksheet.A(a, partialInputs, c);
			}, [SpecialName] () => B.PartialInputs != Severity.Ignore),
			new ZB(A, VH.A(3403), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.A(a, B.UnnecessaryFormulas, c, A3);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Worksheet.A(A, observations.UnnecessaryFormulas, [SpecialName] (Range rng) => new UnnecessaryFormula(B.UnnecessaryFormulas, rng));
			}, [SpecialName] () => B.UnnecessaryFormulas != Severity.Ignore),
			new ZB(A, VH.A(3444), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.D(a, B.ExtraneousSheetNames, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Worksheet.A(A, observations.ExtraneousSheetNames, [SpecialName] (Range rng) => new ExtraneousSheetName(B.ExtraneousSheetNames, rng));
			}, [SpecialName] () => B.ExtraneousSheetNames != Severity.Ignore),
			new ZB(A, VH.A(3487), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				G(a, B.DeprecatedFunctions, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Worksheet.A(A, observations.DeprecatedFunctions, _Closure_0024__.A.A, [SpecialName] (Range rng, string strFunction) => new ExcelAddIn1.Audit.Check.Observations.DeprecatedFunction(B.DeprecatedFunctions, rng, strFunction));
			}, [SpecialName] () => B.DeprecatedFunctions != Severity.Ignore),
			new ZB(A, VH.A(3528), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				H(a, B.VolatileFunctions, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction> volatileFunctions = observations.VolatileFunctions;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction, string> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Func<Range, string, Observation> d;
				if (B2 != null)
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
					d = B2;
				}
				else
				{
					d = (B2 = [SpecialName] (Range rng, string strFunction) => new ExcelAddIn1.Audit.Check.Observations.VolatileFunction(B.VolatileFunctions, rng, strFunction));
				}
				Worksheet.A(a, volatileFunctions, c, d);
			}, [SpecialName] () => B.VolatileFunctions != Severity.Ignore),
			new ZB(A, VH.A(3565), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				A4 = observations.ApproximateMatches;
				AppoximateMatch.A(ref A4, B.ApproximateMatch, c);
				observations.ApproximateMatches = A4;
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch> approximateMatches = observations.ApproximateMatches;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.ApproximateMatch, string> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					c = _Closure_0024__.A;
				}
				Func<Range, string, Observation> d;
				if (C != null)
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
					d = C;
				}
				else
				{
					d = (C = [SpecialName] (Range rng, string strFunctionName) => new ExcelAddIn1.Audit.Check.Observations.ApproximateMatch(B.ApproximateMatch, rng, strFunctionName));
				}
				Worksheet.A(a, approximateMatches, c, d);
			}, [SpecialName] () => B.ApproximateMatch != Severity.Ignore),
			new ZB(A, VH.A(3604), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				A4 = observations.NumericIndexReferences;
				NumericIndexReference.A(ref A4, B.NumericIndexReference, c);
				observations.NumericIndexReferences = A4;
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference> numericIndexReferences = observations.NumericIndexReferences;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.NumericIndexReference, string> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					c = _Closure_0024__.A;
				}
				Worksheet.A(a, numericIndexReferences, c, [SpecialName] (Range rng, string strFunctionName) => new ExcelAddIn1.Audit.Check.Observations.NumericIndexReference(B.NumericIndexReference, rng, strFunctionName));
			}, [SpecialName] () => B.NumericIndexReference != Severity.Ignore),
			new ZB(A, VH.A(3653), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range c, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				Formulas.F(a, B.DoubleMinus, c);
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<Range> doubleMinus = observations.DoubleMinus;
				Func<Range, Observation> c;
				if (I != null)
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
					c = I;
				}
				else
				{
					c = (I = [SpecialName] (Range rng) => new DoubleMinus(B.DoubleMinus, rng));
				}
				Worksheet.A(a, doubleMinus, c);
			}, [SpecialName] () => B.DoubleMinus != Severity.Ignore),
			new ZB(A, VH.A(3678), [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations a, Range range, Microsoft.Office.Interop.Excel.Worksheet C) =>
			{
				if (QB.A(range))
				{
					Formulas.E(a, B.LegacyArrayFormulas, range);
				}
			}, [SpecialName] (ExcelAddIn1.Audit.Check.Observations.Raw.Observations observations, Microsoft.Office.Interop.Excel.Worksheet worksheet) =>
			{
				Analysis a = A;
				List<ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula> legacyArrayFormulas = observations.LegacyArrayFormulas;
				Func<ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula, string> c;
				if (_Closure_0024__.A == null)
				{
					c = (_Closure_0024__.A = _Closure_0024__.A.A);
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
					c = _Closure_0024__.A;
				}
				Func<Range, string, Observation> d;
				if (E != null)
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
					d = E;
				}
				else
				{
					d = (E = [SpecialName] (Range rng, string strFormula) => new ExcelAddIn1.Audit.Check.Observations.LegacyArrayFormula(B.LegacyArrayFormulas, rng, strFormula));
				}
				Worksheet.A(a, legacyArrayFormulas, c, d);
			}, [SpecialName] () => B.LegacyArrayFormulas != Severity.Ignore),
			new SB([SpecialName] (Analysis analysis) =>
			{
				analysis.PrecRetriever.A();
				analysis.DictWsFormulas.Clear();
				analysis.ParenthesisPairs.A();
				A3 = null;
			})
		};
	}

	internal static void A(List<Observation> A, Severity B, Range C)
	{
		if (C == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
				Range range = RangeHelpers.C(C);
				if (range == null)
				{
					return;
				}
				try
				{
					enumerator = range.Areas.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range rng = (Range)enumerator.Current;
						A.Add(new FormulaError(B, rng));
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
				range = null;
				return;
			}
		}
	}

	private static void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Settings B, Range C)
	{
		if (B.FormulasTooLong == Severity.Ignore || QB.B(C).Length <= B.MaxFormulaLength)
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
			A.FormulasTooLong.Add(C);
			return;
		}
	}

	private static void A(Analysis A, ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Settings C, Range D, ref bool E)
	{
		if (C.TooManyPrecedents == Severity.Ignore)
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
			int num = A.PrecRetriever.A(D);
			if (num <= C.MaxNumberOfPrecedents)
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
				B.TooManyPrecedents.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.TooManyPrecedents(D, num));
				E = false;
				return;
			}
		}
	}

	private static void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
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
			if (!C.Errors.get_Item((object)XlErrorChecks.xlEmptyCellReferences).Value)
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
				A.EmptyCellReferences.Add(C);
				return;
			}
		}
	}

	private static void A(Analysis A, ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Severity C, Range D)
	{
		if (C == Severity.Ignore)
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
			ExcelAddIn1.Formulas.Helpers.IsFunctionMatch(D, VH.A(3721));
			return;
		}
	}

	private static void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
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
			if (!C.Errors.get_Item((object)XlErrorChecks.xlOmittedCells).Value)
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
				A.OmittedReferences.Add(C);
				return;
			}
		}
	}

	private static void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Settings B, Range C, ref bool D)
	{
		if (B.TooManyOperators == Severity.Ignore)
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
			int count = Regex.Matches(QB.B(C), VH.A(3728)).Count;
			if (count <= B.MaxNumberOfOperators)
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
				A.TooManyOperators.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.TooManyOperators(C, count));
				D = false;
				return;
			}
		}
	}

	private static void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Settings B, Range C, ref bool D)
	{
		if (B.TooManyFunctions == Severity.Ignore)
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
			int count = Regex.Matches(QB.A(C), VH.A(3753)).Count;
			if (count <= B.MaxNumberOfFunctions)
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
				A.TooManyFunctions.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.TooManyFunctions(C, count));
				D = false;
				return;
			}
		}
	}

	private static void B(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Settings B, Range C)
	{
		if (B.ConditionalComplexity == Severity.Ignore)
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
			int count = ExcelAddIn1.Formulas.Helpers.FunctionMatches(C, VH.A(3794)).Count;
			if (count <= B.MaxNumberOfIfs)
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
				A.ConditionalComplexities.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.ConditionalComplexity(C, count));
				return;
			}
		}
	}

	private static void B(Analysis A, ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Settings C, Range D, ref bool E)
	{
		if (C.TooManyGroupings == Severity.Ignore)
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
			int num = A.ParenthesisPairs.A(D);
			if (num <= C.MaxNumberOfGroupings)
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
				B.TooManyGroupings.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.TooManyGroupings(D, num));
				E = false;
				return;
			}
		}
	}

	private static void C(Analysis A, ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Settings C, Range D, ref bool E)
	{
		if (C.DeepNesting == Severity.Ignore)
		{
			return;
		}
		FB.EB eB = A.ParenthesisPairs.A(D);
		if (Operators.CompareString(eB.MaskedFormula, "", TextCompare: false) == 0)
		{
			return;
		}
		int length = eB.MaskedFormula.Length;
		List<int> list = new List<int>(length);
		checked
		{
			length--;
			int num = length;
			for (int i = 0; i <= num; i++)
			{
				list.Add(0);
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
				using (List<ParenthesesPair>.Enumerator enumerator = eB.Pairs.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						ParenthesesPair current = enumerator.Current;
						try
						{
							length = current.StartIndex + current.Length - 1;
							int startIndex = current.StartIndex;
							int num2 = length;
							for (int j = startIndex; j <= num2; j++)
							{
								list[j]++;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_00e9;
								}
								continue;
								end_IL_00e9:
								break;
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_010d;
						}
						continue;
						end_IL_010d:
						break;
					}
				}
				int num3 = list.Max();
				if (num3 > C.MaxNestingLevel)
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
					B.DeepNesting.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.DeepNesting(D, num3));
					E = false;
				}
				list = null;
				return;
			}
		}
	}

	private static void C(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
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
			if (!ExcelAddIn1.Formulas.Helpers.ContainsPartialInput(C))
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
				A.PartialInputs.Add(C);
				return;
			}
		}
	}

	private static void A(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C, Regex D)
	{
		if (B == Severity.Ignore)
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
			if (!D.IsMatch(QB.A(C)))
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
				A.UnnecessaryFormulas.Add(C);
				return;
			}
		}
	}

	private static void D(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
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
			string text = QB.A(C);
			if (Operators.CompareString(ExcelAddIn1.Formulas.Helpers.RemoveExtraneousSheetName(text, C.Worksheet.Name), text, TextCompare: false) == 0)
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
				A.ExtraneousSheetNames.Add(C);
				return;
			}
		}
	}

	private static void E(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
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
			A.LegacyArrayFormulas.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.LegacyArrayFormula(C));
			return;
		}
	}

	private static void F(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
	{
		if (B == Severity.Ignore)
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
			if (!QB.A(C).Contains(VH.A(3799)))
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
				A.DoubleMinus.Add(C);
				return;
			}
		}
	}

	private static void A(Analysis A, Severity B, Range C, Microsoft.Office.Interop.Excel.Worksheet D)
	{
		if (B == Severity.Ignore || QB.A(C))
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
			int num = A.PrecRetriever.A(C);
			bool flag;
			if (num > 0)
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
				flag = false;
				num = num;
			}
			else
			{
				flag = true;
				try
				{
					num = checked((int)Math.Min(Conversions.ToLong(C.DirectPrecedents.CountLarge), 2147483647L));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (num <= 1)
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
				string text = QB.A(C);
				if (flag)
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
					if (text.Length <= 10)
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
				}
				PB.A(PB.A(A.DictWsFormulas, D.Name), text).Add(C);
				return;
			}
		}
	}

	private static void A(Analysis A, ExcelAddIn1.Audit.Check.Observations.Raw.Observations B, Microsoft.Office.Interop.Excel.Worksheet C)
	{
		try
		{
			checked
			{
				using Dictionary<string, List<Range>>.Enumerator enumerator = PB.A(A.DictWsFormulas, C.Name).GetEnumerator();
				while (enumerator.MoveNext())
				{
					KeyValuePair<string, List<Range>> current = enumerator.Current;
					if (current.Value.Count <= 1)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					int num = current.Value.Count - 1;
					for (int i = 1; i <= num; i++)
					{
						B.DuplicateFormulas.Add(current.Value[i]);
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
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
		}
		catch (object obj) when (((Func<bool>)delegate
		{
			// Could not convert BlockContainer to single expression
			OutOfMemoryException obj2 = obj as OutOfMemoryException;
			System.Runtime.CompilerServices.Unsafe.SkipInit(out int result);
			if (obj2 == null)
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
				result = 0;
			}
			else
			{
				ProjectData.SetProjectError(obj2);
				result = ((obj2.HResult == -2147024882) ? 1 : 0);
			}
			return (byte)result != 0;
		}).Invoke())
		{
			ProjectData.ClearProjectError();
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			throw;
		}
		finally
		{
			A.DictWsFormulas.Remove(C.Name);
		}
	}

	private static void G(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
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
			string strFormula = QB.A(C);
			string[] array = new string[42]
			{
				VH.A(3804),
				VH.A(3821),
				VH.A(3836),
				VH.A(3855),
				VH.A(3870),
				VH.A(3885),
				VH.A(3898),
				VH.A(3913),
				VH.A(3936),
				VH.A(3957),
				VH.A(3968),
				VH.A(3987),
				VH.A(4006),
				VH.A(4017),
				VH.A(4026),
				VH.A(4037),
				VH.A(4054),
				VH.A(4065),
				VH.A(4084),
				VH.A(4101),
				VH.A(4124),
				VH.A(4137),
				VH.A(4160),
				VH.A(4169),
				VH.A(4194),
				VH.A(4211),
				VH.A(4226),
				VH.A(4245),
				VH.A(4262),
				VH.A(4283),
				VH.A(4306),
				VH.A(4321),
				VH.A(4338),
				VH.A(4347),
				VH.A(4358),
				VH.A(4371),
				VH.A(4382),
				VH.A(4391),
				VH.A(4402),
				VH.A(4409),
				VH.A(4418),
				VH.A(4433)
			};
			foreach (string strFunction in array)
			{
				if (ExcelAddIn1.Formulas.Helpers.IsFunctionMatch(strFormula, strFunction))
				{
					A.DeprecatedFunctions.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.DeprecatedFunction(C, strFunction));
				}
			}
			return;
		}
	}

	private static void H(ExcelAddIn1.Audit.Check.Observations.Raw.Observations A, Severity B, Range C)
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
			string strFormula = QB.A(C);
			string[] array = new string[8]
			{
				VH.A(4444),
				VH.A(4457),
				VH.A(4474),
				VH.A(4497),
				VH.A(4504),
				VH.A(4515),
				VH.A(4524),
				VH.A(4533)
			};
			foreach (string strFunction in array)
			{
				if (!ExcelAddIn1.Formulas.Helpers.IsFunctionMatch(strFormula, strFunction))
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
				A.VolatileFunctions.Add(new ExcelAddIn1.Audit.Check.Observations.Raw.VolatileFunction(C, strFunction));
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}

	private static void A(ref List<Observation> A, Settings B, Range C)
	{
	}
}
