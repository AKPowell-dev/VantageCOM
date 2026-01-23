using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class Observations
{
	[CompilerGenerated]
	private List<Range> A;

	[CompilerGenerated]
	private List<Range> B;

	[CompilerGenerated]
	private List<Range> C;

	[CompilerGenerated]
	private List<Range> D;

	[CompilerGenerated]
	private List<Range> E;

	[CompilerGenerated]
	private List<TooManyPrecedents> A;

	[CompilerGenerated]
	private List<Range> F;

	[CompilerGenerated]
	private List<TooManyOperators> A;

	[CompilerGenerated]
	private List<TooManyFunctions> A;

	[CompilerGenerated]
	private List<TooManyGroupings> A;

	[CompilerGenerated]
	private List<DeepNesting> A;

	[CompilerGenerated]
	private List<Range> G;

	[CompilerGenerated]
	private List<LegacyArrayFormula> A;

	[CompilerGenerated]
	private List<ConditionalComplexity> A;

	[CompilerGenerated]
	private List<Range> H;

	[CompilerGenerated]
	private List<Range> I;

	[CompilerGenerated]
	private List<Range> J;

	[CompilerGenerated]
	private List<DeprecatedFunction> A;

	[CompilerGenerated]
	private List<VolatileFunction> A;

	[CompilerGenerated]
	private List<Range> K;

	[CompilerGenerated]
	private List<ApproximateMatch> A;

	[CompilerGenerated]
	private List<NumericIndexReference> A;

	[CompilerGenerated]
	private List<Range> L;

	[CompilerGenerated]
	private List<CellFillColor> A;

	[CompilerGenerated]
	private List<CellBorderColor> A;

	[CompilerGenerated]
	private List<SensitiveData> A;

	[CompilerGenerated]
	private List<Range> M;

	[CompilerGenerated]
	private List<Range> N;

	internal List<Range> EmptyCellReferences
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

	internal List<Range> EmptyCellComments
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

	internal List<Range> EmptyCellNotes
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal List<Range> DoubleSums
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

	internal List<Range> OmittedReferences
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal List<TooManyPrecedents> TooManyPrecedents
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

	internal List<Range> FormulasTooLong
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal List<TooManyOperators> TooManyOperators
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

	internal List<TooManyFunctions> TooManyFunctions
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

	internal List<TooManyGroupings> TooManyGroupings
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

	internal List<DeepNesting> DeepNesting
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

	internal List<Range> DuplicateFormulas
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal List<LegacyArrayFormula> LegacyArrayFormulas
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

	internal List<ConditionalComplexity> ConditionalComplexities
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

	internal List<Range> PartialInputs
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	internal List<Range> UnnecessaryFormulas
	{
		[CompilerGenerated]
		get
		{
			return I;
		}
		[CompilerGenerated]
		set
		{
			I = value;
		}
	}

	internal List<Range> ExtraneousSheetNames
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
		[CompilerGenerated]
		set
		{
			J = value;
		}
	}

	internal List<DeprecatedFunction> DeprecatedFunctions
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

	internal List<VolatileFunction> VolatileFunctions
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

	internal List<Range> DoubleMinus
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[CompilerGenerated]
		set
		{
			K = value;
		}
	}

	internal List<ApproximateMatch> ApproximateMatches
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

	internal List<NumericIndexReference> NumericIndexReferences
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

	internal List<Range> TripleSemicolons
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
		[CompilerGenerated]
		set
		{
			L = value;
		}
	}

	internal List<CellFillColor> CellFillColors
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

	internal List<CellBorderColor> CellBorderColors
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

	internal List<SensitiveData> SensitiveData
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

	internal List<Range> NumbersStoredAsText
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
		[CompilerGenerated]
		set
		{
			M = value;
		}
	}

	internal List<Range> DataValidationFailed
	{
		[CompilerGenerated]
		get
		{
			return N;
		}
		[CompilerGenerated]
		set
		{
			N = value;
		}
	}

	internal Observations()
	{
		EmptyCellReferences = new List<Range>();
		EmptyCellComments = new List<Range>();
		EmptyCellNotes = new List<Range>();
		DoubleSums = new List<Range>();
		OmittedReferences = new List<Range>();
		TooManyPrecedents = new List<TooManyPrecedents>();
		FormulasTooLong = new List<Range>();
		TooManyOperators = new List<TooManyOperators>();
		TooManyFunctions = new List<TooManyFunctions>();
		TooManyGroupings = new List<TooManyGroupings>();
		DeepNesting = new List<DeepNesting>();
		DuplicateFormulas = new List<Range>();
		LegacyArrayFormulas = new List<LegacyArrayFormula>();
		ConditionalComplexities = new List<ConditionalComplexity>();
		PartialInputs = new List<Range>();
		UnnecessaryFormulas = new List<Range>();
		ExtraneousSheetNames = new List<Range>();
		DeprecatedFunctions = new List<DeprecatedFunction>();
		VolatileFunctions = new List<VolatileFunction>();
		DoubleMinus = new List<Range>();
		ApproximateMatches = new List<ApproximateMatch>();
		NumericIndexReferences = new List<NumericIndexReference>();
		TripleSemicolons = new List<Range>();
		CellFillColors = new List<CellFillColor>();
		CellBorderColors = new List<CellBorderColor>();
		SensitiveData = new List<SensitiveData>();
		NumbersStoredAsText = new List<Range>();
		DataValidationFailed = new List<Range>();
	}
}
