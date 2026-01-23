using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros.Config.Settings;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check;

public sealed class Settings
{
	private readonly string m_A;

	internal readonly string B;

	internal readonly string C;

	internal readonly string D;

	internal readonly string E;

	internal readonly string F;

	internal readonly string G;

	[CompilerGenerated]
	private string m_H;

	internal readonly string I;

	internal readonly string J;

	internal readonly string K;

	internal readonly string L;

	internal readonly string M;

	internal readonly string N;

	internal readonly string O;

	internal readonly string P;

	internal readonly string Q;

	internal readonly string R;

	internal readonly string S;

	internal readonly string T;

	internal readonly string U;

	internal readonly string V;

	internal readonly string W;

	internal readonly string X;

	internal readonly string Y;

	internal readonly string Z;

	internal readonly string AB;

	internal readonly string BB;

	internal readonly string CB;

	internal readonly string DB;

	internal readonly string EB;

	internal readonly string FB;

	internal readonly string GB;

	internal readonly string HB;

	internal readonly string IB;

	internal readonly string JB;

	internal readonly string KB;

	internal readonly string LB;

	internal readonly string MB;

	internal readonly string NB;

	internal readonly string OB;

	internal readonly string PB;

	internal readonly string QB;

	internal readonly string RB;

	internal readonly string SB;

	internal readonly string TB;

	internal readonly string UB;

	internal readonly string VB;

	internal readonly string WB;

	internal readonly string XB;

	internal readonly string YB;

	internal readonly string ZB;

	internal readonly string AC;

	internal readonly string BC;

	[CompilerGenerated]
	private Severity m_A;

	[CompilerGenerated]
	private Severity m_B;

	[CompilerGenerated]
	private Severity m_C;

	[CompilerGenerated]
	private Severity m_D;

	[CompilerGenerated]
	private Severity m_E;

	[CompilerGenerated]
	private Severity m_F;

	[CompilerGenerated]
	private Severity m_G;

	[CompilerGenerated]
	private Severity m_H;

	[CompilerGenerated]
	private Severity m_I;

	[CompilerGenerated]
	private Severity m_J;

	[CompilerGenerated]
	private Severity m_K;

	[CompilerGenerated]
	private Severity L;

	[CompilerGenerated]
	private Severity M;

	[CompilerGenerated]
	private Severity N;

	[CompilerGenerated]
	private Severity O;

	[CompilerGenerated]
	private Severity P;

	[CompilerGenerated]
	private Severity Q;

	[CompilerGenerated]
	private Severity R;

	[CompilerGenerated]
	private Severity S;

	[CompilerGenerated]
	private Severity T;

	[CompilerGenerated]
	private Severity U;

	[CompilerGenerated]
	private Severity V;

	[CompilerGenerated]
	private Severity W;

	[CompilerGenerated]
	private Severity X;

	[CompilerGenerated]
	private Severity Y;

	[CompilerGenerated]
	private Severity Z;

	[CompilerGenerated]
	private Severity AB;

	[CompilerGenerated]
	private Severity BB;

	[CompilerGenerated]
	private Severity CB;

	[CompilerGenerated]
	private Severity DB;

	[CompilerGenerated]
	private Severity EB;

	[CompilerGenerated]
	private Severity FB;

	[CompilerGenerated]
	private Severity GB;

	[CompilerGenerated]
	private Severity HB;

	[CompilerGenerated]
	private Severity IB;

	[CompilerGenerated]
	private Severity JB;

	[CompilerGenerated]
	private Severity KB;

	[CompilerGenerated]
	private Severity LB;

	[CompilerGenerated]
	private Severity MB;

	[CompilerGenerated]
	private Severity NB;

	[CompilerGenerated]
	private Severity OB;

	[CompilerGenerated]
	private Severity PB;

	[CompilerGenerated]
	private Severity QB;

	[CompilerGenerated]
	private Severity RB;

	[CompilerGenerated]
	private Severity SB;

	[CompilerGenerated]
	private Severity TB;

	[CompilerGenerated]
	private Severity UB;

	[CompilerGenerated]
	private Severity VB;

	[CompilerGenerated]
	private Severity WB;

	[CompilerGenerated]
	private Severity XB;

	[CompilerGenerated]
	private Severity YB;

	[CompilerGenerated]
	private Severity ZB;

	[CompilerGenerated]
	private Severity AC;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private int m_C;

	[CompilerGenerated]
	private int m_D;

	[CompilerGenerated]
	private int m_E;

	[CompilerGenerated]
	private int m_F;

	[CompilerGenerated]
	private int m_G;

	[CompilerGenerated]
	private int m_H;

	[CompilerGenerated]
	private int m_I;

	[CompilerGenerated]
	private long m_A;

	[CompilerGenerated]
	private int m_J;

	[CompilerGenerated]
	private static int m_K = 50;

	internal string ID_UNNECESSARY_FMLA
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal Severity CoverMissing
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

	internal Severity FormulaErrors
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal Severity EmptyCellReferences
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal Severity EmptyCellCommentsNotes
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal Severity UnusedNumericInputs
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal Severity PartialInputs
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal Severity UnnecessaryFormulas
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal Severity FormulaInterruption
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal Severity FormulasTooLong
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal Severity TooManyPrecedents
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal Severity TooManyOperators
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal Severity TooManyFunctions
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

	internal Severity TooManyGroupings
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

	internal Severity ConditionalComplexity
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

	internal Severity DuplicateFormulas
	{
		[CompilerGenerated]
		get
		{
			return O;
		}
		[CompilerGenerated]
		set
		{
			O = value;
		}
	}

	internal Severity DeepNesting
	{
		[CompilerGenerated]
		get
		{
			return P;
		}
		[CompilerGenerated]
		set
		{
			P = value;
		}
	}

	internal Severity ExtraneousSheetNames
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	internal Severity LegacyArrayFormulas
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	internal Severity VolatileFunctions
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	internal Severity DeprecatedFunctions
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	internal Severity ApproximateMatch
	{
		[CompilerGenerated]
		get
		{
			return U;
		}
		[CompilerGenerated]
		set
		{
			U = value;
		}
	}

	internal Severity NumericIndexReference
	{
		[CompilerGenerated]
		get
		{
			return V;
		}
		[CompilerGenerated]
		set
		{
			V = value;
		}
	}

	internal Severity DoubleMinus
	{
		[CompilerGenerated]
		get
		{
			return W;
		}
		[CompilerGenerated]
		set
		{
			W = value;
		}
	}

	internal Severity DoubleSums
	{
		[CompilerGenerated]
		get
		{
			return X;
		}
		[CompilerGenerated]
		set
		{
			X = value;
		}
	}

	internal Severity OmittedReferences
	{
		[CompilerGenerated]
		get
		{
			return Y;
		}
		[CompilerGenerated]
		set
		{
			Y = value;
		}
	}

	internal Severity CircularReferences
	{
		[CompilerGenerated]
		get
		{
			return Z;
		}
		[CompilerGenerated]
		set
		{
			Z = value;
		}
	}

	internal Severity InputsNotColored
	{
		[CompilerGenerated]
		get
		{
			return AB;
		}
		[CompilerGenerated]
		set
		{
			AB = value;
		}
	}

	internal Severity MergedCells
	{
		[CompilerGenerated]
		get
		{
			return BB;
		}
		[CompilerGenerated]
		set
		{
			BB = value;
		}
	}

	internal Severity ExcessConditionalFormatting
	{
		[CompilerGenerated]
		get
		{
			return CB;
		}
		[CompilerGenerated]
		set
		{
			CB = value;
		}
	}

	internal Severity TripleSemicolonNumFormat
	{
		[CompilerGenerated]
		get
		{
			return DB;
		}
		[CompilerGenerated]
		set
		{
			DB = value;
		}
	}

	internal Severity SensitiveData
	{
		[CompilerGenerated]
		get
		{
			return EB;
		}
		[CompilerGenerated]
		set
		{
			EB = value;
		}
	}

	internal Severity CommentsAndNotes
	{
		[CompilerGenerated]
		get
		{
			return FB;
		}
		[CompilerGenerated]
		set
		{
			FB = value;
		}
	}

	internal Severity HiddenSheets
	{
		[CompilerGenerated]
		get
		{
			return GB;
		}
		[CompilerGenerated]
		set
		{
			GB = value;
		}
	}

	internal Severity VeryHiddenSheets
	{
		[CompilerGenerated]
		get
		{
			return HB;
		}
		[CompilerGenerated]
		set
		{
			HB = value;
		}
	}

	internal Severity HiddenRowsColumns
	{
		[CompilerGenerated]
		get
		{
			return IB;
		}
		[CompilerGenerated]
		set
		{
			IB = value;
		}
	}

	internal Severity CollapsedRowsColumns
	{
		[CompilerGenerated]
		get
		{
			return JB;
		}
		[CompilerGenerated]
		set
		{
			JB = value;
		}
	}

	internal Severity OldFile
	{
		[CompilerGenerated]
		get
		{
			return KB;
		}
		[CompilerGenerated]
		set
		{
			KB = value;
		}
	}

	internal Severity LegacyFileType
	{
		[CompilerGenerated]
		get
		{
			return LB;
		}
		[CompilerGenerated]
		set
		{
			LB = value;
		}
	}

	internal Severity LargeFileSize
	{
		[CompilerGenerated]
		get
		{
			return MB;
		}
		[CompilerGenerated]
		set
		{
			MB = value;
		}
	}

	internal Severity CalculationModeManual
	{
		[CompilerGenerated]
		get
		{
			return NB;
		}
		[CompilerGenerated]
		set
		{
			NB = value;
		}
	}

	internal Severity DisplayDrawingObjects
	{
		[CompilerGenerated]
		get
		{
			return OB;
		}
		[CompilerGenerated]
		set
		{
			OB = value;
		}
	}

	internal Severity ShapesOverNonEmptyCells
	{
		[CompilerGenerated]
		get
		{
			return PB;
		}
		[CompilerGenerated]
		set
		{
			PB = value;
		}
	}

	internal Severity NamesWithExternalReferences
	{
		[CompilerGenerated]
		get
		{
			return QB;
		}
		[CompilerGenerated]
		set
		{
			QB = value;
		}
	}

	internal Severity ExcessNames
	{
		[CompilerGenerated]
		get
		{
			return RB;
		}
		[CompilerGenerated]
		set
		{
			RB = value;
		}
	}

	internal Severity HiddenNames
	{
		[CompilerGenerated]
		get
		{
			return SB;
		}
		[CompilerGenerated]
		set
		{
			SB = value;
		}
	}

	internal Severity UnusedNames
	{
		[CompilerGenerated]
		get
		{
			return TB;
		}
		[CompilerGenerated]
		set
		{
			TB = value;
		}
	}

	internal Severity ExcessStyles
	{
		[CompilerGenerated]
		get
		{
			return UB;
		}
		[CompilerGenerated]
		set
		{
			UB = value;
		}
	}

	internal Severity NumbersStoredAsText
	{
		[CompilerGenerated]
		get
		{
			return VB;
		}
		[CompilerGenerated]
		set
		{
			VB = value;
		}
	}

	internal Severity DataOutliers
	{
		[CompilerGenerated]
		get
		{
			return WB;
		}
		[CompilerGenerated]
		set
		{
			WB = value;
		}
	}

	internal Severity DataValidationIgnored
	{
		[CompilerGenerated]
		get
		{
			return XB;
		}
		[CompilerGenerated]
		set
		{
			XB = value;
		}
	}

	internal Severity UsedRangeInflation
	{
		[CompilerGenerated]
		get
		{
			return YB;
		}
		[CompilerGenerated]
		set
		{
			YB = value;
		}
	}

	internal Severity CellFillColor
	{
		[CompilerGenerated]
		get
		{
			return ZB;
		}
		[CompilerGenerated]
		set
		{
			ZB = value;
		}
	}

	internal Severity CellBorderColor
	{
		[CompilerGenerated]
		get
		{
			return AC;
		}
		[CompilerGenerated]
		set
		{
			AC = value;
		}
	}

	internal int MaxFormulaLength
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

	internal int MaxNumberOfPrecedents
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal int MaxNumberOfOperators
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal int MaxNumberOfFunctions
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal int MaxNumberOfGroupings
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	internal int MaxNumberOfIfs
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	internal int MaxNestingLevel
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[CompilerGenerated]
		set
		{
			this.m_G = value;
		}
	}

	internal int MaxNamesCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[CompilerGenerated]
		set
		{
			this.m_H = value;
		}
	}

	internal int MaxStylesCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal long MaxFileSize
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

	internal int MaxFileAgeInMonths
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal static int MaxObsPerTitleTypeSeverity
	{
		[CompilerGenerated]
		get
		{
			return Settings.m_K;
		}
	}

	public Settings()
	{
		this.m_A = VH.A(38064);
		this.B = VH.A(38091);
		this.C = VH.A(38116);
		this.D = VH.A(38143);
		this.E = VH.A(38182);
		this.F = VH.A(38227);
		this.G = VH.A(38266);
		ID_UNNECESSARY_FMLA = VH.A(38293);
		this.I = VH.A(38332);
		this.J = VH.A(38371);
		this.K = VH.A(38402);
		this.L = VH.A(38437);
		this.M = VH.A(38470);
		this.N = VH.A(38503);
		this.O = VH.A(38536);
		this.P = VH.A(38579);
		this.Q = VH.A(38614);
		this.R = VH.A(38637);
		this.S = VH.A(38678);
		this.T = VH.A(38717);
		this.U = VH.A(38752);
		this.V = VH.A(38791);
		this.W = VH.A(38814);
		this.X = VH.A(38857);
		this.Y = VH.A(38880);
		this.Z = VH.A(38901);
		this.AB = VH.A(38936);
		this.BB = VH.A(38973);
		this.CB = VH.A(39006);
		this.DB = VH.A(39029);
		this.EB = VH.A(39070);
		this.FB = VH.A(39119);
		this.GB = VH.A(39146);
		this.HB = VH.A(39179);
		this.IB = VH.A(39204);
		this.JB = VH.A(39237);
		this.KB = VH.A(39272);
		this.LB = VH.A(39313);
		this.MB = VH.A(39328);
		this.NB = VH.A(39357);
		this.OB = VH.A(39384);
		this.PB = VH.A(39413);
		this.QB = VH.A(39456);
		this.RB = VH.A(39503);
		this.SB = VH.A(39526);
		this.TB = VH.A(39549);
		this.UB = VH.A(39572);
		this.VB = VH.A(39603);
		this.WB = VH.A(39628);
		this.XB = VH.A(39667);
		this.YB = VH.A(39692);
		this.ZB = VH.A(39735);
		this.AC = VH.A(39772);
		BC = VH.A(39799);
		XmlNode xmlNode = B();
		if (xmlNode == null)
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
			xmlNode = A();
		}
		FormulaErrors = (Severity)A(xmlNode, this.C);
		EmptyCellReferences = (Severity)A(xmlNode, this.D);
		EmptyCellCommentsNotes = (Severity)A(xmlNode, this.E);
		UnusedNumericInputs = (Severity)A(xmlNode, this.F);
		PartialInputs = (Severity)A(xmlNode, this.G);
		UnnecessaryFormulas = (Severity)A(xmlNode, ID_UNNECESSARY_FMLA);
		FormulaInterruption = (Severity)A(xmlNode, this.I);
		FormulasTooLong = (Severity)A(xmlNode, this.J);
		TooManyPrecedents = (Severity)A(xmlNode, this.K);
		TooManyOperators = (Severity)A(xmlNode, this.L);
		TooManyFunctions = (Severity)A(xmlNode, this.M);
		TooManyGroupings = (Severity)A(xmlNode, this.N);
		ConditionalComplexity = (Severity)A(xmlNode, this.O);
		DuplicateFormulas = (Severity)A(xmlNode, this.P);
		DeepNesting = (Severity)A(xmlNode, this.Q);
		ExtraneousSheetNames = (Severity)A(xmlNode, this.R);
		LegacyArrayFormulas = (Severity)A(xmlNode, this.S);
		VolatileFunctions = (Severity)A(xmlNode, this.T);
		DeprecatedFunctions = (Severity)A(xmlNode, this.U);
		ApproximateMatch = (Severity)A(xmlNode, this.V);
		NumericIndexReference = (Severity)A(xmlNode, this.W);
		DoubleMinus = (Severity)A(xmlNode, this.X);
		DoubleSums = (Severity)A(xmlNode, this.Y);
		OmittedReferences = (Severity)A(xmlNode, this.Z);
		CircularReferences = (Severity)A(xmlNode, this.AB);
		InputsNotColored = (Severity)A(xmlNode, this.BB);
		MergedCells = (Severity)A(xmlNode, this.CB);
		ExcessConditionalFormatting = (Severity)A(xmlNode, this.DB);
		TripleSemicolonNumFormat = (Severity)A(xmlNode, this.EB);
		SensitiveData = (Severity)A(xmlNode, this.FB);
		CommentsAndNotes = (Severity)A(xmlNode, this.GB);
		HiddenSheets = (Severity)A(xmlNode, this.HB);
		VeryHiddenSheets = (Severity)A(xmlNode, this.IB);
		HiddenRowsColumns = (Severity)A(xmlNode, this.JB);
		CollapsedRowsColumns = (Severity)A(xmlNode, this.KB);
		OldFile = (Severity)A(xmlNode, this.LB);
		LegacyFileType = (Severity)A(xmlNode, this.MB);
		LargeFileSize = (Severity)A(xmlNode, this.NB);
		CalculationModeManual = (Severity)A(xmlNode, this.OB);
		CellFillColor = (Severity)A(xmlNode, this.AC);
		CellBorderColor = (Severity)A(xmlNode, BC);
		CoverMissing = (Severity)A(xmlNode, this.B);
		ShapesOverNonEmptyCells = (Severity)A(xmlNode, this.QB);
		NamesWithExternalReferences = (Severity)A(xmlNode, this.UB);
		HiddenNames = (Severity)A(xmlNode, this.SB);
		UnusedNames = (Severity)A(xmlNode, this.TB);
		ExcessNames = (Severity)A(xmlNode, this.RB);
		ExcessStyles = (Severity)A(xmlNode, this.VB);
		NumbersStoredAsText = (Severity)A(xmlNode, this.WB);
		DataOutliers = (Severity)A(xmlNode, this.XB);
		DataValidationIgnored = (Severity)A(xmlNode, this.YB);
		UsedRangeInflation = (Severity)A(xmlNode, this.ZB);
		DisplayDrawingObjects = (Severity)A(xmlNode, this.PB);
		MaxFormulaLength = B(xmlNode, this.J);
		MaxNumberOfPrecedents = B(xmlNode, this.K);
		MaxNumberOfOperators = B(xmlNode, this.L);
		MaxNumberOfFunctions = B(xmlNode, this.M);
		MaxNumberOfGroupings = B(xmlNode, this.N);
		MaxNestingLevel = B(xmlNode, this.Q);
		MaxNumberOfIfs = B(xmlNode, this.O);
		MaxNamesCount = B(xmlNode, this.RB);
		MaxStylesCount = B(xmlNode, this.VB);
		MaxFileSize = B(xmlNode, this.NB);
		MaxFileAgeInMonths = B(xmlNode, this.LB);
		xmlNode = null;
	}

	internal void A(int A)
	{
		this.A(this.J, A);
		MaxFormulaLength = A;
	}

	internal void B(int A)
	{
		this.A(this.K, A);
		MaxNumberOfPrecedents = A;
	}

	internal void C(int A)
	{
		this.A(this.L, A);
		MaxNumberOfOperators = A;
	}

	internal void D(int A)
	{
		this.A(this.M, A);
		MaxNumberOfFunctions = A;
	}

	internal void E(int A)
	{
		this.A(this.N, A);
		MaxNumberOfGroupings = A;
	}

	internal void F(int A)
	{
		this.A(this.O, A);
		MaxNumberOfIfs = A;
	}

	internal void G(int A)
	{
		this.A(this.Q, A);
		MaxNestingLevel = A;
	}

	internal void H(int A)
	{
		this.A(this.RB, A);
		MaxNamesCount = A;
	}

	internal void I(int A)
	{
		this.A(this.VB, A);
		MaxStylesCount = A;
	}

	internal void J(int A)
	{
		checked
		{
			this.A(this.NB, A * 1000);
			MaxFileSize = A * 1000;
		}
	}

	internal void K(int A)
	{
		this.A(this.LB, A);
		MaxFileAgeInMonths = A;
	}

	internal void A(string A, Severity B)
	{
		XmlDocument xml = Manage.GetXml(false);
		XmlAttribute xmlAttribute = this.A(this.A(xml), A).Attributes[VH.A(38014)];
		int num = (int)B;
		xmlAttribute.Value = num.ToString();
		Manage.Save(xml, true);
		xml = null;
	}

	private void A(string A, int B)
	{
		XmlDocument xml = Manage.GetXml(false);
		this.A(this.A(xml), A).Attributes[VH.A(38031)].Value = B.ToString();
		Manage.Save(xml, true);
		xml = null;
	}

	private XmlNode A()
	{
		XmlDocument xml = Manage.GetXml(false);
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(global::A.J.DefaultAuditSettings);
		XmlDocument xmlDocument2 = xml;
		xmlDocument2.DocumentElement.InsertAfter(xmlDocument2.ImportNode(xmlDocument.DocumentElement, deep: true), xmlDocument2.DocumentElement.LastChild);
		xmlDocument2 = null;
		Manage.Save(xml, true);
		xmlDocument = null;
		return B();
	}

	private XmlNode A(string A)
	{
		XmlDocument xml = Manage.GetXml(false);
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(global::A.J.DefaultAuditSettings);
		XmlDocument xmlDocument2 = xml;
		xmlDocument2.DocumentElement.SelectSingleNode(this.m_A).InsertAfter(xmlDocument2.ImportNode(xmlDocument.DocumentElement.SelectSingleNode(VH.A(38038) + A + VH.A(38059)), deep: true), xmlDocument2.DocumentElement.SelectSingleNode(this.m_A).LastChild);
		xmlDocument2 = null;
		Manage.Save(xml, true);
		xmlDocument = null;
		return this.A(this.A(xml), A);
	}

	private XmlNode B()
	{
		return A(KH.A.SettingsXml);
	}

	private XmlNode A(XmlDocument A)
	{
		return A.DocumentElement.SelectSingleNode(this.m_A);
	}

	private XmlNode A(XmlNode A, string B)
	{
		return A.SelectSingleNode(VH.A(38038) + B + VH.A(38059));
	}

	private int A(XmlNode A, string B)
	{
		XmlNode xmlNode = this.A(A, B);
		if (xmlNode == null)
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
			xmlNode = this.A(B);
		}
		return Conversions.ToInteger(xmlNode.Attributes[VH.A(38014)].Value);
	}

	private int B(XmlNode A, string B)
	{
		return Conversions.ToInteger(this.A(A, B).Attributes[VH.A(38031)].Value);
	}
}
