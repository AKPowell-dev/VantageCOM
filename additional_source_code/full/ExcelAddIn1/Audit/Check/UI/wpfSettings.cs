using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using System.Windows.Media;
using A;
using Foo.Controls;
using MacabacusMacros;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.UI;

[DesignerGenerated]
public sealed class wpfSettings : UserControl, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private Settings m_A;

	[CompilerGenerated]
	private wpfAudit m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFormulaErrors")]
	private Button m_B;

	[AccessedThroughProperty("btnFormulaInterrupt")]
	[CompilerGenerated]
	private Button m_C;

	[AccessedThroughProperty("btnRefEmptyCells")]
	[CompilerGenerated]
	private Button D;

	[AccessedThroughProperty("btnOmittedRefs")]
	[CompilerGenerated]
	private Button E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUnusedInputs")]
	private Button F;

	[AccessedThroughProperty("btnDoubleMinus")]
	[CompilerGenerated]
	private Button G;

	[AccessedThroughProperty("btnRedundantSums")]
	[CompilerGenerated]
	private Button H;

	[AccessedThroughProperty("btnFormulasTooLong")]
	[CompilerGenerated]
	private Button I;

	[AccessedThroughProperty("numMaxFormulaLength")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTooManyPrecedents")]
	private Button J;

	[AccessedThroughProperty("numMaxPrecedents")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTooManyOperators")]
	private Button K;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxOperators")]
	private MacNumericUpDown m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTooManyFunctions")]
	private Button L;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxFunctions")]
	private MacNumericUpDown D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTooManyGroupings")]
	private Button M;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxGroupings")]
	private MacNumericUpDown E;

	[AccessedThroughProperty("btnDeepNesting")]
	[CompilerGenerated]
	private Button N;

	[AccessedThroughProperty("numMaxNestingLevel")]
	[CompilerGenerated]
	private MacNumericUpDown F;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCondComplexity")]
	private Button O;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxIfs")]
	private MacNumericUpDown G;

	[CompilerGenerated]
	[AccessedThroughProperty("btnExtraneousSheetNames")]
	private Button P;

	[AccessedThroughProperty("btnCircularReferences")]
	[CompilerGenerated]
	private Button Q;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPartialInputs")]
	private Button R;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDuplicateFormulas")]
	private Button S;

	[AccessedThroughProperty("btnApproximateMatch")]
	[CompilerGenerated]
	private Button T;

	[CompilerGenerated]
	[AccessedThroughProperty("btnNumericIndexRef")]
	private Button U;

	[AccessedThroughProperty("btnDeprecatedFxns")]
	[CompilerGenerated]
	private Button V;

	[AccessedThroughProperty("btnLegacyArrayFormulas")]
	[CompilerGenerated]
	private Button W;

	[AccessedThroughProperty("btnUnnecessaryFormulas")]
	[CompilerGenerated]
	private Button X;

	[CompilerGenerated]
	[AccessedThroughProperty("btnVolatileFxns")]
	private Button Y;

	[AccessedThroughProperty("btnConditionalFormats")]
	[CompilerGenerated]
	private Button Z;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUsedRange")]
	private Button AB;

	[AccessedThroughProperty("btnExcessNames")]
	[CompilerGenerated]
	private Button BB;

	[AccessedThroughProperty("numMaxNamesCount")]
	[CompilerGenerated]
	private MacNumericUpDown H;

	[AccessedThroughProperty("btnExcessStyles")]
	[CompilerGenerated]
	private Button CB;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxStylesCount")]
	private MacNumericUpDown I;

	[AccessedThroughProperty("btnUnusedNames")]
	[CompilerGenerated]
	private Button DB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDataOutliers")]
	private Button EB;

	[AccessedThroughProperty("btnDataValidation")]
	[CompilerGenerated]
	private Button FB;

	[AccessedThroughProperty("btnNumsAsText")]
	[CompilerGenerated]
	private Button GB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCalcMode")]
	private Button HB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHiddenSheets")]
	private Button IB;

	[AccessedThroughProperty("btnVeryHiddenSheets")]
	[CompilerGenerated]
	private Button JB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHiddenRowsCols")]
	private Button KB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCollapsedRowsCols")]
	private Button LB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTripleSemicolon")]
	private Button MB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnShapesOverCells")]
	private Button NB;

	[AccessedThroughProperty("btnDisplayDwgObjects")]
	[CompilerGenerated]
	private Button OB;

	[AccessedThroughProperty("btnHiddenNames")]
	[CompilerGenerated]
	private Button PB;

	[AccessedThroughProperty("btnInputsNotColored")]
	[CompilerGenerated]
	private Button QB;

	[AccessedThroughProperty("btnMergedCells")]
	[CompilerGenerated]
	private Button RB;

	[AccessedThroughProperty("btnCoverMissing")]
	[CompilerGenerated]
	private Button SB;

	[AccessedThroughProperty("btnCellFillColor")]
	[CompilerGenerated]
	private Button TB;

	[AccessedThroughProperty("btnCellBorderColor")]
	[CompilerGenerated]
	private Button UB;

	[AccessedThroughProperty("btnSensitiveData")]
	[CompilerGenerated]
	private Button VB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCommentsNotes")]
	private Button WB;

	[AccessedThroughProperty("btnEmptyCellCmtNote")]
	[CompilerGenerated]
	private Button XB;

	[AccessedThroughProperty("btnLargeFileSize")]
	[CompilerGenerated]
	private Button YB;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxFileSize")]
	private MacNumericUpDown J;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOldFile")]
	private Button ZB;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxFileAge")]
	private MacNumericUpDown K;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLegacyFileType")]
	private Button AC;

	[AccessedThroughProperty("btnNamesExtRef")]
	[CompilerGenerated]
	private Button BC;

	private bool m_A;

	private Settings Prefs
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

	private wpfAudit ParentView
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

	internal virtual Button btnClose
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
			RoutedEventHandler value2 = CloseView;
			Button button = this.m_A;
			if (button != null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnFormulaErrors
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

	internal virtual Button btnFormulaInterrupt
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

	internal virtual Button btnRefEmptyCells
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	internal virtual Button btnOmittedRefs
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	internal virtual Button btnUnusedInputs
	{
		[CompilerGenerated]
		get
		{
			return this.F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.F = value;
		}
	}

	internal virtual Button btnDoubleMinus
	{
		[CompilerGenerated]
		get
		{
			return this.G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.G = value;
		}
	}

	internal virtual Button btnRedundantSums
	{
		[CompilerGenerated]
		get
		{
			return this.H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.H = value;
		}
	}

	internal virtual Button btnFormulasTooLong
	{
		[CompilerGenerated]
		get
		{
			return this.I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.I = value;
		}
	}

	internal virtual MacNumericUpDown numMaxFormulaLength
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

	internal virtual Button btnTooManyPrecedents
	{
		[CompilerGenerated]
		get
		{
			return this.J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.J = value;
		}
	}

	internal virtual MacNumericUpDown numMaxPrecedents
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

	internal virtual Button btnTooManyOperators
	{
		[CompilerGenerated]
		get
		{
			return this.K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.K = value;
		}
	}

	internal virtual MacNumericUpDown numMaxOperators
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

	internal virtual Button btnTooManyFunctions
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

	internal virtual MacNumericUpDown numMaxFunctions
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
			D = value;
		}
	}

	internal virtual Button btnTooManyGroupings
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

	internal virtual MacNumericUpDown numMaxGroupings
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal virtual Button btnDeepNesting
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

	internal virtual MacNumericUpDown numMaxNestingLevel
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal virtual Button btnCondComplexity
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

	internal virtual MacNumericUpDown numMaxIfs
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal virtual Button btnExtraneousSheetNames
	{
		[CompilerGenerated]
		get
		{
			return P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			P = value;
		}
	}

	internal virtual Button btnCircularReferences
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	internal virtual Button btnPartialInputs
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	internal virtual Button btnDuplicateFormulas
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	internal virtual Button btnApproximateMatch
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	internal virtual Button btnNumericIndexRef
	{
		[CompilerGenerated]
		get
		{
			return U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			U = value;
		}
	}

	internal virtual Button btnDeprecatedFxns
	{
		[CompilerGenerated]
		get
		{
			return V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			V = value;
		}
	}

	internal virtual Button btnLegacyArrayFormulas
	{
		[CompilerGenerated]
		get
		{
			return W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			W = value;
		}
	}

	internal virtual Button btnUnnecessaryFormulas
	{
		[CompilerGenerated]
		get
		{
			return X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			X = value;
		}
	}

	internal virtual Button btnVolatileFxns
	{
		[CompilerGenerated]
		get
		{
			return Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Y = value;
		}
	}

	internal virtual Button btnConditionalFormats
	{
		[CompilerGenerated]
		get
		{
			return Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Z = value;
		}
	}

	internal virtual Button btnUsedRange
	{
		[CompilerGenerated]
		get
		{
			return AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			AB = value;
		}
	}

	internal virtual Button btnExcessNames
	{
		[CompilerGenerated]
		get
		{
			return BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			BB = value;
		}
	}

	internal virtual MacNumericUpDown numMaxNamesCount
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	internal virtual Button btnExcessStyles
	{
		[CompilerGenerated]
		get
		{
			return CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			CB = value;
		}
	}

	internal virtual MacNumericUpDown numMaxStylesCount
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

	internal virtual Button btnUnusedNames
	{
		[CompilerGenerated]
		get
		{
			return DB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			DB = value;
		}
	}

	internal virtual Button btnDataOutliers
	{
		[CompilerGenerated]
		get
		{
			return EB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EB = value;
		}
	}

	internal virtual Button btnDataValidation
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

	internal virtual Button btnNumsAsText
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

	internal virtual Button btnCalcMode
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

	internal virtual Button btnHiddenSheets
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

	internal virtual Button btnVeryHiddenSheets
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

	internal virtual Button btnHiddenRowsCols
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

	internal virtual Button btnCollapsedRowsCols
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

	internal virtual Button btnTripleSemicolon
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

	internal virtual Button btnShapesOverCells
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

	internal virtual Button btnDisplayDwgObjects
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

	internal virtual Button btnHiddenNames
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

	internal virtual Button btnInputsNotColored
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

	internal virtual Button btnMergedCells
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

	internal virtual Button btnCoverMissing
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

	internal virtual Button btnCellFillColor
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

	internal virtual Button btnCellBorderColor
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

	internal virtual Button btnSensitiveData
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

	internal virtual Button btnCommentsNotes
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

	internal virtual Button btnEmptyCellCmtNote
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

	internal virtual Button btnLargeFileSize
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

	internal virtual MacNumericUpDown numMaxFileSize
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

	internal virtual Button btnOldFile
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

	internal virtual MacNumericUpDown numMaxFileAge
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

	internal virtual Button btnLegacyFileType
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

	internal virtual Button btnNamesExtRef
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

	public wpfSettings(wpfAudit parent)
	{
		base.Unloaded += ViewUnloaded;
		InitializeComponent();
		ParentView = parent;
		A();
		B();
	}

	private void A()
	{
		Prefs = new Settings();
		Settings prefs = Prefs;
		A(btnFormulaErrors, prefs.C, prefs.FormulaErrors);
		A(btnRefEmptyCells, prefs.D, prefs.EmptyCellReferences);
		A(btnEmptyCellCmtNote, prefs.E, prefs.EmptyCellCommentsNotes);
		A(btnOmittedRefs, prefs.Z, prefs.OmittedReferences);
		A(btnUnusedInputs, prefs.F, prefs.UnusedNumericInputs);
		A(btnCircularReferences, prefs.AB, prefs.CircularReferences);
		A(btnPartialInputs, prefs.G, prefs.PartialInputs);
		A(btnUnnecessaryFormulas, prefs.ID_UNNECESSARY_FMLA, prefs.UnnecessaryFormulas);
		A(btnFormulaInterrupt, prefs.I, prefs.FormulaInterruption);
		A(btnFormulasTooLong, prefs.J, prefs.FormulasTooLong);
		A(btnTooManyPrecedents, prefs.K, prefs.TooManyPrecedents);
		A(btnTooManyOperators, prefs.L, prefs.TooManyOperators);
		A(btnTooManyFunctions, prefs.M, prefs.TooManyFunctions);
		A(btnTooManyGroupings, prefs.N, prefs.TooManyGroupings);
		A(btnCondComplexity, prefs.O, prefs.ConditionalComplexity);
		A(btnDeepNesting, prefs.Q, prefs.DeepNesting);
		A(btnDuplicateFormulas, prefs.P, prefs.DuplicateFormulas);
		A(btnExtraneousSheetNames, prefs.R, prefs.ExtraneousSheetNames);
		A(btnLegacyArrayFormulas, prefs.S, prefs.LegacyArrayFormulas);
		A(btnVolatileFxns, prefs.T, prefs.VolatileFunctions);
		A(btnDeprecatedFxns, prefs.U, prefs.DeprecatedFunctions);
		A(btnApproximateMatch, prefs.V, prefs.ApproximateMatch);
		A(btnNumericIndexRef, prefs.W, prefs.NumericIndexReference);
		A(btnDoubleMinus, prefs.X, prefs.DoubleMinus);
		A(btnRedundantSums, prefs.Y, prefs.DoubleSums);
		A(btnInputsNotColored, prefs.BB, prefs.InputsNotColored);
		A(btnMergedCells, prefs.CB, prefs.MergedCells);
		A(btnConditionalFormats, prefs.DB, prefs.ExcessConditionalFormatting);
		A(btnTripleSemicolon, prefs.EB, prefs.TripleSemicolonNumFormat);
		A(btnSensitiveData, prefs.FB, prefs.SensitiveData);
		A(btnCommentsNotes, prefs.GB, prefs.CommentsAndNotes);
		A(btnHiddenSheets, prefs.HB, prefs.HiddenSheets);
		A(btnVeryHiddenSheets, prefs.IB, prefs.VeryHiddenSheets);
		A(btnHiddenRowsCols, prefs.JB, prefs.HiddenRowsColumns);
		A(btnCollapsedRowsCols, prefs.KB, prefs.CollapsedRowsColumns);
		A(btnOldFile, prefs.LB, prefs.OldFile);
		A(btnLegacyFileType, prefs.MB, prefs.LegacyFileType);
		A(btnLargeFileSize, prefs.NB, prefs.LargeFileSize);
		A(btnCalcMode, prefs.OB, prefs.CalculationModeManual);
		A(btnShapesOverCells, prefs.QB, prefs.ShapesOverNonEmptyCells);
		A(btnNamesExtRef, prefs.UB, prefs.NamesWithExternalReferences);
		A(btnExcessNames, prefs.RB, prefs.ExcessNames);
		A(btnNumsAsText, prefs.WB, prefs.NumbersStoredAsText);
		A(btnDataOutliers, prefs.XB, prefs.DataOutliers);
		A(btnDataValidation, prefs.YB, prefs.DataValidationIgnored);
		A(btnUsedRange, prefs.ZB, prefs.UsedRangeInflation);
		A(btnUnusedNames, prefs.TB, prefs.UnusedNames);
		A(btnDisplayDwgObjects, prefs.PB, prefs.DisplayDrawingObjects);
		A(btnHiddenNames, prefs.SB, prefs.HiddenNames);
		A(btnCoverMissing, prefs.B, prefs.CoverMissing);
		A(btnCellFillColor, prefs.AC, prefs.CellFillColor);
		A(btnCellBorderColor, prefs.BC, prefs.CellBorderColor);
		numMaxFormulaLength.Value = prefs.MaxFormulaLength;
		numMaxPrecedents.Value = prefs.MaxNumberOfPrecedents;
		numMaxOperators.Value = prefs.MaxNumberOfOperators;
		numMaxFunctions.Value = prefs.MaxNumberOfFunctions;
		numMaxGroupings.Value = prefs.MaxNumberOfGroupings;
		numMaxIfs.Value = prefs.MaxNumberOfIfs;
		numMaxNestingLevel.Value = prefs.MaxNestingLevel;
		numMaxNamesCount.Value = prefs.MaxNamesCount;
		numMaxStylesCount.Value = prefs.MaxStylesCount;
		numMaxFileSize.Value = (double)prefs.MaxFileSize / 1000.0;
		numMaxFileAge.Value = prefs.MaxFileAgeInMonths;
		prefs = null;
	}

	private void B()
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Expected O, but got Unknown
		//IL_0088: Unknown result type (might be due to invalid IL or missing references)
		//IL_0092: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Expected O, but got Unknown
		//IL_00cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d9: Expected O, but got Unknown
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Expected O, but got Unknown
		//IL_00ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0109: Expected O, but got Unknown
		numMaxFormulaLength.ValueChanged += new MacRangeBaseValueChangedHandler(MaxFormulaLengthChanged);
		numMaxPrecedents.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNumPrecedentsChanged);
		numMaxOperators.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNumOperatorsChanged);
		numMaxFunctions.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNumFunctionsChanged);
		numMaxGroupings.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNumGroupingsChanged);
		numMaxIfs.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNumIfsChanged);
		numMaxNestingLevel.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNestingLevelChanged);
		numMaxNamesCount.ValueChanged += new MacRangeBaseValueChangedHandler(MaxNamesCountChanged);
		numMaxStylesCount.ValueChanged += new MacRangeBaseValueChangedHandler(MaxStylesCountChanged);
		numMaxFileSize.ValueChanged += new MacRangeBaseValueChangedHandler(MaxFileSizeChanged);
		numMaxFileAge.ValueChanged += new MacRangeBaseValueChangedHandler(MaxFileAgeChanged);
	}

	private void C()
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Expected O, but got Unknown
		//IL_0086: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		//IL_009f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Expected O, but got Unknown
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Expected O, but got Unknown
		//IL_00d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Expected O, but got Unknown
		//IL_00ea: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f4: Expected O, but got Unknown
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_010b: Expected O, but got Unknown
		numMaxFormulaLength.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxFormulaLengthChanged);
		numMaxPrecedents.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNumPrecedentsChanged);
		numMaxOperators.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNumOperatorsChanged);
		numMaxFunctions.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNumFunctionsChanged);
		numMaxGroupings.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNumGroupingsChanged);
		numMaxIfs.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNumIfsChanged);
		numMaxNestingLevel.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNestingLevelChanged);
		numMaxNamesCount.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxNamesCountChanged);
		numMaxStylesCount.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxStylesCountChanged);
		numMaxFileSize.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxFileSizeChanged);
		numMaxFileAge.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxFileAgeChanged);
	}

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		C();
		Prefs = null;
		ParentView = null;
	}

	private void CycleSeverity(object sender, RoutedEventArgs e)
	{
		Button a = (Button)sender;
		(string, Severity) tuple = A(a);
		int item = (int)tuple.Item2;
		int num;
		if ((uint)item <= 2u)
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
			num = checked(item + 1);
		}
		else
		{
			num = 0;
		}
		A(a, tuple.Item1, (Severity)num);
		Prefs.A(tuple.Item1, (Severity)num);
		a = null;
	}

	private void A(Button A, string B, Severity C)
	{
		Color color;
		Color color2;
		switch (C)
		{
		case Severity.Low:
			color = clsColors.SeverityColorBlue();
			color2 = color;
			break;
		case Severity.Medium:
			color = clsColors.SeverityColorYellow();
			color2 = color;
			break;
		case Severity.High:
			color = clsColors.SeverityColorRed();
			color2 = color;
			break;
		default:
			color = clsColors.SeverityColorGray();
			color2 = Colors.White;
			break;
		}
		A.BorderBrush = new SolidColorBrush(color);
		A.Background = new SolidColorBrush(color2);
		this.B(A, B, C);
	}

	private void B(Button A, string B, Severity C)
	{
		A.Tag = (B, C);
	}

	private (string Id, Severity Severity) A(Button A)
	{
		object tag = A.Tag;
		if (tag == null)
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
					return default((string, Severity));
				}
			}
		}
		return ((string, Severity))tag;
	}

	private void MaxFormulaLengthChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.A(checked((int)Math.Round(numMaxFormulaLength.Value.Value)));
	}

	private void MaxNumPrecedentsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.B(checked((int)Math.Round(numMaxPrecedents.Value.Value)));
	}

	private void MaxNumOperatorsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.C(checked((int)Math.Round(numMaxOperators.Value.Value)));
	}

	private void MaxNumFunctionsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.D(checked((int)Math.Round(numMaxFunctions.Value.Value)));
	}

	private void MaxNumGroupingsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.E(checked((int)Math.Round(numMaxGroupings.Value.Value)));
	}

	private void MaxNumIfsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.F(checked((int)Math.Round(numMaxIfs.Value.Value)));
	}

	private void MaxNestingLevelChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.G(checked((int)Math.Round(numMaxNestingLevel.Value.Value)));
	}

	private void MaxNamesCountChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.H(checked((int)Math.Round(numMaxNamesCount.Value.Value)));
	}

	private void MaxStylesCountChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.I(checked((int)Math.Round(numMaxStylesCount.Value.Value)));
	}

	private void MaxFileSizeChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.J(checked((int)Math.Round(numMaxFileSize.Value.Value)));
	}

	private void MaxFileAgeChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Prefs.K(checked((int)Math.Round(numMaxFileAge.Value.Value)));
	}

	private void CloseView(object sender, RoutedEventArgs e)
	{
		ParentView.J();
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
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(37897), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fb: Expected O, but got Unknown
		//IL_011f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0129: Expected O, but got Unknown
		//IL_0157: Unknown result type (might be due to invalid IL or missing references)
		//IL_0161: Expected O, but got Unknown
		//IL_0185: Unknown result type (might be due to invalid IL or missing references)
		//IL_018f: Expected O, but got Unknown
		//IL_01b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01bd: Expected O, but got Unknown
		//IL_01eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f5: Expected O, but got Unknown
		//IL_0223: Unknown result type (might be due to invalid IL or missing references)
		//IL_022d: Expected O, but got Unknown
		//IL_036f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0379: Expected O, but got Unknown
		//IL_03a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b1: Expected O, but got Unknown
		//IL_05f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0603: Expected O, but got Unknown
		//IL_061d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0627: Expected O, but got Unknown
		if (connectionId == 2)
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
					btnClose = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnFormulaErrors = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			btnFormulaInterrupt = (Button)target;
			return;
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnRefEmptyCells = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			btnOmittedRefs = (Button)target;
			return;
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
					btnUnusedInputs = (Button)target;
					return;
				}
			}
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
					btnDoubleMinus = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnRedundantSums = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			btnFormulasTooLong = (Button)target;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numMaxFormulaLength = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			btnTooManyPrecedents = (Button)target;
			return;
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
					numMaxPrecedents = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnTooManyOperators = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					numMaxOperators = (MacNumericUpDown)target;
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
					btnTooManyFunctions = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			numMaxFunctions = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnTooManyGroupings = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			numMaxGroupings = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnDeepNesting = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numMaxNestingLevel = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCondComplexity = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numMaxIfs = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnExtraneousSheetNames = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnCircularReferences = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnPartialInputs = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnDuplicateFormulas = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			btnApproximateMatch = (Button)target;
			return;
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnNumericIndexRef = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			btnDeprecatedFxns = (Button)target;
			return;
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnLegacyArrayFormulas = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			btnUnnecessaryFormulas = (Button)target;
			return;
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnVolatileFxns = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			btnConditionalFormats = (Button)target;
			return;
		}
		if (connectionId == 35)
		{
			btnUsedRange = (Button)target;
			return;
		}
		if (connectionId == 36)
		{
			btnExcessNames = (Button)target;
			return;
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numMaxNamesCount = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnExcessStyles = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numMaxStylesCount = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnUnusedNames = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnDataOutliers = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnDataValidation = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnNumsAsText = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnCalcMode = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnHiddenSheets = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 46)
		{
			btnVeryHiddenSheets = (Button)target;
			return;
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnHiddenRowsCols = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCollapsedRowsCols = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 49)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnTripleSemicolon = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 50)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnShapesOverCells = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 51)
		{
			btnDisplayDwgObjects = (Button)target;
			return;
		}
		if (connectionId == 52)
		{
			btnHiddenNames = (Button)target;
			return;
		}
		if (connectionId == 53)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnInputsNotColored = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 54)
		{
			btnMergedCells = (Button)target;
			return;
		}
		if (connectionId == 55)
		{
			btnCoverMissing = (Button)target;
			return;
		}
		if (connectionId == 56)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnCellFillColor = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCellBorderColor = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 58)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnSensitiveData = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 59)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnCommentsNotes = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 60)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnEmptyCellCmtNote = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 61)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnLargeFileSize = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 62)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numMaxFileSize = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 63)
		{
			btnOldFile = (Button)target;
			return;
		}
		if (connectionId == 64)
		{
			numMaxFileAge = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 65)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnLegacyFileType = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 66)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnNamesExtRef = (Button)target;
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 1)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = ButtonBase.ClickEvent;
			eventSetter.Handler = new RoutedEventHandler(CycleSeverity);
			((Style)target).Setters.Add(eventSetter);
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
