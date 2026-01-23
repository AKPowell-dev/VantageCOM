using System;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using MacabacusMacros.Config.Settings;
using MacabacusMacros.Proofing;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Settings : Settings
{
	private readonly string m_A;

	internal readonly string B;

	internal readonly string C;

	internal readonly string D;

	internal readonly string E;

	internal readonly string F;

	internal readonly string G;

	internal readonly string H;

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
	private Severity m_L;

	[CompilerGenerated]
	private Severity m_M;

	[CompilerGenerated]
	private Severity m_N;

	[CompilerGenerated]
	private Severity m_O;

	[CompilerGenerated]
	private Severity m_P;

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
	private int m_A;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private int m_C;

	internal Severity BulletPunctuation
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_A = value;
		}
	}

	internal Severity BulletSize
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_B = value;
		}
	}

	internal Severity BulletFontFamily
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_C = value;
		}
	}

	internal Severity BulletIndentation
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_D = value;
		}
	}

	internal Severity SlideTitles
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_E = value;
		}
	}

	internal Severity SlideNumbers
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_F;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_F = value;
		}
	}

	internal Severity HiddenSlides
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_G;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_G = value;
		}
	}

	internal Severity HiddenShapes
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_H;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_H = value;
		}
	}

	internal Severity Hyperlinks
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_I;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_I = value;
		}
	}

	internal Severity MisalignedShapes
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_J;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_J = value;
		}
	}

	internal Severity RotatedShapes
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_K;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_K = value;
		}
	}

	internal Severity OverlappingText
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_L;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_L = value;
		}
	}

	internal Severity LineSpacing
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_M;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_M = value;
		}
	}

	internal Severity ShrinkTextOnOverflow
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_N;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_N = value;
		}
	}

	internal Severity CheckPlaceholderLayoutMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_O;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_O = value;
		}
	}

	internal Severity CheckPlaceholderFillMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return this.m_P;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			this.m_P = value;
		}
	}

	internal Severity CheckPlaceholderFontColorMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return Q;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			Q = value;
		}
	}

	internal Severity CheckPlaceholderFontStyleMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return R;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			R = value;
		}
	}

	internal Severity CheckPlaceholderBulletMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return S;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			S = value;
		}
	}

	internal Severity CheckPlaceholderIndentMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return T;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			T = value;
		}
	}

	internal Severity CheckPlaceholderMarginMismatch
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return U;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			U = value;
		}
	}

	internal Severity Footnotes
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return V;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			V = value;
		}
	}

	internal Severity TemplateRules
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return W;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			W = value;
		}
	}

	internal Severity MultipleSlideMasters
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return X;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			X = value;
		}
	}

	internal Severity CheckAgendaUpdated
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return Y;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			Y = value;
		}
	}

	internal Severity AirplaneMode
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return Z;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			Z = value;
		}
	}

	internal Severity MasterShapePosition
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return AB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			AB = value;
		}
	}

	internal Severity ShapeEffects
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return BB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			BB = value;
		}
	}

	internal Severity TextEffects
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return CB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			CB = value;
		}
	}

	internal Severity Animation
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return DB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			DB = value;
		}
	}

	internal Severity SlideCount
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return EB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			EB = value;
		}
	}

	internal Severity SlideWordCount
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return FB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			FB = value;
		}
	}

	internal Severity BulletWordCount
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return GB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			GB = value;
		}
	}

	internal Severity MinMaxFontSize
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return HB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			HB = value;
		}
	}

	internal Severity FractionalFontSize
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return IB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			IB = value;
		}
	}

	internal Severity IllegalFonts
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return JB;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			JB = value;
		}
	}

	internal int MaxSlides
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

	internal int MaxSlideWords
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

	internal int MaxBulletWords
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

	public Settings()
	{
		this.m_A = AH.A(53503);
		this.B = AH.A(53530);
		this.C = AH.A(53587);
		this.D = AH.A(53630);
		this.E = AH.A(53685);
		this.F = AH.A(53742);
		this.G = AH.A(53815);
		this.H = AH.A(53884);
		this.I = AH.A(53963);
		this.J = AH.A(54042);
		this.K = AH.A(54115);
		this.L = AH.A(54188);
		this.M = AH.A(54261);
		this.N = AH.A(54306);
		this.O = AH.A(54353);
		this.P = AH.A(54400);
		this.Q = AH.A(54447);
		this.R = AH.A(54490);
		this.S = AH.A(54545);
		this.T = AH.A(54594);
		this.U = AH.A(54647);
		this.V = AH.A(54692);
		this.W = AH.A(54755);
		this.X = AH.A(54796);
		this.Y = AH.A(54851);
		this.Z = AH.A(54906);
		this.AB = AH.A(54955);
		this.BB = AH.A(55018);
		this.CB = AH.A(55067);
		this.DB = AH.A(55114);
		this.EB = AH.A(55175);
		this.FB = AH.A(55222);
		this.GB = AH.A(55267);
		this.HB = AH.A(55308);
		this.IB = AH.A(55351);
		this.JB = AH.A(55402);
		KB = AH.A(55455);
		LB = AH.A(55506);
		MB = AH.A(55565);
		NB = AH.A(55612);
		OB = AH.A(55631);
		PB = AH.A(55658);
		XmlNode a = A();
		try
		{
			BulletPunctuation = (Severity)A(a, this.B);
			BulletSize = (Severity)A(a, this.C);
			BulletFontFamily = (Severity)A(a, this.D);
			BulletIndentation = (Severity)A(a, this.E);
			SlideTitles = (Severity)A(a, this.M);
			SlideNumbers = (Severity)A(a, this.N);
			HiddenSlides = (Severity)A(a, this.O);
			HiddenShapes = (Severity)A(a, this.P);
			Hyperlinks = (Severity)A(a, this.Q);
			try
			{
				MisalignedShapes = (Severity)A(a, this.R);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				MisalignedShapes = (Severity)0;
				ProjectData.ClearProjectError();
			}
			try
			{
				RotatedShapes = (Severity)A(a, this.S);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				RotatedShapes = (Severity)0;
				ProjectData.ClearProjectError();
			}
			OverlappingText = (Severity)A(a, this.T);
			LineSpacing = (Severity)A(a, this.U);
			ShrinkTextOnOverflow = (Severity)A(a, this.V);
			CheckPlaceholderLayoutMismatch = (Severity)A(a, this.F);
			CheckPlaceholderFillMismatch = (Severity)A(a, this.G);
			CheckPlaceholderFontColorMismatch = (Severity)A(a, this.H);
			CheckPlaceholderFontStyleMismatch = (Severity)A(a, this.I);
			CheckPlaceholderBulletMismatch = (Severity)A(a, this.J);
			CheckPlaceholderIndentMismatch = (Severity)A(a, this.K);
			CheckPlaceholderMarginMismatch = (Severity)A(a, this.L);
			Footnotes = (Severity)A(a, this.W);
			((Settings)this).TableCellMargins = (Severity)A(a, this.X);
			((Settings)this).ShapeOutOfBounds = (Severity)A(a, this.Y);
			TemplateRules = (Severity)A(a, this.Z);
			MultipleSlideMasters = (Severity)A(a, this.AB);
			CheckAgendaUpdated = (Severity)A(a, this.BB);
			AirplaneMode = (Severity)A(a, this.CB);
			MasterShapePosition = (Severity)A(a, this.DB);
			ShapeEffects = (Severity)A(a, this.EB);
			TextEffects = (Severity)A(a, this.FB);
			Animation = (Severity)A(a, this.GB);
			SlideCount = (Severity)A(a, this.HB);
			SlideWordCount = (Severity)A(a, this.IB);
			try
			{
				BulletWordCount = (Severity)A(a, this.JB);
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				BulletWordCount = (Severity)0;
				ProjectData.ClearProjectError();
			}
			try
			{
				MinMaxFontSize = (Severity)A(a, KB);
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				MinMaxFontSize = (Severity)0;
				ProjectData.ClearProjectError();
			}
			try
			{
				FractionalFontSize = (Severity)A(a, LB);
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				FractionalFontSize = (Severity)0;
				ProjectData.ClearProjectError();
			}
			try
			{
				IllegalFonts = (Severity)A(a, MB);
			}
			catch (Exception ex11)
			{
				ProjectData.SetProjectError(ex11);
				Exception ex12 = ex11;
				IllegalFonts = (Severity)0;
				ProjectData.ClearProjectError();
			}
			XmlDocument xmlDoc = base.xmlDoc;
			MaxSlides = Conversions.ToInteger(xmlDoc.SelectSingleNode(Settings.CONVENTIONS_NODE + NB).InnerText);
			MaxSlideWords = Conversions.ToInteger(xmlDoc.SelectSingleNode(Settings.CONVENTIONS_NODE + OB).InnerText);
			try
			{
				MaxBulletWords = Conversions.ToInteger(xmlDoc.SelectSingleNode(Settings.CONVENTIONS_NODE + PB).InnerText);
			}
			catch (Exception ex13)
			{
				ProjectData.SetProjectError(ex13);
				Exception ex14 = ex13;
				MaxBulletWords = 40;
				ProjectData.ClearProjectError();
			}
			xmlDoc = null;
		}
		catch (Exception ex15)
		{
			ProjectData.SetProjectError(ex15);
			Exception ex16 = ex15;
			((Settings)this).LoadError(ex16);
			ProjectData.ClearProjectError();
		}
		base.xmlDoc = null;
	}

	internal void A(int A)
	{
		this.A(NB, A);
		MaxSlides = A;
	}

	internal void B(int A)
	{
		this.A(OB, A);
		MaxSlideWords = A;
	}

	internal void C(int A)
	{
		this.A(PB, A);
		MaxBulletWords = A;
	}

	internal void D(int A)
	{
		this.A(base.ID_CONV_MAX_FONT_SIZES, A);
		((Settings)this).MaxFontSizes = A;
	}

	internal void E(int A)
	{
		this.A(base.ID_CONV_SENTENCE_SP, A);
		((Settings)this).SentenceSpacingConvention = (SpacesBetweenSentences)A;
	}

	internal void F(int A)
	{
		this.A(base.ID_CONV_COLON_SP, A);
		((Settings)this).ColonSpacingConvention = (SpacesAfterColon)A;
	}

	internal void G(int A)
	{
		this.A(base.ID_CONV_SLASH_SP, A);
		((Settings)this).SlashSpacingConvention = (SlashSpacing)A;
	}

	internal void H(int A)
	{
		this.A(base.ID_CONV_DASH_SP, A);
		((Settings)this).DashSpacingConvention = (DashSpacing)A;
	}

	internal void I(int A)
	{
		this.A(base.ID_CONV_UNITS_SP, A);
		((Settings)this).UnitsSpacingConvention = (UnitsSpacing)A;
	}

	internal void J(int A)
	{
		this.A(base.ID_CONV_PROOF_LANG_ID, A);
		((Settings)this).DefaultLanguageId = A;
	}

	internal void K(int A)
	{
		this.A(base.ID_CONV_QUOTES_STYLE, A);
		((Settings)this).QuotesStyleConvention = (DoubleSingleQuotesStyle)A;
	}

	internal void L(int A)
	{
		this.A(base.ID_CONV_SLD_TITLE_CASE, A);
		((Settings)this).SlideTitleCaseConvention = (SlideTitleCase)A;
	}

	internal void M(int A)
	{
		this.A(base.ID_CONV_TRAIL_BULLET_PUNCT, A);
		((Settings)this).BulletPunctuationConvention = (TrailingBulletPunctuation)A;
	}

	internal void N(int A)
	{
		this.A(base.ID_CONV_TRAIL_IE_EG_COMMA, A);
		((Settings)this).IeEgCommaConvention = (IeEgTrailingComma)A;
	}

	internal void O(int A)
	{
		this.A(base.ID_CONV_SPELL_CANCELED, A);
		((Settings)this).CanceledSpellingConvention = (CanceledSpelling)A;
	}

	internal void P(int A)
	{
		this.A(base.ID_CONV_SPELL_ADVISER, A);
		((Settings)this).AdviserSpellingConvention = (AdviserSpelling)A;
	}

	internal void A(string A)
	{
		this.A(base.ID_CONV_MILLIONS_ABBREV, A);
		((Settings)this).MillionsAbbreviationConvention = A;
	}

	internal void B(string A)
	{
		this.A(base.ID_CONV_BILLIONS_ABBREV, A);
		((Settings)this).BillionsAbbreviationConvention = A;
	}

	internal void A(string A, Severity B)
	{
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Expected I4, but got Unknown
		XmlDocument xml = Manage.GetXml(false);
		((Settings)this).GetRule(this.A(xml), A).Attributes[AH.A(53486)].Value = ((int)B).ToString();
		Manage.Save(xml, true);
		xml = null;
	}

	private void A(string A, int B)
	{
		XmlDocument xml = Manage.GetXml(false);
		xml.SelectSingleNode(Settings.CONVENTIONS_NODE + A).InnerText = B.ToString();
		Manage.Save(xml, true);
	}

	private void A(string A, string B)
	{
		XmlDocument xml = Manage.GetXml(false);
		xml.SelectSingleNode(Settings.CONVENTIONS_NODE + A).InnerText = B;
		Manage.Save(xml, true);
	}

	private XmlNode A()
	{
		return A(KG.A.SettingsXml);
	}

	private XmlNode A(XmlDocument A)
	{
		return A.DocumentElement.SelectSingleNode(this.m_A);
	}

	private int A(XmlNode A, string B)
	{
		return Conversions.ToInteger(((Settings)this).GetRule(A, B).Attributes[AH.A(53486)].Value);
	}
}
