using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml;
using A;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.UI;

[DesignerGenerated]
public sealed class wpfSettings : UserControl, IComponentConnector, IStyleConnector
{
	private List<string> m_A;

	[CompilerGenerated]
	private wpfPane m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnColors")]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSemiTransparent")]
	private Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFillGradients")]
	private Button D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLineSpacing")]
	private Button E;

	[AccessedThroughProperty("btnBulletIndents")]
	[CompilerGenerated]
	private Button F;

	[AccessedThroughProperty("btnBulletSize")]
	[CompilerGenerated]
	private Button G;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBulletFont")]
	private Button H;

	[CompilerGenerated]
	[AccessedThroughProperty("btnStrikethru")]
	private Button I;

	[CompilerGenerated]
	[AccessedThroughProperty("btnShrinkText")]
	private Button J;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFractionalFont")]
	private Button K;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFontCount")]
	private Button L;

	[AccessedThroughProperty("numMaxFontSizes")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("btnMultiFonts")]
	[CompilerGenerated]
	private Button M;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTableCellsMargins")]
	private Button N;

	[AccessedThroughProperty("btnShapeEffects")]
	[CompilerGenerated]
	private Button O;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTextEffects")]
	private Button P;

	[AccessedThroughProperty("btnIllegalFonts")]
	[CompilerGenerated]
	private Button Q;

	[AccessedThroughProperty("btnMinMaxFontSize")]
	[CompilerGenerated]
	private Button R;

	[AccessedThroughProperty("btnPlaceholderFontStyle")]
	[CompilerGenerated]
	private Button S;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderFontColor")]
	private Button T;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderFill")]
	private Button U;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderIndent")]
	private Button V;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderBullet")]
	private Button W;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderMargin")]
	private Button X;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPlaceholderLayout")]
	private Button Y;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTemplateRules")]
	private Button Z;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMultipleMasters")]
	private Button AB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMasterShapePosition")]
	private Button BB;

	[AccessedThroughProperty("btnSlideNumbers")]
	[CompilerGenerated]
	private Button CB;

	[AccessedThroughProperty("btnShapeOutOfBounds")]
	[CompilerGenerated]
	private Button DB;

	[AccessedThroughProperty("btnShapeOutsideMargins")]
	[CompilerGenerated]
	private Button EB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMisalignedShapes")]
	private Button FB;

	[AccessedThroughProperty("btnRotatedShapes")]
	[CompilerGenerated]
	private Button GB;

	[AccessedThroughProperty("btnShapeOverlapsText")]
	[CompilerGenerated]
	private Button HB;

	[AccessedThroughProperty("btnFootnotes")]
	[CompilerGenerated]
	private Button IB;

	[AccessedThroughProperty("btnRepeatedWords")]
	[CompilerGenerated]
	private Button JB;

	[AccessedThroughProperty("btnDummyText")]
	[CompilerGenerated]
	private Button KB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSlideTitles")]
	private Button LB;

	[AccessedThroughProperty("btnChartElements")]
	[CompilerGenerated]
	private Button MB;

	[AccessedThroughProperty("btnImageDistortion")]
	[CompilerGenerated]
	private Button NB;

	[AccessedThroughProperty("btnImageCropping")]
	[CompilerGenerated]
	private Button OB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLinks")]
	private Button PB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHyperlinks")]
	private Button QB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLinkedPictures")]
	private Button RB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAnimation")]
	private Button SB;

	[AccessedThroughProperty("btnInk")]
	[CompilerGenerated]
	private Button TB;

	[AccessedThroughProperty("btnSlideVisibility")]
	[CompilerGenerated]
	private Button UB;

	[AccessedThroughProperty("btnShapeVisibility")]
	[CompilerGenerated]
	private Button VB;

	[AccessedThroughProperty("btnSlideCount")]
	[CompilerGenerated]
	private Button WB;

	[AccessedThroughProperty("numMaxSlides")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnWordCount")]
	private Button XB;

	[AccessedThroughProperty("numMaxSlideWords")]
	[CompilerGenerated]
	private MacNumericUpDown m_C;

	[AccessedThroughProperty("btnBulletWords")]
	[CompilerGenerated]
	private Button YB;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxBulletWords")]
	private MacNumericUpDown D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBulletPunct")]
	private Button ZB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPunctMissing")]
	private Button AC;

	[AccessedThroughProperty("btnPunctSpacingIncorrect")]
	[CompilerGenerated]
	private Button BC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnPunctSpacingInconsistent")]
	private Button CC;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxSlashSpacing")]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxDashSpacing")]
	private ComboBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnSentenceSpacing")]
	private Button DC;

	[AccessedThroughProperty("cbxSentenceSpacing")]
	[CompilerGenerated]
	private ComboBox m_C;

	[AccessedThroughProperty("btnColonSpacing")]
	[CompilerGenerated]
	private Button EC;

	[AccessedThroughProperty("cbxColonSpacing")]
	[CompilerGenerated]
	private ComboBox D;

	[AccessedThroughProperty("btnNumberAbbrev")]
	[CompilerGenerated]
	private Button FC;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxMillions")]
	private ComboBox E;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxBillions")]
	private ComboBox F;

	[AccessedThroughProperty("cbxUnitsSpacing")]
	[CompilerGenerated]
	private ComboBox G;

	[AccessedThroughProperty("btnQuotesStyle")]
	[CompilerGenerated]
	private Button GC;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxQuotes")]
	private ComboBox H;

	[CompilerGenerated]
	[AccessedThroughProperty("btnIeEgComma")]
	private Button HC;

	[AccessedThroughProperty("cbxIeEgComma")]
	[CompilerGenerated]
	private ComboBox I;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUnnecessaryPeriods")]
	private Button IC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHyphenWordsImproper")]
	private Button JC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHyphenWordsInconsistent")]
	private Button KC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCorpDict")]
	private Button LC;

	[AccessedThroughProperty("btnCanceled")]
	[CompilerGenerated]
	private Button MC;

	[AccessedThroughProperty("cbxCanceledSpelling")]
	[CompilerGenerated]
	private ComboBox J;

	[AccessedThroughProperty("btnAdviser")]
	[CompilerGenerated]
	private Button NC;

	[AccessedThroughProperty("cbxAdviserSpelling")]
	[CompilerGenerated]
	private ComboBox K;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMyriadOf")]
	private Button OC;

	[AccessedThroughProperty("btnAsPer")]
	[CompilerGenerated]
	private Button PC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnConfusedWords")]
	private Button QC;

	[AccessedThroughProperty("btnIeEg")]
	[CompilerGenerated]
	private Button RC;

	[CompilerGenerated]
	[AccessedThroughProperty("btnProofLanguage")]
	private Button SC;

	[AccessedThroughProperty("cbxLanguages")]
	[CompilerGenerated]
	private ComboBox L;

	[AccessedThroughProperty("btnPassiveVoice")]
	[CompilerGenerated]
	private Button TC;

	[AccessedThroughProperty("btnCasualWriting")]
	[CompilerGenerated]
	private Button UC;

	private bool m_B;

	private wpfPane ParentView
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

	internal bool SettingsDirty
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnColors
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

	internal virtual Button btnSemiTransparent
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

	internal virtual Button btnFillGradients
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

	internal virtual Button btnLineSpacing
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

	internal virtual Button btnBulletIndents
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

	internal virtual Button btnBulletSize
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

	internal virtual Button btnBulletFont
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

	internal virtual Button btnStrikethru
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

	internal virtual Button btnShrinkText
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

	internal virtual Button btnFractionalFont
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

	internal virtual Button btnFontCount
	{
		[CompilerGenerated]
		get
		{
			return this.L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.L = value;
		}
	}

	internal virtual MacNumericUpDown numMaxFontSizes
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

	internal virtual Button btnMultiFonts
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

	internal virtual Button btnTableCellsMargins
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

	internal virtual Button btnShapeEffects
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

	internal virtual Button btnTextEffects
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

	internal virtual Button btnIllegalFonts
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

	internal virtual Button btnMinMaxFontSize
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

	internal virtual Button btnPlaceholderFontStyle
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

	internal virtual Button btnPlaceholderFontColor
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

	internal virtual Button btnPlaceholderFill
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

	internal virtual Button btnPlaceholderIndent
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

	internal virtual Button btnPlaceholderBullet
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

	internal virtual Button btnPlaceholderMargin
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

	internal virtual Button btnPlaceholderLayout
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

	internal virtual Button btnTemplateRules
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

	internal virtual Button btnMultipleMasters
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

	internal virtual Button btnMasterShapePosition
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

	internal virtual Button btnSlideNumbers
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

	internal virtual Button btnShapeOutOfBounds
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

	internal virtual Button btnShapeOutsideMargins
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

	internal virtual Button btnMisalignedShapes
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

	internal virtual Button btnRotatedShapes
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

	internal virtual Button btnShapeOverlapsText
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

	internal virtual Button btnFootnotes
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

	internal virtual Button btnRepeatedWords
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

	internal virtual Button btnDummyText
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

	internal virtual Button btnSlideTitles
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

	internal virtual Button btnChartElements
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

	internal virtual Button btnImageDistortion
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

	internal virtual Button btnImageCropping
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

	internal virtual Button btnLinks
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

	internal virtual Button btnHyperlinks
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

	internal virtual Button btnLinkedPictures
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

	internal virtual Button btnAnimation
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

	internal virtual Button btnInk
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

	internal virtual Button btnSlideVisibility
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

	internal virtual Button btnShapeVisibility
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

	internal virtual Button btnSlideCount
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

	internal virtual MacNumericUpDown numMaxSlides
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

	internal virtual Button btnWordCount
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

	internal virtual MacNumericUpDown numMaxSlideWords
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

	internal virtual Button btnBulletWords
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

	internal virtual MacNumericUpDown numMaxBulletWords
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

	internal virtual Button btnBulletPunct
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

	internal virtual Button btnPunctMissing
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

	internal virtual Button btnPunctSpacingIncorrect
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

	internal virtual Button btnPunctSpacingInconsistent
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

	internal virtual ComboBox cbxSlashSpacing
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

	internal virtual ComboBox cbxDashSpacing
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

	internal virtual Button btnSentenceSpacing
	{
		[CompilerGenerated]
		get
		{
			return DC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			DC = value;
		}
	}

	internal virtual ComboBox cbxSentenceSpacing
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

	internal virtual Button btnColonSpacing
	{
		[CompilerGenerated]
		get
		{
			return EC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EC = value;
		}
	}

	internal virtual ComboBox cbxColonSpacing
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

	internal virtual Button btnNumberAbbrev
	{
		[CompilerGenerated]
		get
		{
			return FC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			FC = value;
		}
	}

	internal virtual ComboBox cbxMillions
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

	internal virtual ComboBox cbxBillions
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

	internal virtual ComboBox cbxUnitsSpacing
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

	internal virtual Button btnQuotesStyle
	{
		[CompilerGenerated]
		get
		{
			return GC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			GC = value;
		}
	}

	internal virtual ComboBox cbxQuotes
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

	internal virtual Button btnIeEgComma
	{
		[CompilerGenerated]
		get
		{
			return HC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			HC = value;
		}
	}

	internal virtual ComboBox cbxIeEgComma
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

	internal virtual Button btnUnnecessaryPeriods
	{
		[CompilerGenerated]
		get
		{
			return IC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			IC = value;
		}
	}

	internal virtual Button btnHyphenWordsImproper
	{
		[CompilerGenerated]
		get
		{
			return JC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			JC = value;
		}
	}

	internal virtual Button btnHyphenWordsInconsistent
	{
		[CompilerGenerated]
		get
		{
			return KC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			KC = value;
		}
	}

	internal virtual Button btnCorpDict
	{
		[CompilerGenerated]
		get
		{
			return LC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			LC = value;
		}
	}

	internal virtual Button btnCanceled
	{
		[CompilerGenerated]
		get
		{
			return MC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			MC = value;
		}
	}

	internal virtual ComboBox cbxCanceledSpelling
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

	internal virtual Button btnAdviser
	{
		[CompilerGenerated]
		get
		{
			return NC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			NC = value;
		}
	}

	internal virtual ComboBox cbxAdviserSpelling
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

	internal virtual Button btnMyriadOf
	{
		[CompilerGenerated]
		get
		{
			return OC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			OC = value;
		}
	}

	internal virtual Button btnAsPer
	{
		[CompilerGenerated]
		get
		{
			return PC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			PC = value;
		}
	}

	internal virtual Button btnConfusedWords
	{
		[CompilerGenerated]
		get
		{
			return QC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			QC = value;
		}
	}

	internal virtual Button btnIeEg
	{
		[CompilerGenerated]
		get
		{
			return RC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RC = value;
		}
	}

	internal virtual Button btnProofLanguage
	{
		[CompilerGenerated]
		get
		{
			return SC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			SC = value;
		}
	}

	internal virtual ComboBox cbxLanguages
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

	internal virtual Button btnPassiveVoice
	{
		[CompilerGenerated]
		get
		{
			return TC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			TC = value;
		}
	}

	internal virtual Button btnCasualWriting
	{
		[CompilerGenerated]
		get
		{
			return UC;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			UC = value;
		}
	}

	public wpfSettings(wpfPane parent)
	{
		base.Unloaded += ViewUnloaded;
		SettingsDirty = false;
		InitializeComponent();
		ParentView = parent;
		A();
		B();
	}

	private void A()
	{
		//IL_00c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0131: Unknown result type (might be due to invalid IL or missing references)
		//IL_014b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0150: Unknown result type (might be due to invalid IL or missing references)
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_016c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_019b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_01cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0205: Unknown result type (might be due to invalid IL or missing references)
		//IL_021f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0239: Unknown result type (might be due to invalid IL or missing references)
		//IL_023e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0255: Unknown result type (might be due to invalid IL or missing references)
		//IL_026d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0272: Unknown result type (might be due to invalid IL or missing references)
		//IL_0289: Unknown result type (might be due to invalid IL or missing references)
		//IL_028e: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_02db: Unknown result type (might be due to invalid IL or missing references)
		//IL_02e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0313: Unknown result type (might be due to invalid IL or missing references)
		//IL_0318: Unknown result type (might be due to invalid IL or missing references)
		//IL_032d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0347: Unknown result type (might be due to invalid IL or missing references)
		//IL_034c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0363: Unknown result type (might be due to invalid IL or missing references)
		//IL_0368: Unknown result type (might be due to invalid IL or missing references)
		//IL_037f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0384: Unknown result type (might be due to invalid IL or missing references)
		//IL_039b: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_03cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0407: Unknown result type (might be due to invalid IL or missing references)
		//IL_0421: Unknown result type (might be due to invalid IL or missing references)
		//IL_043b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0440: Unknown result type (might be due to invalid IL or missing references)
		//IL_0457: Unknown result type (might be due to invalid IL or missing references)
		//IL_045c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0473: Unknown result type (might be due to invalid IL or missing references)
		//IL_0478: Unknown result type (might be due to invalid IL or missing references)
		//IL_048f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0494: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0500: Unknown result type (might be due to invalid IL or missing references)
		//IL_0515: Unknown result type (might be due to invalid IL or missing references)
		//IL_051a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0531: Unknown result type (might be due to invalid IL or missing references)
		//IL_054b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0550: Unknown result type (might be due to invalid IL or missing references)
		//IL_0567: Unknown result type (might be due to invalid IL or missing references)
		//IL_056c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0583: Unknown result type (might be due to invalid IL or missing references)
		//IL_0588: Unknown result type (might be due to invalid IL or missing references)
		//IL_059f: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_05be: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_05f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0609: Unknown result type (might be due to invalid IL or missing references)
		//IL_060e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0625: Unknown result type (might be due to invalid IL or missing references)
		//IL_062a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0641: Unknown result type (might be due to invalid IL or missing references)
		//IL_0646: Unknown result type (might be due to invalid IL or missing references)
		//IL_065b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0675: Unknown result type (might be due to invalid IL or missing references)
		//IL_067a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0691: Unknown result type (might be due to invalid IL or missing references)
		//IL_0696: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_06df: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_06f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0711: Unknown result type (might be due to invalid IL or missing references)
		//IL_0716: Unknown result type (might be due to invalid IL or missing references)
		//IL_072d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0732: Unknown result type (might be due to invalid IL or missing references)
		//IL_0749: Unknown result type (might be due to invalid IL or missing references)
		//IL_074e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0765: Unknown result type (might be due to invalid IL or missing references)
		//IL_076a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0781: Unknown result type (might be due to invalid IL or missing references)
		//IL_0786: Unknown result type (might be due to invalid IL or missing references)
		//IL_079d: Unknown result type (might be due to invalid IL or missing references)
		//IL_07a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_07b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_07be: Unknown result type (might be due to invalid IL or missing references)
		//IL_07d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_07ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_07f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0809: Unknown result type (might be due to invalid IL or missing references)
		//IL_080e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0825: Unknown result type (might be due to invalid IL or missing references)
		//IL_082a: Unknown result type (might be due to invalid IL or missing references)
		//IL_083a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0844: Expected I4, but got Unknown
		//IL_084b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0855: Expected I4, but got Unknown
		//IL_085e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0863: Unknown result type (might be due to invalid IL or missing references)
		//IL_086a: Expected I4, but got Unknown
		//IL_0873: Unknown result type (might be due to invalid IL or missing references)
		//IL_0878: Unknown result type (might be due to invalid IL or missing references)
		//IL_087f: Expected I4, but got Unknown
		//IL_0888: Unknown result type (might be due to invalid IL or missing references)
		//IL_088d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0894: Expected I4, but got Unknown
		//IL_0922: Unknown result type (might be due to invalid IL or missing references)
		//IL_0927: Unknown result type (might be due to invalid IL or missing references)
		//IL_092e: Expected I4, but got Unknown
		//IL_0937: Unknown result type (might be due to invalid IL or missing references)
		//IL_093c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0943: Expected I4, but got Unknown
		//IL_094a: Unknown result type (might be due to invalid IL or missing references)
		//IL_094f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0956: Expected I4, but got Unknown
		//IL_095d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0962: Unknown result type (might be due to invalid IL or missing references)
		//IL_0969: Expected I4, but got Unknown
		cbxSentenceSpacing.ItemsSource = SettingsPane.GetSentenceSpacingItems();
		cbxColonSpacing.ItemsSource = SettingsPane.GetColonSpacingItems();
		cbxSlashSpacing.ItemsSource = SettingsPane.GetSlashSpacingItems();
		cbxDashSpacing.ItemsSource = SettingsPane.GetDashSpacingItems();
		cbxUnitsSpacing.ItemsSource = SettingsPane.GetUnitsSpacingItems();
		cbxQuotes.ItemsSource = SettingsPane.GetQuotesStyleItems();
		cbxIeEgComma.ItemsSource = SettingsPane.GetIeEgCommaItems();
		cbxCanceledSpelling.ItemsSource = SettingsPane.GetCanceledSpellingItems();
		cbxAdviserSpelling.ItemsSource = SettingsPane.GetAdviserSpellingItems();
		Settings options = Main.Analysis.Options;
		A(btnPunctSpacingInconsistent, ((Settings)options).ID_PUNCT_SP_INCONSIST, ((Settings)options).PunctuationSpacingInconsistent);
		A(btnPunctSpacingIncorrect, ((Settings)options).ID_PUNCT_SP_INCORRECT, ((Settings)options).PunctuationSpacingIncorrect);
		A(btnPunctMissing, ((Settings)options).ID_PUNCT_MISSING, ((Settings)options).PunctuationMissing);
		A(btnRepeatedWords, ((Settings)options).ID_REPEAT_WORDS, ((Settings)options).RepeatedWords);
		A(btnHyphenWordsInconsistent, ((Settings)options).ID_HYPHEN_WORDS_INCONSIST, ((Settings)options).HyphenWordsInconsistent);
		A(btnQuotesStyle, ((Settings)options).ID_QUOTES_STYLE, ((Settings)options).QuotesStyle);
		A(btnBulletPunct, options.B, options.BulletPunctuation);
		A(btnNumberAbbrev, ((Settings)options).ID_NUM_ABBREV, ((Settings)options).MillionsBillionsAbbreviation);
		A(btnProofLanguage, ((Settings)options).ID_PROOF_LANG, ((Settings)options).ProofingLanguage);
		A(btnChartElements, ((Settings)options).ID_CHART_ELEMENTS, ((Settings)options).CheckChartElements);
		A(btnSlideTitles, options.M, options.SlideTitles);
		A(btnSlideNumbers, options.N, options.SlideNumbers);
		A(btnSlideVisibility, options.O, options.HiddenSlides);
		A(btnShapeVisibility, options.P, options.HiddenShapes);
		A(btnFootnotes, options.W, options.Footnotes);
		A(btnDummyText, ((Settings)options).ID_DUMMY_TEXT, ((Settings)options).DummyText);
		A(btnImageDistortion, ((Settings)options).ID_IMG_DISTORTION, ((Settings)options).ImageDistortion);
		A(btnImageCropping, ((Settings)options).ID_IMG_CROP, ((Settings)options).ImageCropping);
		A(btnLinks, ((Settings)options).ID_LINKS, ((Settings)options).CheckLinks);
		A(btnAnimation, options.GB, options.Animation);
		A(btnInk, ((Settings)options).ID_INK, ((Settings)options).Ink);
		A(btnHyperlinks, options.Q, options.Hyperlinks);
		A(btnLinkedPictures, ((Settings)options).ID_LINKED_PICS, ((Settings)options).LinkedPictures);
		A(btnSlideCount, options.HB, options.SlideCount);
		A(btnWordCount, options.IB, options.SlideWordCount);
		A(btnBulletWords, options.JB, options.BulletWordCount);
		A(btnLineSpacing, options.U, options.LineSpacing);
		A(btnBulletIndents, options.E, options.BulletIndentation);
		A(btnBulletFont, options.D, options.BulletFontFamily);
		A(btnBulletSize, options.C, options.BulletSize);
		A(btnColors, ((Settings)options).ID_COLOR_PALETTE, ((Settings)options).ColorPalette);
		A(btnSemiTransparent, ((Settings)options).ID_FILL_TRANS, ((Settings)options).FillTransparency);
		A(btnFillGradients, ((Settings)options).ID_FILL_GRADIENTS, ((Settings)options).FillGradients);
		A(btnShrinkText, options.V, options.ShrinkTextOnOverflow);
		A(btnStrikethru, ((Settings)options).ID_STRIKETHRU, ((Settings)options).StrikethroughFont);
		A(btnFractionalFont, options.LB, options.FractionalFontSize);
		A(btnIllegalFonts, options.MB, options.IllegalFonts);
		A(btnMinMaxFontSize, options.KB, options.MinMaxFontSize);
		A(btnMultiFonts, ((Settings)options).ID_MULT_FONT_FAMS, ((Settings)options).MultipleFontFamilies);
		A(btnFontCount, ((Settings)options).ID_FONT_FAM_SIZE_CNT, ((Settings)options).FontFamilySizeCount);
		A(btnPlaceholderFill, options.G, options.CheckPlaceholderFillMismatch);
		A(btnPlaceholderFontColor, options.H, options.CheckPlaceholderFontColorMismatch);
		A(btnPlaceholderFontStyle, options.I, options.CheckPlaceholderFontStyleMismatch);
		A(btnPlaceholderBullet, options.J, options.CheckPlaceholderBulletMismatch);
		A(btnPlaceholderIndent, options.K, options.CheckPlaceholderIndentMismatch);
		A(btnShapeEffects, options.EB, options.ShapeEffects);
		A(btnTextEffects, options.FB, options.TextEffects);
		A(btnShapeOutOfBounds, options.Y, ((Settings)options).ShapeOutOfBounds);
		A(btnShapeOutsideMargins, ((Settings)options).ID_SHP_OUTSIDE_MGNS, ((Settings)options).ShapeOutsideMargins);
		A(btnMisalignedShapes, options.R, options.MisalignedShapes);
		A(btnRotatedShapes, options.S, options.RotatedShapes);
		A(btnShapeOverlapsText, options.T, options.OverlappingText);
		A(btnMasterShapePosition, options.DB, options.MasterShapePosition);
		A(btnPlaceholderLayout, options.F, options.CheckPlaceholderLayoutMismatch);
		A(btnPlaceholderMargin, options.L, options.CheckPlaceholderMarginMismatch);
		A(btnTableCellsMargins, options.X, ((Settings)options).TableCellMargins);
		A(btnTemplateRules, options.Z, options.TemplateRules);
		A(btnMultipleMasters, options.AB, options.MultipleSlideMasters);
		A(btnSentenceSpacing, ((Settings)options).ID_SENTENCE_SPACING, ((Settings)options).SentenceSpacing);
		A(btnColonSpacing, ((Settings)options).ID_COLON_SPACING, ((Settings)options).ColonSpacing);
		A(btnMyriadOf, ((Settings)options).ID_MYRIAD_OF, ((Settings)options).GrammarMyriadOf);
		A(btnAsPer, ((Settings)options).ID_AS_PER, ((Settings)options).GrammarAsPer);
		A(btnConfusedWords, ((Settings)options).ID_CONFUSED_WORDS, ((Settings)options).ConfusedWords);
		A(btnIeEg, ((Settings)options).ID_IE_EG, ((Settings)options).GrammarIeEg);
		A(btnCanceled, ((Settings)options).ID_SPELL_CANCELED, ((Settings)options).SpellingCanceled);
		A(btnAdviser, ((Settings)options).ID_SPELL_ADVISER, ((Settings)options).SpellingAdviser);
		A(btnHyphenWordsImproper, ((Settings)options).ID_HYPHEN_WORDS_IMPROPER, ((Settings)options).HyphenWordsImproper);
		A(btnIeEgComma, ((Settings)options).ID_IE_EG_COMMA, ((Settings)options).IeEgComma);
		A(btnPassiveVoice, ((Settings)options).ID_PASSIVE_VOICE, ((Settings)options).PassiveVoice);
		A(btnCasualWriting, ((Settings)options).ID_CASUAL_WRITING, ((Settings)options).CasualWriting);
		A(btnUnnecessaryPeriods, ((Settings)options).ID_UNNECESSARY_PERIODS, ((Settings)options).UnnecessaryPeriods);
		cbxSentenceSpacing.SelectedIndex = (int)((Settings)options).SentenceSpacingConvention;
		cbxColonSpacing.SelectedIndex = (int)((Settings)options).ColonSpacingConvention;
		cbxSlashSpacing.SelectedIndex = (int)((Settings)options).SlashSpacingConvention;
		cbxDashSpacing.SelectedIndex = (int)((Settings)options).DashSpacingConvention;
		cbxUnitsSpacing.SelectedIndex = (int)((Settings)options).UnitsSpacingConvention;
		if (((Settings)options).MillionsAbbreviationConvention.Length == 0)
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
			cbxMillions.SelectedIndex = 0;
		}
		else
		{
			cbxMillions.SelectedValue = ((Settings)options).MillionsAbbreviationConvention;
		}
		if (((Settings)options).BillionsAbbreviationConvention.Length == 0)
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
			cbxBillions.SelectedIndex = 0;
		}
		else
		{
			cbxBillions.SelectedValue = ((Settings)options).BillionsAbbreviationConvention;
		}
		cbxQuotes.SelectedIndex = (int)((Settings)options).QuotesStyleConvention;
		cbxIeEgComma.SelectedIndex = (int)((Settings)options).IeEgCommaConvention;
		cbxCanceledSpelling.SelectedIndex = (int)((Settings)options).CanceledSpellingConvention;
		cbxAdviserSpelling.SelectedIndex = (int)((Settings)options).AdviserSpellingConvention;
		numMaxFontSizes.Value = ((Settings)options).MaxFontSizes;
		numMaxSlides.Value = options.MaxSlides;
		numMaxSlideWords.Value = options.MaxSlideWords;
		numMaxBulletWords.Value = options.MaxBulletWords;
		options = null;
		List<string> list = new List<string>();
		this.m_A = new List<string>();
		XmlNodeList languageNodes = SettingsPane.GetLanguageNodes();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = languageNodes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				XmlNode xmlNode = (XmlNode)enumerator.Current;
				list.Add(xmlNode.Attributes[AH.A(58224)].Value);
				this.m_A.Add(xmlNode.Attributes[AH.A(58243)].Value);
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
		cbxLanguages.ItemsSource = list;
		cbxLanguages.SelectedIndex = this.m_A.IndexOf(Conversions.ToString(((Settings)Main.Analysis.Options).DefaultLanguageId));
		languageNodes = null;
		list = null;
		Settings options2 = Main.Analysis.Options;
		if (((Settings)options2).CorporateDictionaryRules != null)
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
			if (((Settings)options2).CorporateDictionaryRules.Count != 0)
			{
				btnMyriadOf.IsEnabled = false;
				btnAsPer.IsEnabled = false;
				btnCanceled.IsEnabled = false;
				cbxCanceledSpelling.Visibility = Visibility.Hidden;
				cbxCanceledSpelling.SelectedIndex = -1;
				btnAdviser.IsEnabled = false;
				cbxAdviserSpelling.Visibility = Visibility.Hidden;
				cbxAdviserSpelling.SelectedIndex = -1;
				goto IL_0b8f;
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
		btnCorpDict.Background = new SolidColorBrush(System.Windows.Media.Colors.White);
		goto IL_0b8f;
		IL_0b8f:
		options2 = null;
	}

	private void B()
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0041: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Expected O, but got Unknown
		numMaxSlides.ValueChanged += new MacRangeBaseValueChangedHandler(MaxSlidesChanged);
		numMaxSlideWords.ValueChanged += new MacRangeBaseValueChangedHandler(MaxSlideWordsChanged);
		numMaxBulletWords.ValueChanged += new MacRangeBaseValueChangedHandler(MaxBulletWordsChanged);
		numMaxFontSizes.ValueChanged += new MacRangeBaseValueChangedHandler(MaxFontSizesChanged);
		cbxLanguages.SelectionChanged += ProofingLanguageIdChanged;
		cbxSentenceSpacing.SelectionChanged += SentenceSpacingChanged;
		cbxColonSpacing.SelectionChanged += ColonSpacingChanged;
		cbxSlashSpacing.SelectionChanged += SlashSpacingChanged;
		cbxDashSpacing.SelectionChanged += DashSpacingChanged;
		cbxUnitsSpacing.SelectionChanged += UnitsSpacingChanged;
		cbxQuotes.SelectionChanged += QuotesStyleChanged;
		cbxIeEgComma.SelectionChanged += TrailingIeEgCommaChanged;
		cbxCanceledSpelling.SelectionChanged += CanceledSpellingChanged;
		cbxAdviserSpelling.SelectionChanged += AdviserSpellingChanged;
		cbxMillions.SelectionChanged += MillionsAbbrevChanged;
		cbxBillions.SelectionChanged += BillionsAbbrevChanged;
	}

	private void C()
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Expected O, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Expected O, but got Unknown
		numMaxSlides.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxSlidesChanged);
		numMaxSlideWords.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxSlideWordsChanged);
		numMaxBulletWords.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxBulletWordsChanged);
		numMaxFontSizes.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxFontSizesChanged);
		cbxLanguages.SelectionChanged -= ProofingLanguageIdChanged;
		cbxSentenceSpacing.SelectionChanged -= SentenceSpacingChanged;
		cbxColonSpacing.SelectionChanged -= ColonSpacingChanged;
		cbxSlashSpacing.SelectionChanged -= SlashSpacingChanged;
		cbxDashSpacing.SelectionChanged -= DashSpacingChanged;
		cbxUnitsSpacing.SelectionChanged -= UnitsSpacingChanged;
		cbxQuotes.SelectionChanged -= QuotesStyleChanged;
		cbxIeEgComma.SelectionChanged -= TrailingIeEgCommaChanged;
		cbxCanceledSpelling.SelectionChanged -= CanceledSpellingChanged;
		cbxAdviserSpelling.SelectionChanged -= AdviserSpellingChanged;
		cbxMillions.SelectionChanged -= MillionsAbbrevChanged;
		cbxBillions.SelectionChanged -= BillionsAbbrevChanged;
	}

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		C();
		ParentView = null;
		this.m_A = null;
	}

	private void CycleSeverity(object sender, RoutedEventArgs e)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected I4, but got Unknown
		Button a = (Button)sender;
		(string, Severity) tuple = A(a);
		int num = (int)tuple.Item2;
		int num2;
		if ((uint)num <= 2u)
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
			num2 = checked(num + 1);
		}
		else
		{
			num2 = 0;
		}
		A(a, tuple.Item1, (Severity)num2);
		Main.Analysis.Options.A(tuple.Item1, (Severity)num2);
		SettingsDirty = true;
		a = null;
	}

	private void A(Button A, string B, Severity C)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected I4, but got Unknown
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		Color color;
		Color color2;
		switch (C - 1)
		{
		case 0:
			color = clsColors.SeverityColorBlue();
			color2 = color;
			break;
		case 1:
			color = clsColors.SeverityColorYellow();
			color2 = color;
			break;
		case 2:
			color = clsColors.SeverityColorRed();
			color2 = color;
			break;
		default:
			color = clsColors.SeverityColorGray();
			color2 = System.Windows.Media.Colors.White;
			break;
		}
		A.BorderBrush = new SolidColorBrush(color);
		A.Background = new SolidColorBrush(color2);
		this.B(A, B, C);
	}

	private void B(Button A, string B, Severity C)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		A.Tag = (B, C);
	}

	private (string Id, Severity Severity) A(Button A)
	{
		object tag = A.Tag;
		if (tag == null)
		{
			return default((string, Severity));
		}
		return ((string, Severity))tag;
	}

	private void MaxSlidesChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Main.Analysis.Options.A(checked((int)Math.Round(numMaxSlides.Value.Value)));
		SettingsDirty = true;
	}

	private void MaxSlideWordsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Main.Analysis.Options.B(checked((int)Math.Round(numMaxSlideWords.Value.Value)));
		SettingsDirty = true;
	}

	private void MaxBulletWordsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Main.Analysis.Options.C(checked((int)Math.Round(numMaxBulletWords.Value.Value)));
		SettingsDirty = true;
	}

	private void MaxFontSizesChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		Main.Analysis.Options.D(checked((int)Math.Round(numMaxFontSizes.Value.Value)));
		SettingsDirty = true;
	}

	private void ProofingLanguageIdChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.J(Conversions.ToInteger(this.m_A[cbxLanguages.SelectedIndex]));
		SettingsDirty = true;
	}

	private void SentenceSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.E(cbxSentenceSpacing.SelectedIndex);
		SettingsDirty = true;
	}

	private void ColonSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.F(cbxColonSpacing.SelectedIndex);
		SettingsDirty = true;
	}

	private void SlashSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.G(cbxSlashSpacing.SelectedIndex);
		SettingsDirty = true;
	}

	private void DashSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.H(cbxDashSpacing.SelectedIndex);
		SettingsDirty = true;
	}

	private void UnitsSpacingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.I(cbxUnitsSpacing.SelectedIndex);
		SettingsDirty = true;
	}

	private void QuotesStyleChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.K(cbxQuotes.SelectedIndex);
		SettingsDirty = true;
	}

	private void TrailingIeEgCommaChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.N(cbxIeEgComma.SelectedIndex);
		SettingsDirty = true;
	}

	private void TrailingBulletPunctChanged(object sender, SelectionChangedEventArgs e)
	{
		throw new NotImplementedException();
	}

	private void SlideTitleCaseChanged(object sender, SelectionChangedEventArgs e)
	{
		throw new NotImplementedException();
	}

	private void CanceledSpellingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.O(cbxCanceledSpelling.SelectedIndex);
		SettingsDirty = true;
	}

	private void AdviserSpellingChanged(object sender, SelectionChangedEventArgs e)
	{
		Main.Analysis.Options.P(cbxAdviserSpelling.SelectedIndex);
		SettingsDirty = true;
	}

	private void MillionsAbbrevChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxMillions.SelectedIndex > -1)
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
			Main.Analysis.Options.A(Conversions.ToString(cbxMillions.SelectedValue));
		}
		else
		{
			Main.Analysis.Options.A("");
		}
		SettingsDirty = true;
	}

	private void BillionsAbbrevChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxBillions.SelectedIndex > -1)
		{
			Main.Analysis.Options.B(Conversions.ToString(cbxBillions.SelectedValue));
		}
		else
		{
			Main.Analysis.Options.B("");
		}
		SettingsDirty = true;
	}

	private void CloseView(object sender, RoutedEventArgs e)
	{
		ParentView.F();
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
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(58248), UriKind.Relative);
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
		//IL_013b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0145: Expected O, but got Unknown
		//IL_04eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f5: Expected O, but got Unknown
		//IL_0519: Unknown result type (might be due to invalid IL or missing references)
		//IL_0523: Expected O, but got Unknown
		//IL_0547: Unknown result type (might be due to invalid IL or missing references)
		//IL_0551: Expected O, but got Unknown
		if (connectionId == 2)
		{
			btnClose = (Button)target;
			return;
		}
		if (connectionId == 3)
		{
			btnColors = (Button)target;
			return;
		}
		if (connectionId == 4)
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
					btnSemiTransparent = (Button)target;
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
					btnFillGradients = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnLineSpacing = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnBulletIndents = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnBulletSize = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnBulletFont = (Button)target;
					return;
				}
			}
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
					btnStrikethru = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnShrinkText = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			btnFractionalFont = (Button)target;
			return;
		}
		if (connectionId == 13)
		{
			btnFontCount = (Button)target;
			return;
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
					numMaxFontSizes = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			btnMultiFonts = (Button)target;
			return;
		}
		if (connectionId == 16)
		{
			btnTableCellsMargins = (Button)target;
			return;
		}
		if (connectionId == 17)
		{
			btnShapeEffects = (Button)target;
			return;
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnTextEffects = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			btnIllegalFonts = (Button)target;
			return;
		}
		if (connectionId == 20)
		{
			btnMinMaxFontSize = (Button)target;
			return;
		}
		if (connectionId == 21)
		{
			btnPlaceholderFontStyle = (Button)target;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnPlaceholderFontColor = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			btnPlaceholderFill = (Button)target;
			return;
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnPlaceholderIndent = (Button)target;
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
					btnPlaceholderBullet = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnPlaceholderMargin = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnPlaceholderLayout = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnTemplateRules = (Button)target;
					return;
				}
			}
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
					btnMultipleMasters = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnMasterShapePosition = (Button)target;
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
					btnSlideNumbers = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnShapeOutOfBounds = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnShapeOutsideMargins = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			btnMisalignedShapes = (Button)target;
			return;
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnRotatedShapes = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 36)
		{
			btnShapeOverlapsText = (Button)target;
			return;
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnFootnotes = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			btnRepeatedWords = (Button)target;
			return;
		}
		if (connectionId == 39)
		{
			btnDummyText = (Button)target;
			return;
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
					btnSlideTitles = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnChartElements = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnImageDistortion = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			btnImageCropping = (Button)target;
			return;
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnLinks = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnHyperlinks = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 46)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnLinkedPictures = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnAnimation = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnInk = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 49)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnSlideVisibility = (Button)target;
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
					btnShapeVisibility = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 51)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnSlideCount = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 52)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numMaxSlides = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 53)
		{
			btnWordCount = (Button)target;
			return;
		}
		if (connectionId == 54)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numMaxSlideWords = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 55)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnBulletWords = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 56)
		{
			numMaxBulletWords = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnBulletPunct = (Button)target;
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
					btnPunctMissing = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 59)
		{
			btnPunctSpacingIncorrect = (Button)target;
			return;
		}
		if (connectionId == 60)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnPunctSpacingInconsistent = (Button)target;
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
					cbxSlashSpacing = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 62)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxDashSpacing = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 63)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnSentenceSpacing = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 64)
		{
			cbxSentenceSpacing = (ComboBox)target;
			return;
		}
		if (connectionId == 65)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnColonSpacing = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 66)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxColonSpacing = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 67)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnNumberAbbrev = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 68)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					cbxMillions = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 69)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxBillions = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 70)
		{
			cbxUnitsSpacing = (ComboBox)target;
			return;
		}
		if (connectionId == 71)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnQuotesStyle = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 72)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					cbxQuotes = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 73)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnIeEgComma = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 74)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					cbxIeEgComma = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 75)
		{
			btnUnnecessaryPeriods = (Button)target;
			return;
		}
		if (connectionId == 76)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnHyphenWordsImproper = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 77)
		{
			btnHyphenWordsInconsistent = (Button)target;
			return;
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
					btnCorpDict = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 79)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnCanceled = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 80)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxCanceledSpelling = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 81)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnAdviser = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 82)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxAdviserSpelling = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 83)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnMyriadOf = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 84)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnAsPer = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 85)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnConfusedWords = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 86)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnIeEg = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 87)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnProofLanguage = (Button)target;
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
					cbxLanguages = (ComboBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 89:
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				btnPassiveVoice = (Button)target;
				return;
			}
		case 90:
			btnCasualWriting = (Button)target;
			break;
		default:
			this.m_B = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 1)
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
