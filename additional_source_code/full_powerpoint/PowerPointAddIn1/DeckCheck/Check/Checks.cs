using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Checks : Checks
{
	[CompilerGenerated]
	private ShapeOutsideMargins A;

	[CompilerGenerated]
	private ShapeOutOfBounds A;

	[CompilerGenerated]
	private MisalignedShapes A;

	[CompilerGenerated]
	private RotatedShapes A;

	[CompilerGenerated]
	private OverlappingText A;

	[CompilerGenerated]
	private MasterShapePosition A;

	[CompilerGenerated]
	private PlaceholderFontStyleMismatch A;

	[CompilerGenerated]
	private PlaceholderFontColorMismatch A;

	[CompilerGenerated]
	private PlaceholderFillMismatch A;

	[CompilerGenerated]
	private PlaceholderMarginsMismatch A;

	[CompilerGenerated]
	private PlaceholderBulletMismatch A;

	[CompilerGenerated]
	private PlaceholderIndentMismatch A;

	[CompilerGenerated]
	private SlideTitleCapitalization A;

	[CompilerGenerated]
	private SlideNumbers A;

	[CompilerGenerated]
	private HiddenSlides A;

	[CompilerGenerated]
	private ProofingLanguage A;

	[CompilerGenerated]
	private LineSpacing A;

	[CompilerGenerated]
	private BulletPunctuation A;

	[CompilerGenerated]
	private BulletIndentation A;

	[CompilerGenerated]
	private BulletSize A;

	[CompilerGenerated]
	private BulletFontFamily A;

	[CompilerGenerated]
	private MultipleFontFamilies A;

	[CompilerGenerated]
	private FootnoteExplanationMissing A;

	[CompilerGenerated]
	private FootnoteReferenceMissing A;

	[CompilerGenerated]
	private FootnoteSequence A;

	[CompilerGenerated]
	private List<BaseTextCheck> A;

	[CompilerGenerated]
	private List<BaseTextCheck> B;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private HiddenShapes A;

	public ShapeOutsideMargins ShapeOutsideMargins
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

	public ShapeOutOfBounds ShapeOutOfBounds
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

	public MisalignedShapes MisalignedShapes
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

	public RotatedShapes RotatedShapes
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

	public OverlappingText OverlappingText
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

	public MasterShapePosition MasterShapePosition
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

	public PlaceholderFontStyleMismatch PlaceholderFontStyleMismatch
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

	public PlaceholderFontColorMismatch PlaceholderFontColorMismatch
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

	public PlaceholderFillMismatch PlaceholderFillMismatch
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

	public PlaceholderMarginsMismatch PlaceholderMarginsMismatch
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

	public PlaceholderBulletMismatch PlaceholderBulletMismatch
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

	public PlaceholderIndentMismatch PlaceholderIndentMismatch
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

	public SlideTitleCapitalization SlideTitleCapitalization
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

	public SlideNumbers SlideNumbers
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

	public HiddenSlides HiddenSlides
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

	public ProofingLanguage ProofingLanguage
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

	public LineSpacing LineSpacing
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

	public BulletPunctuation BulletPunctuation
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

	public BulletIndentation BulletIndentation
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

	public BulletSize BulletSize
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

	public BulletFontFamily BulletFontFamily
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

	public MultipleFontFamilies MultipleFontFamilies
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

	public FootnoteExplanationMissing FootnoteExplanationMissing
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

	public FootnoteReferenceMissing FootnoteReferenceMissing
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

	public FootnoteSequence FootnoteSequence
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

	public List<BaseTextCheck> ShapeTextChecks
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

	public List<BaseTextCheck> ParagraphTextChecks
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

	public bool CheckPlaceholders
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

	public HiddenShapes HiddenShapes
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

	public Checks(Settings options, Conventions conv, Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		//IL_036b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0370: Unknown result type (might be due to invalid IL or missing references)
		//IL_0390: Unknown result type (might be due to invalid IL or missing references)
		//IL_0395: Unknown result type (might be due to invalid IL or missing references)
		//IL_0161: Unknown result type (might be due to invalid IL or missing references)
		//IL_0166: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01da: Unknown result type (might be due to invalid IL or missing references)
		//IL_0175: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_03fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_041c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0217: Unknown result type (might be due to invalid IL or missing references)
		//IL_021c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0187: Unknown result type (might be due to invalid IL or missing references)
		//IL_018c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0438: Unknown result type (might be due to invalid IL or missing references)
		//IL_043d: Unknown result type (might be due to invalid IL or missing references)
		//IL_023e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0243: Unknown result type (might be due to invalid IL or missing references)
		//IL_0191: Unknown result type (might be due to invalid IL or missing references)
		//IL_0196: Unknown result type (might be due to invalid IL or missing references)
		//IL_045e: Unknown result type (might be due to invalid IL or missing references)
		//IL_025b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0260: Unknown result type (might be due to invalid IL or missing references)
		//IL_019b: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0483: Unknown result type (might be due to invalid IL or missing references)
		//IL_0488: Unknown result type (might be due to invalid IL or missing references)
		//IL_0276: Unknown result type (might be due to invalid IL or missing references)
		//IL_027b: Unknown result type (might be due to invalid IL or missing references)
		//IL_04aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_04af: Unknown result type (might be due to invalid IL or missing references)
		//IL_029d: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_01af: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c7: Invalid comparison between Unknown and I4
		//IL_051d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0522: Unknown result type (might be due to invalid IL or missing references)
		//IL_0544: Unknown result type (might be due to invalid IL or missing references)
		//IL_0549: Unknown result type (might be due to invalid IL or missing references)
		//IL_0561: Unknown result type (might be due to invalid IL or missing references)
		//IL_0566: Unknown result type (might be due to invalid IL or missing references)
		//IL_0588: Unknown result type (might be due to invalid IL or missing references)
		//IL_058d: Unknown result type (might be due to invalid IL or missing references)
		//IL_05be: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0603: Unknown result type (might be due to invalid IL or missing references)
		//IL_0608: Unknown result type (might be due to invalid IL or missing references)
		//IL_0633: Unknown result type (might be due to invalid IL or missing references)
		//IL_065d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0662: Unknown result type (might be due to invalid IL or missing references)
		//IL_0620: Unknown result type (might be due to invalid IL or missing references)
		//IL_0625: Unknown result type (might be due to invalid IL or missing references)
		//IL_069f: Unknown result type (might be due to invalid IL or missing references)
		//IL_064c: Unknown result type (might be due to invalid IL or missing references)
		//IL_06e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_067a: Unknown result type (might be due to invalid IL or missing references)
		//IL_068c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0691: Unknown result type (might be due to invalid IL or missing references)
		//IL_072b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0730: Unknown result type (might be due to invalid IL or missing references)
		//IL_071a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0765: Unknown result type (might be due to invalid IL or missing references)
		//IL_076a: Unknown result type (might be due to invalid IL or missing references)
		//IL_073e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0743: Unknown result type (might be due to invalid IL or missing references)
		//IL_0752: Unknown result type (might be due to invalid IL or missing references)
		//IL_0757: Unknown result type (might be due to invalid IL or missing references)
		//IL_07bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_07c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_07ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_07f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_081f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0824: Unknown result type (might be due to invalid IL or missing references)
		//IL_07dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_07e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_083d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0842: Unknown result type (might be due to invalid IL or missing references)
		//IL_080c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0811: Unknown result type (might be due to invalid IL or missing references)
		//IL_0865: Unknown result type (might be due to invalid IL or missing references)
		//IL_0889: Unknown result type (might be due to invalid IL or missing references)
		//IL_0876: Unknown result type (might be due to invalid IL or missing references)
		//IL_087b: Unknown result type (might be due to invalid IL or missing references)
		//IL_08a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_08aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_08c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_08c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_08eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0987: Unknown result type (might be due to invalid IL or missing references)
		//IL_098c: Unknown result type (might be due to invalid IL or missing references)
		//IL_09d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_09d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a21: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a26: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a7b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a80: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a99: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a9e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0abf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ac4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b05: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b3a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b49: Unknown result type (might be due to invalid IL or missing references)
		ShapeOutsideMargins = null;
		ShapeOutOfBounds = null;
		MisalignedShapes = null;
		RotatedShapes = null;
		OverlappingText = null;
		MasterShapePosition = null;
		PlaceholderFontStyleMismatch = null;
		PlaceholderFontColorMismatch = null;
		PlaceholderFillMismatch = null;
		PlaceholderMarginsMismatch = null;
		PlaceholderBulletMismatch = null;
		PlaceholderIndentMismatch = null;
		SlideTitleCapitalization = null;
		SlideNumbers = null;
		HiddenSlides = null;
		ProofingLanguage = null;
		LineSpacing = null;
		BulletPunctuation = null;
		BulletIndentation = null;
		BulletSize = null;
		BulletFontFamily = null;
		MultipleFontFamilies = null;
		FootnoteExplanationMissing = null;
		FootnoteReferenceMissing = null;
		FootnoteSequence = null;
		ShapeTextChecks = null;
		ParagraphTextChecks = null;
		CheckPlaceholders = false;
		HiddenShapes = null;
		bool flag = false;
		MsoLanguageID defaultLanguageId = (MsoLanguageID)((Settings)options).DefaultLanguageId;
		if (defaultLanguageId <= MsoLanguageID.msoLanguageIDEnglishAUS)
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
			if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishUS)
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
				if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishUK)
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
					if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishAUS)
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
						goto IL_015d;
					}
				}
			}
		}
		else if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishCanadian)
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
			if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishNewZealand)
			{
				if (defaultLanguageId != MsoLanguageID.msoLanguageIDEnglishIreland)
				{
					goto IL_015d;
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
		}
		flag = true;
		goto IL_015d;
		IL_0952:
		List<BaseTextCheck> shapeTextChecks;
		shapeTextChecks.Add(new HyphenSpacingImproper());
		if (flag)
		{
			shapeTextChecks.Add(new GrammarAn());
		}
		shapeTextChecks = null;
		ParagraphTextChecks = new List<BaseTextCheck>();
		List<BaseTextCheck> paragraphTextChecks = ParagraphTextChecks;
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationSpacingIncorrect))
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
			paragraphTextChecks.Add(new SpaceBeforeClosing());
			paragraphTextChecks.Add(new SpaceBeforeOpening());
			paragraphTextChecks.Add(new SpaceAfterClosing());
			paragraphTextChecks.Add(new SpaceAfterOpening());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationMissing))
		{
			paragraphTextChecks.Add(new MissingStraightQuotes());
			paragraphTextChecks.Add(new MissingFancyQuotes());
			paragraphTextChecks.Add(new MissingParentheses());
			paragraphTextChecks.Add(new MissingBrackets());
			paragraphTextChecks.Add(new MissingBraces());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ConfusedWords))
		{
			paragraphTextChecks.Add(new GrammarEnsureInsure());
			paragraphTextChecks.Add(new GrammarAffectEffect());
			paragraphTextChecks.Add(new GrammarAdverseAverse());
			paragraphTextChecks.Add(new GrammarAcceptExcept());
			paragraphTextChecks.Add(new GrammarComplimentComplement());
			paragraphTextChecks.Add(new GrammarIts());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).GrammarIeEg))
		{
			paragraphTextChecks.Add(new GrammarIeEg());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).RepeatedWords))
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
			paragraphTextChecks.Add(new RepeatedWords());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).HyphenWordsInconsistent) && ((Conventions)conv).HyphenatedWords.Count > 0)
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
			paragraphTextChecks.Add(new HyphenWordsInconsistent(((Conventions)conv).HyphenatedWords, ((Conventions)conv).UnhyphenatedWords));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).CasualWriting))
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
			paragraphTextChecks.Add(new Contractions());
		}
		paragraphTextChecks.Add(new Uncontractions());
		paragraphTextChecks = null;
		((Checks)this).IsCheckEnabled(options.CheckAgendaUpdated);
		if (((Checks)this).IsCheckEnabled(((Settings)options).ColorPalette))
		{
			((Checks)this).PopulatePaletteColors();
		}
		return;
		IL_01cc:
		int checkPlaceholders;
		CheckPlaceholders = (byte)checkPlaceholders != 0;
		Settings settings = null;
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderFontStyleMismatch))
		{
			PlaceholderFontStyleMismatch = new PlaceholderFontStyleMismatch();
		}
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderFontColorMismatch))
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
			PlaceholderFontColorMismatch = new PlaceholderFontColorMismatch();
		}
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderFillMismatch))
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
			PlaceholderFillMismatch = new PlaceholderFillMismatch();
		}
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderMarginMismatch))
		{
			PlaceholderMarginsMismatch = new PlaceholderMarginsMismatch();
		}
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderBulletMismatch))
		{
			PlaceholderBulletMismatch = new PlaceholderBulletMismatch();
		}
		if (((Checks)this).IsCheckEnabled(options.CheckPlaceholderIndentMismatch))
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
			PlaceholderIndentMismatch = new PlaceholderIndentMismatch();
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ShapeOutsideMargins))
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
				PowerPointAddIn1.Template.Settings settings2 = new PowerPointAddIn1.Template.Settings(pres);
				if (settings2.SlideMargins.HasValue)
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
					PowerPointAddIn1.Template.Settings.Margins value = settings2.SlideMargins.Value;
					ShapeOutsideMargins = new ShapeOutsideMargins(value.Left, pres.PageSetup.SlideWidth - value.Right);
				}
				else
				{
					Microsoft.Office.Interop.PowerPoint.Shape bodyPlaceholder = Helpers.GetBodyPlaceholder(pres);
					ShapeOutsideMargins = new ShapeOutsideMargins(bodyPlaceholder.Left, bodyPlaceholder.Left + bodyPlaceholder.Width);
					bodyPlaceholder = null;
				}
				settings2 = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ShapeOutOfBounds))
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
			ShapeOutOfBounds = new ShapeOutOfBounds();
		}
		if (((Checks)this).IsCheckEnabled(options.MisalignedShapes))
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
			MisalignedShapes = new MisalignedShapes();
		}
		if (((Checks)this).IsCheckEnabled(options.RotatedShapes))
		{
			RotatedShapes = new RotatedShapes();
		}
		if (((Checks)this).IsCheckEnabled(options.OverlappingText))
		{
			OverlappingText = new OverlappingText(conv.TextRangeBounds);
		}
		if (((Checks)this).IsCheckEnabled(options.MasterShapePosition))
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
			MasterShapePosition = new MasterShapePosition();
		}
		if (((Checks)this).IsCheckEnabled(options.SlideTitles))
		{
			SlideTitleCapitalization = new SlideTitleCapitalization(conv);
		}
		if (((Checks)this).IsCheckEnabled(options.SlideNumbers))
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
			SlideNumbers = new SlideNumbers(conv);
		}
		if (((Checks)this).IsCheckEnabled(options.HiddenSlides))
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
			HiddenSlides = new HiddenSlides();
		}
		if (((Checks)this).IsCheckEnabled(options.HiddenShapes))
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
			HiddenShapes = new HiddenShapes();
		}
		if (((Checks)this).IsCheckEnabled(options.LineSpacing))
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
			LineSpacing = new LineSpacing();
		}
		if (((Checks)this).IsCheckEnabled(options.BulletPunctuation))
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
			BulletPunctuation = new BulletPunctuation();
		}
		if (((Checks)this).IsCheckEnabled(options.BulletSize))
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
			BulletSize = new BulletSize();
		}
		if (((Checks)this).IsCheckEnabled(options.BulletIndentation))
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
			BulletIndentation = new BulletIndentation();
		}
		if (((Checks)this).IsCheckEnabled(options.BulletFontFamily))
		{
			BulletFontFamily = new BulletFontFamily();
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).MultipleFontFamilies))
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
			MultipleFontFamilies = new MultipleFontFamilies();
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ProofingLanguage))
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
			ProofingLanguage = new ProofingLanguage(Conversions.ToString(((Settings)options).DefaultLanguageId));
		}
		if (((Checks)this).IsCheckEnabled(options.Footnotes))
		{
			FootnoteExplanationMissing = new FootnoteExplanationMissing(conv.FootnoteNumbers);
			FootnoteReferenceMissing = new FootnoteReferenceMissing();
		}
		ShapeTextChecks = new List<BaseTextCheck>();
		shapeTextChecks = ShapeTextChecks;
		if (((Checks)this).IsCheckEnabled(((Settings)options).SentenceSpacing))
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
			shapeTextChecks.Add(new SentenceSpacing(((Settings)options).SentenceSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ColonSpacing))
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
			shapeTextChecks.Add(new ColonSpacing(((Settings)options).ColonSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationSpacingInconsistent))
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
			shapeTextChecks.Add(new SlashSpacingInconsistent(((Settings)options).SlashSpacingConvention));
			shapeTextChecks.Add(new DashSpacingInconsistent(((Settings)options).DashSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationSpacingIncorrect))
		{
			shapeTextChecks.Add(new SlashSpacingUnbalanced(((Settings)options).SlashSpacingConvention));
			shapeTextChecks.Add(new DashSpacingUnbalanced(((Settings)options).DashSpacingConvention));
			shapeTextChecks.Add(new DoubleSpace());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).MillionsBillionsAbbreviation))
		{
			shapeTextChecks.Add(new AbbreviationMillions(((Settings)options).MillionsAbbreviationConvention));
			shapeTextChecks.Add(new AbbreviationBillions(((Settings)options).BillionsAbbreviationConvention));
			shapeTextChecks.Add(new AbbreviationSpacing(((Settings)options).UnitsSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).QuotesStyle))
		{
			shapeTextChecks.Add(new DoubleQuoteStyle(((Settings)options).QuotesStyleConvention));
			shapeTextChecks.Add(new SingleQuoteStyle(((Settings)options).QuotesStyleConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).DummyText))
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
			shapeTextChecks.Add(new DummyText());
			shapeTextChecks.Add(new LoremIpsum());
		}
		if (((Settings)options).IsCorporateDictionaryInUse())
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
			shapeTextChecks.Add(new CorporateDictionary(((Settings)options).CorporateDictionaryRules));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).SpellingCanceled))
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
			shapeTextChecks.Add(new SpellingCanceled(((Settings)options).CanceledSpellingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).SpellingAdviser))
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
			shapeTextChecks.Add(new SpellingAdviser(((Settings)options).AdviserSpellingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).GrammarMyriadOf))
		{
			shapeTextChecks.Add(new GrammarMyriadOf());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).GrammarAsPer))
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
			shapeTextChecks.Add(new GrammarAsPer());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).IeEgComma))
		{
			shapeTextChecks.Add(new IeEgComma(((Settings)options).IeEgCommaConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).UnnecessaryPeriods))
		{
			shapeTextChecks.Add(new UnnecessaryPeriods());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PassiveVoice))
		{
			shapeTextChecks.Add(new PassiveVoice());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).CasualWriting))
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
			shapeTextChecks.Add(new QuestionMark());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).HyphenWordsImproper))
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
			shapeTextChecks.Add(new HyphenWordsImproper());
			shapeTextChecks.Add(new HyphenatedAdverb());
			if (((Settings)options).CorporateDictionaryRules != null)
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
				if (((Settings)options).CorporateDictionaryRules.Count != 0)
				{
					goto IL_0952;
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
			shapeTextChecks.Add(new HyphenMissing());
		}
		goto IL_0952;
		IL_015d:
		settings = options;
		if ((int)settings.CheckPlaceholderLayoutMismatch == 0)
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
			if ((int)settings.CheckPlaceholderFillMismatch == 0)
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
				if ((int)settings.CheckPlaceholderFontColorMismatch == 0 && (int)settings.CheckPlaceholderFontStyleMismatch == 0 && (int)settings.CheckPlaceholderMarginMismatch == 0)
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
					if ((int)settings.CheckPlaceholderBulletMismatch == 0)
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
						checkPlaceholders = (((int)settings.CheckPlaceholderIndentMismatch > 0) ? 1 : 0);
						goto IL_01cc;
					}
				}
			}
		}
		checkPlaceholders = 1;
		goto IL_01cc;
	}
}
