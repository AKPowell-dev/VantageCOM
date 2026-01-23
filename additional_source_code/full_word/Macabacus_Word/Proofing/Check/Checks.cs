using System;
using System.Collections.Generic;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class Checks : Checks
{
	private ShapeOutsideMargins A;

	private ProofingLanguage A;

	private MultipleFontFamilies A;

	private List<BaseTextCheck> A;

	public ShapeOutsideMargins ShapeOutsideMargins
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public ProofingLanguage ProofingLanguage
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public MultipleFontFamilies MultipleFontFamilies
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public List<BaseTextCheck> TextChecks
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public Checks(Settings options, Conventions conv, Microsoft.Office.Interop.Word.Document doc)
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0194: Unknown result type (might be due to invalid IL or missing references)
		//IL_0199: Unknown result type (might be due to invalid IL or missing references)
		//IL_014a: Unknown result type (might be due to invalid IL or missing references)
		//IL_014f: Unknown result type (might be due to invalid IL or missing references)
		//IL_015d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0162: Unknown result type (might be due to invalid IL or missing references)
		//IL_0170: Unknown result type (might be due to invalid IL or missing references)
		//IL_0181: Unknown result type (might be due to invalid IL or missing references)
		//IL_0186: Unknown result type (might be due to invalid IL or missing references)
		//IL_0129: Unknown result type (might be due to invalid IL or missing references)
		//IL_020d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0212: Unknown result type (might be due to invalid IL or missing references)
		//IL_0260: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_02c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_027a: Unknown result type (might be due to invalid IL or missing references)
		//IL_027f: Unknown result type (might be due to invalid IL or missing references)
		//IL_028d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0292: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_02f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0313: Unknown result type (might be due to invalid IL or missing references)
		//IL_0318: Unknown result type (might be due to invalid IL or missing references)
		//IL_033a: Unknown result type (might be due to invalid IL or missing references)
		//IL_033f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0361: Unknown result type (might be due to invalid IL or missing references)
		//IL_0366: Unknown result type (might be due to invalid IL or missing references)
		//IL_037e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0383: Unknown result type (might be due to invalid IL or missing references)
		//IL_039b: Unknown result type (might be due to invalid IL or missing references)
		//IL_03be: Unknown result type (might be due to invalid IL or missing references)
		//IL_0405: Unknown result type (might be due to invalid IL or missing references)
		//IL_040a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0450: Unknown result type (might be due to invalid IL or missing references)
		//IL_0455: Unknown result type (might be due to invalid IL or missing references)
		this.A = null;
		this.A = null;
		this.A = null;
		A = null;
		if (((Checks)this).IsCheckEnabled(((Settings)options).ShapeOutsideMargins))
		{
			try
			{
				PageSetup pageSetup = doc.PageSetup;
				ShapeOutsideMargins = new ShapeOutsideMargins(Math.Round(pageSetup.LeftMargin, 4), Math.Round(pageSetup.RightMargin, 4), Math.Round(pageSetup.TopMargin, 4), Math.Round(pageSetup.BottomMargin, 4));
				pageSetup = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ProofingLanguage))
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
			ProofingLanguage = new ProofingLanguage(Conversions.ToString(((Settings)options).DefaultLanguageId));
		}
		TextChecks = new List<BaseTextCheck>();
		List<BaseTextCheck> textChecks = TextChecks;
		if (((Checks)this).IsCheckEnabled(((Settings)options).MillionsBillionsAbbreviation))
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
			textChecks.Add(new AbbreviationMillions(((Settings)options).MillionsAbbreviationConvention));
			textChecks.Add(new AbbreviationBllions(((Settings)options).BillionsAbbreviationConvention));
			textChecks.Add(new AbbreviationSpacing(((Settings)options).UnitsSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationSpacingInconsistent))
		{
			textChecks.Add(new SentenceSpacing(((Settings)options).SentenceSpacingConvention));
			textChecks.Add(new ColonSpacing(((Settings)options).ColonSpacingConvention));
			textChecks.Add(new SlashSpacingInconsistent(((Settings)options).SlashSpacingConvention));
			textChecks.Add(new DashSpacingInconsistent(((Settings)options).DashSpacingConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationSpacingIncorrect))
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
			textChecks.Add(new SlashSpacingUnbalanced(((Settings)options).SlashSpacingConvention));
			textChecks.Add(new DashSpacingUnbalanced(((Settings)options).DashSpacingConvention));
			textChecks.Add(new SpaceBeforeClosing());
			textChecks.Add(new SpaceBeforeOpening());
			textChecks.Add(new SpaceAfterClosing());
			textChecks.Add(new SpaceAfterOpening());
			textChecks.Add(new DoubleSpace());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).PunctuationMissing))
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
			textChecks.Add(new MissingStraightQuotes());
			textChecks.Add(new MissingFancyQuotes());
			textChecks.Add(new MissingParentheses());
			textChecks.Add(new MissingBrackets());
			textChecks.Add(new MissingBraces());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).QuotesStyle))
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
			textChecks.Add(new DoubleQuoteStyle(((Settings)options).QuotesStyleConvention));
			textChecks.Add(new SingleQuoteStyle(((Settings)options).QuotesStyleConvention));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).GrammarMyriadOf))
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
			textChecks.Add(new GrammarMyriadOf());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).GrammarAsPer))
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
			textChecks.Add(new GrammarAsPer());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).ConfusedWords))
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
			textChecks.Add(new GrammarEnsureInsure());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).SpellingCanceled))
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
			textChecks.Add(new SpellingCanceled());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).RepeatedWords))
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
			textChecks.Add(new RepeatedWords());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).HyphenWordsImproper))
		{
			textChecks.Add(new HyphenWordsImproper());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).IeEgComma))
		{
			textChecks.Add(new CommaMissing());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).DummyText))
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
			textChecks.Add(new DummyText());
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).HyphenWordsInconsistent) && ((Conventions)conv).HyphenatedWords.Count > 0)
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
			textChecks.Add(new HyphenWordsInconsistent(((Conventions)conv).HyphenatedWords, ((Conventions)conv).UnhyphenatedWords));
		}
		if (((Checks)this).IsCheckEnabled(((Settings)options).CasualWriting))
		{
			textChecks.Add(new QuestionMark());
			textChecks.Add(new Contractions());
		}
		textChecks.Add(new HyphenSpacingImproper());
		textChecks.Add(new HyphenatedAdverb());
		textChecks.Add(new Uncontractions());
		textChecks = null;
		if (!((Checks)this).IsCheckEnabled(((Settings)options).ColorPalette))
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
			((Checks)this).PopulatePaletteColors();
			return;
		}
	}
}
