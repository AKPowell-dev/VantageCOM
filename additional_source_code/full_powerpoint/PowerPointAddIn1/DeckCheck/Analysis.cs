using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Check;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Analysis
{
	[CompilerGenerated]
	private Settings A;

	[CompilerGenerated]
	private Conventions A;

	[CompilerGenerated]
	private Checks A;

	[CompilerGenerated]
	private List<BaseError> A;

	public Settings Options
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

	public Conventions Conventions
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

	public Checks Checks
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

	public List<BaseError> Errors
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

	public Analysis(Microsoft.Office.Interop.PowerPoint.Presentation pres, List<string> unexpectedErrors)
	{
		try
		{
			Options = new Settings();
			Conventions = new Conventions(pres, Options, unexpectedErrors);
			Checks = new Checks(Options, Conventions, pres);
			Errors = new List<BaseError>();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			unexpectedErrors.Add(string.Format(AH.A(47123), AH.A(7894), ex2.Message));
			ProjectData.ClearProjectError();
		}
	}
}
