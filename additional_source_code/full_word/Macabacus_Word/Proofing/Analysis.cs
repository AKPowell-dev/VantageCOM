using System.Collections.Generic;
using Macabacus_Word.Proofing.Check;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing;

public sealed class Analysis
{
	private Settings A;

	private Conventions A;

	private Checks A;

	private List<BaseError> A;

	public Settings Options
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

	public Conventions Conventions
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

	public Checks Checks
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

	public List<BaseError> Errors
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

	public Analysis(Microsoft.Office.Interop.Word.Document doc)
	{
		Options = new Settings();
		Conventions = new Conventions(doc, Options);
		Checks = new Checks(Options, Conventions, doc);
		Errors = new List<BaseError>();
	}
}
