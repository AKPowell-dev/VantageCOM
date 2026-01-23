using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class SensitiveData : RawObservation
{
	[CompilerGenerated]
	private string A;

	internal string DataType
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

	public SensitiveData(Range rng, string strDataType)
		: base(rng)
	{
		DataType = strDataType;
	}
}
