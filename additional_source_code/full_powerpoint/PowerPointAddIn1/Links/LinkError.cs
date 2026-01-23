using System.Runtime.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class LinkError
{
	[CompilerGenerated]
	private object A;

	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	public object LinkedObject
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public string Name
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

	public string ErrorMessage
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

	public LinkError(object obj, string strName, string strError)
	{
		LinkedObject = RuntimeHelpers.GetObjectValue(obj);
		Name = strName;
		ErrorMessage = strError;
	}
}
