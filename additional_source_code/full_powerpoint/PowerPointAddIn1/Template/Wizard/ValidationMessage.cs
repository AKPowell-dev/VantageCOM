using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros;

namespace PowerPointAddIn1.Template.Wizard;

public sealed class ValidationMessage
{
	[CompilerGenerated]
	private string A;

	[CompilerGenerated]
	private string B;

	[CompilerGenerated]
	private SolidColorBrush A;

	public string Message
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

	public string IconData
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

	public SolidColorBrush IconColor
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

	public ValidationMessage(string strMessage, ValidationType type)
	{
		Message = strMessage;
		Color color;
		switch (type)
		{
		case ValidationType.WarningLevel:
			IconData = AH.A(121061);
			color = clsColors.SeverityColorYellow();
			break;
		case ValidationType.ErrorLevel:
			IconData = AH.A(121158);
			color = clsColors.SeverityColorRed();
			break;
		case ValidationType.InfoLevel:
			IconData = AH.A(121344);
			color = clsColors.SeverityColorBlue();
			break;
		default:
		{
			IconData = AH.A(121554);
			object obj = ColorConverter.ConvertFromString(AH.A(121830));
			Color obj2;
			if (obj == null)
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
				obj2 = default(Color);
			}
			else
			{
				obj2 = (Color)obj;
			}
			color = obj2;
			break;
		}
		}
		IconColor = new SolidColorBrush(color);
	}
}
