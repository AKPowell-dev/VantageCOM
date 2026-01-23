using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Resources;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
[StandardModule]
[HideModuleName]
[CompilerGenerated]
[DebuggerNonUserCode]
internal sealed class M
{
	private static ResourceManager A;

	private static CultureInfo A;

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	internal static ResourceManager ResourceManager
	{
		get
		{
			if (object.ReferenceEquals(M.A, null))
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
				M.A = new ResourceManager(XC.A(341), typeof(M).Assembly);
			}
			return M.A;
		}
	}

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	internal static CultureInfo Culture
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

	internal static Bitmap BorderColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(390), A));

	internal static Bitmap ChartSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(425), A));

	internal static Bitmap CloseOthers => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(446), A));

	internal static Bitmap ContentTypeShapes => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(469), A));

	internal static string DefaultShortcuts => ResourceManager.GetString(XC.A(504), A);

	internal static Bitmap Download => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(537), A));

	internal static Bitmap EmbeddedExcel => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(554), A));

	internal static Bitmap FillColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(581), A));

	internal static Bitmap Find => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(612), A));

	internal static Bitmap FolderOpen => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(621), A));

	internal static Bitmap FontColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(642), A));

	internal static Bitmap Gear => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(673), A));

	internal static Bitmap GearSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(682), A));

	internal static Bitmap GroupFormatText => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(701), A));

	internal static Bitmap HighlightColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(732), A));

	internal static Bitmap ImportToWord => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(773), A));

	internal static Bitmap LinkManager => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(798), A));

	internal static Bitmap LinkWizard => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(821), A));

	internal static Bitmap NoBorder => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(842), A));

	internal static Bitmap NoColor => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(859), A));

	internal static Bitmap NoFill => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(874), A));

	internal static Bitmap NoHighlight => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(887), A));

	internal static Bitmap PdfIconSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(910), A));

	internal static Bitmap Picture => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(935), A));

	internal static Bitmap PrintArea => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(950), A));

	internal static Bitmap Proof => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(969), A));

	internal static Bitmap Redact => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(980), A));

	internal static string Ribbon => ResourceManager.GetString(XC.A(993), A);

	internal static Bitmap Table => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(XC.A(1006), A));
}
