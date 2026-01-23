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

[StandardModule]
[HideModuleName]
[DebuggerNonUserCode]
[GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0")]
[CompilerGenerated]
internal sealed class OB
{
	private static ResourceManager A;

	private static CultureInfo A;

	[EditorBrowsable(EditorBrowsableState.Advanced)]
	internal static ResourceManager ResourceManager
	{
		get
		{
			if (object.ReferenceEquals(OB.A, null))
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
				OB.A = new ResourceManager(AH.A(1885), typeof(OB).Assembly);
			}
			return OB.A;
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

	internal static Bitmap AirplaneMode => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(1938), A));

	internal static Bitmap AnchorBottomLeft => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(1963), A));

	internal static Bitmap AnchorBottomRight => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(1996), A));

	internal static Bitmap AnchorCenter => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2031), A));

	internal static Bitmap AnchorSwap => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2056), A));

	internal static Bitmap AnchorTopLeft => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2077), A));

	internal static Bitmap AnchorTopRight => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2104), A));

	internal static Bitmap AutofitToggle => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2133), A));

	internal static Bitmap BlankSlide => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2160), A));

	internal static Bitmap BorderColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2181), A));

	internal static Bitmap ChartSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2216), A));

	internal static Bitmap CloseOthers => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2237), A));

	internal static Bitmap Contact => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2260), A));

	internal static Bitmap ContentTypeShapes => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2275), A));

	internal static Bitmap Copy => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2310), A));

	internal static Bitmap DistributeHorizontally => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2319), A));

	internal static Bitmap DistributeVertically => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2364), A));

	internal static Bitmap Download => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2405), A));

	internal static Bitmap EmbeddedExcel => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2422), A));

	internal static Icon favicon => (Icon)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2449), A));

	internal static Bitmap FileCheckIn => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2464), A));

	internal static Bitmap FileSaveAsPowerPointPptx => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2487), A));

	internal static Bitmap FillColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2536), A));

	internal static Bitmap Find => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2567), A));

	internal static Bitmap FixBullets => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2576), A));

	internal static Bitmap Flysheet => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2597), A));

	internal static Bitmap FlysheetStyleAgenda => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2614), A));

	internal static Bitmap FlysheetStyleTopic => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2653), A));

	internal static Bitmap FolderOpen => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2690), A));

	internal static Bitmap FontColorPicker => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2711), A));

	internal static Bitmap FootnoteAdd => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2742), A));

	internal static Bitmap FootnoteRemove => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2765), A));

	internal static Bitmap Gear => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2794), A));

	internal static Bitmap GearSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2803), A));

	internal static string Generic => ResourceManager.GetString(AH.A(2822), A);

	internal static Bitmap GroupFormatText => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2837), A));

	internal static Bitmap Help => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2868), A));

	internal static Bitmap ImportFromExcel => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2877), A));

	internal static Bitmap InsertDrawingCanvas => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2908), A));

	internal static Bitmap LinkManager => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2947), A));

	internal static Bitmap LinkWizard => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2970), A));

	internal static Bitmap LogoExcel => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(2991), A));

	internal static Bitmap LogoWord => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3010), A));

	internal static Bitmap MemorizeLeft => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3027), A));

	internal static Bitmap MemorizePosition => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3052), A));

	internal static Bitmap MemorizeTop => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3085), A));

	internal static Bitmap MergeText => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3108), A));

	internal static Bitmap MultiplyShape => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3127), A));

	internal static Bitmap NoBorder => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3154), A));

	internal static Bitmap NoColor => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3171), A));

	internal static Bitmap NoFill => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3186), A));

	internal static Bitmap Numbering => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3199), A));

	internal static Bitmap OmitDoubleFlysheets => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3218), A));

	internal static Bitmap Paste => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3257), A));

	internal static Bitmap PdfIconSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3268), A));

	internal static Bitmap Picture => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3293), A));

	internal static Bitmap PrintArea => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3308), A));

	internal static Bitmap Proof => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3327), A));

	internal static Bitmap ProofOkSmall => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3338), A));

	internal static Bitmap Record => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3363), A));

	internal static Bitmap Redact => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3376), A));

	internal static string Ribbon => ResourceManager.GetString(AH.A(3389), A);

	internal static Bitmap SaveAll => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3402), A));

	internal static Bitmap SectionAdd => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3417), A));

	internal static Bitmap SetPertWeights => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3438), A));

	internal static Bitmap SkipDoubleFlysheets => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3467), A));

	internal static Bitmap SlideLayoutGallery => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3506), A));

	internal static Bitmap SplitText => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3543), A));

	internal static Bitmap StackDown => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3562), A));

	internal static Bitmap StackLeft => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3581), A));

	internal static Bitmap StackRight => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3600), A));

	internal static Bitmap StackUp => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3621), A));

	internal static Bitmap SwapText => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3636), A));

	internal static Bitmap Table => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3653), A));

	internal static Bitmap TitlePage => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3664), A));

	internal static Bitmap TocExclude => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3683), A));

	internal static Bitmap TocInclude => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3704), A));

	internal static Bitmap TocShowSubsections => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3725), A));

	internal static Bitmap Topic => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3762), A));

	internal static Bitmap TurboShapeArrow => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3773), A));

	internal static Bitmap TurboShapeHarvey => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3804), A));

	internal static Bitmap TurboShapePentagon => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3837), A));

	internal static Bitmap TurboShapeProgress => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3874), A));

	internal static Bitmap TurboShapeRectangle => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3911), A));

	internal static Bitmap TurboShapeSwitch => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3950), A));

	internal static Bitmap TurboShapeTachometer => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(3983), A));

	internal static Bitmap TurboShapeThermometer => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(4024), A));

	internal static Bitmap TurboShapeTrafficLight => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(4067), A));

	internal static Bitmap UngroupTable => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(4112), A));

	internal static Bitmap Wizard => (Bitmap)RuntimeHelpers.GetObjectValue(ResourceManager.GetObject(AH.A(4137), A));
}
