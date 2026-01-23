using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes;

public sealed class Animation
{
	public static void Remove(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		AnimationSettings animationSettings = shp.AnimationSettings;
		animationSettings.AfterEffect = PpAfterEffect.ppAfterEffectNothing;
		animationSettings.AnimateBackground = MsoTriState.msoFalse;
		animationSettings.EntryEffect = PpEntryEffect.ppEffectNone;
		animationSettings.SoundEffect.Type = PpSoundEffectType.ppSoundNone;
		animationSettings.TextLevelEffect = PpTextLevelEffect.ppAnimateLevelNone;
		_ = null;
	}
}
