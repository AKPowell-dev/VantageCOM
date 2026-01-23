using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Links;

public sealed class Slides
{
	internal static void A(Slide A, string B, string C, bool D)
	{
		string authorFromSlide = Common.GetAuthorFromSlide(A);
		string text = Base.LastUpdate();
		string lastModifiedTime = Updates.GetLastModifiedTime(B);
		Tags tags = A.Tags;
		tags.Add(Value: Add.GenerateXml(tags[Base.TAG_LINK_XML], (ImportType)10, B, lastModifiedTime, text, authorFromSlide, "", C, "", "", (bool?)D), Name: Base.TAG_LINK_XML);
		_ = null;
	}
}
