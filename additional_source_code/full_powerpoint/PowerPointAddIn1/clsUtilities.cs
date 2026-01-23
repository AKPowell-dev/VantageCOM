using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1;

public sealed class clsUtilities
{
	private struct GG
	{
		public string A;

		public string B;

		public string C;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<GG, string> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal string A(GG A)
		{
			return A.C;
		}
	}

	public static void StartupProcedures(XmlDocument xmlSettings)
	{
		AirplaneMode.Startup();
		KG.A = new clsSettings(xmlSettings);
	}

	public static void EnumerateLanguages()
	{
		XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
		StringBuilder stringBuilder = new StringBuilder();
		List<GG> list = new List<GG>();
		Application application = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid(AH.A(161986))));
		application.Visible = false;
		MsoLanguageID msoLanguageID = MsoLanguageID.msoLanguageIDArabic;
		object Index;
		do
		{
			try
			{
				Languages languages = application.Languages;
				Index = msoLanguageID;
				Language language = languages[ref Index];
				string a = ((int)language.ID).ToString();
				string name = language.Name;
				string nameLocal = language.NameLocal;
				_ = null;
				list.Add(new GG
				{
					A = a,
					B = name,
					C = nameLocal
				});
				GG gG = default(GG);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			msoLanguageID++;
		}
		while (msoLanguageID <= MsoLanguageID.msoLanguageIDSpanishPuertoRico);
		Application application2 = application;
		Index = RuntimeHelpers.GetObjectValue(Missing.Value);
		object OriginalFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
		object RouteDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
		application2.Quit(ref Index, ref OriginalFormat, ref RouteDocument);
		application = null;
		List<GG> source = list;
		Func<GG, string> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (GG A) => A.C);
		}
		else
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
			keySelector = _Closure_0024__.A;
		}
		list = source.OrderBy(keySelector).ToList();
		xmlWriterSettings.Indent = true;
		xmlWriterSettings.CloseOutput = true;
		xmlWriterSettings.NewLineHandling = NewLineHandling.None;
		XmlWriter xmlWriter = XmlWriter.Create(stringBuilder, xmlWriterSettings);
		xmlWriter.WriteStartDocument();
		xmlWriter.WriteStartElement(AH.A(10641));
		foreach (GG item in list)
		{
			XmlWriter xmlWriter2 = xmlWriter;
			xmlWriter2.WriteStartElement(AH.A(162059));
			xmlWriter2.WriteAttributeString(AH.A(58243), item.A);
			xmlWriter2.WriteAttributeString(AH.A(63335), item.B);
			xmlWriter2.WriteAttributeString(AH.A(58224), item.C);
			xmlWriter2.WriteEndElement();
			_ = null;
		}
		xmlWriter.WriteEndElement();
		xmlWriter.WriteEndDocument();
		xmlWriter.Flush();
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(stringBuilder.ToString());
		xmlDocument.Save(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), AH.A(162076)));
		_ = null;
		xmlWriter = null;
		xmlWriterSettings = null;
		stringBuilder = null;
		list = null;
	}
}
