using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.Versioning;

public sealed class Check
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ShapeItem, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(ShapeItem A)
		{
			return ((ContentItem)A).IsOutdated;
		}
	}

	[CompilerGenerated]
	internal sealed class P
	{
		public Document A;

		public Action A;

		[SpecialName]
		internal void A()
		{
			try
			{
				Dispatcher a = Check.m_A;
				Action callback;
				if (this.A != null)
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
					callback = this.A;
				}
				else
				{
					callback = (this.A = [SpecialName] () =>
					{
						Check.A(this.A, B: false);
						List<ShapeItem> libraryObjects = LibraryObjects;
						Func<ShapeItem, bool> predicate;
						if (_Closure_0024__.A == null)
						{
							predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
						}
						else
						{
							while (true)
							{
								switch (1)
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
							predicate = _Closure_0024__.A;
						}
						if (libraryObjects.Where(predicate).Count() > 0)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									if (UIFormsExtensions.AskYesNo((System.Windows.Window)null, XC.A(43960), true, true))
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												Check.A();
												return;
											}
										}
									}
									return;
								}
							}
						}
					});
				}
				a.Invoke(callback);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}

		[SpecialName]
		internal void B()
		{
			Check.A(this.A, B: false);
			List<ShapeItem> libraryObjects = LibraryObjects;
			Func<ShapeItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
			}
			else
			{
				while (true)
				{
					switch (1)
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
				predicate = _Closure_0024__.A;
			}
			if (libraryObjects.Where(predicate).Count() <= 0)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (!UIFormsExtensions.AskYesNo((System.Windows.Window)null, XC.A(43960), true, true))
				{
					return;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					Check.A();
					return;
				}
			}
		}
	}

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static List<ShapeItem> m_A;

	private static Dispatcher m_A;

	internal static bool CheckOutdatedLibraryContent
	{
		[CompilerGenerated]
		get
		{
			return Check.m_A;
		}
		[CompilerGenerated]
		set
		{
			Check.m_A = value;
		}
	}

	internal static List<ShapeItem> LibraryObjects
	{
		[CompilerGenerated]
		get
		{
			return Check.m_A;
		}
		[CompilerGenerated]
		set
		{
			Check.m_A = value;
		}
	}

	internal static void A(Document A)
	{
		if (!CheckOutdatedLibraryContent)
		{
			return;
		}
		Action A2 = default(Action);
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (A.Windows.Count <= 0)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				Check.m_A = Dispatcher.CurrentDispatcher;
				Thread thread = new Thread([SpecialName] () =>
				{
					try
					{
						Dispatcher a = Check.m_A;
						Action callback;
						if (A2 != null)
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
							callback = A2;
						}
						else
						{
							callback = (A2 = [SpecialName] () =>
							{
								Check.A(A, B: false);
								List<ShapeItem> libraryObjects = LibraryObjects;
								Func<ShapeItem, bool> predicate;
								if (_Closure_0024__.A == null)
								{
									predicate = (_Closure_0024__.A = _Closure_0024__.A.A);
								}
								else
								{
									while (true)
									{
										switch (1)
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
									predicate = _Closure_0024__.A;
								}
								if (libraryObjects.Where(predicate).Count() > 0)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											if (UIFormsExtensions.AskYesNo((System.Windows.Window)null, XC.A(43960), true, true))
											{
												while (true)
												{
													switch (5)
													{
													case 0:
														break;
													default:
														Check.A();
														return;
													}
												}
											}
											return;
										}
									}
								}
							});
						}
						a.Invoke(callback);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				});
				thread.IsBackground = true;
				thread.SetApartmentState(ApartmentState.STA);
				thread.Start();
				return;
			}
		}
	}

	private static void A()
	{
		try
		{
			Pane.A();
			Pane.B();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
	}

	internal static void B()
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		if (application.Documents.Count > 0)
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
			A(application.ActiveDocument, B: true);
			if (LibraryObjects.Count > 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				Pane.A();
			}
			else
			{
				Forms.InfoMessage(XC.A(7273));
			}
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(7358));
		}
		else
		{
			Forms.WarningMessage(XC.A(7399));
		}
		application = null;
	}

	private static void A(Document A, bool B)
	{
		List<ShapeItem> B2 = new List<ShapeItem>();
		try
		{
			_ = A.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.StoryRanges.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					do
					{
						enumerator2 = range.InlineShapes.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Check.A((InlineShape)enumerator2.Current, ref B2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator2 as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
						if (Helpers.A(range))
						{
							foreach (Microsoft.Office.Interop.Word.Shape item in range.ShapeRange)
							{
								Check.A(item, ref B2);
							}
						}
						range = range.NextStoryRange;
					}
					while (range != null);
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0123;
					}
					continue;
					end_IL_0123:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			if (B)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				Forms.ErrorMessage(ex2.Message);
			}
			ProjectData.ClearProjectError();
		}
		finally
		{
			LibraryObjects = B2;
		}
		B2 = null;
	}

	private static void A(InlineShape A, ref List<ShapeItem> B)
	{
		if (!Tagging.A(A))
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Check.B(A, ref B);
			return;
		}
	}

	private static void B(InlineShape A, ref List<ShapeItem> B)
	{
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		ContentInfo? val = Tagging.A(A);
		if (!val.HasValue)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ContentInfo value = val.Value;
			string libraryPath = Content.GetLibraryPath(value);
			if (Content.PathsAreDifferent(libraryPath, value))
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
				Tagging.A(A, libraryPath);
			}
			ManifestInfo? manifestInfo = Content.GetManifestInfo(libraryPath, value);
			if (manifestInfo.HasValue)
			{
				B.Add(new ShapeItem(A, value, manifestInfo.Value));
			}
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, ref List<ShapeItem> B)
	{
		if (A.Type == MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (2)
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
			if (!Tagging.A(A))
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = A.GroupItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Check.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, ref B);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		Check.B(A, ref B);
	}

	private static void B(Microsoft.Office.Interop.Word.Shape A, ref List<ShapeItem> B)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Unknown result type (might be due to invalid IL or missing references)
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		ContentInfo? val = Tagging.A(A);
		if (val.HasValue)
		{
			ContentInfo value = val.Value;
			string libraryPath = Content.GetLibraryPath(value);
			if (Content.PathsAreDifferent(libraryPath, value))
			{
				Tagging.A(A, libraryPath);
			}
			ManifestInfo? manifestInfo = Content.GetManifestInfo(libraryPath, value);
			if (manifestInfo.HasValue)
			{
				B.Add(new ShapeItem(A, value, manifestInfo.Value));
			}
		}
	}
}
