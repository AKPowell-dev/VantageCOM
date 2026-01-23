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
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2.Versioning;

public sealed class Check
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ShapeItem, bool> A;

		public static Action A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A()
		{
			List<ShapeItem> libraryShapes = LibraryShapes;
			Func<ShapeItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (ShapeItem A) => ((ContentItem)A).IsOutdated);
			}
			else
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
				predicate = _Closure_0024__.A;
			}
			if (libraryShapes.Where(predicate).Count() <= 0)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (UIFormsExtensions.AskYesNo((System.Windows.Window)null, VH.A(211966), true, true))
				{
					Check.A();
				}
				return;
			}
		}

		[SpecialName]
		internal bool A(ShapeItem A)
		{
			return ((ContentItem)A).IsOutdated;
		}
	}

	[CompilerGenerated]
	internal sealed class QE
	{
		public Microsoft.Office.Interop.Excel.Workbook A;

		[SpecialName]
		internal void A()
		{
			try
			{
				Check.A(this.A, B: false);
				_disp.Invoke(_Closure_0024__.A.A);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	[CompilerGenerated]
	private static bool m_A;

	[CompilerGenerated]
	private static List<ShapeItem> m_A;

	public static Dispatcher _disp;

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

	internal static List<ShapeItem> LibraryShapes
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

	internal static void A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (!CheckOutdatedLibraryContent)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!Workbooks.IsValid(A))
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
				_disp = Dispatcher.CurrentDispatcher;
				Thread thread = new Thread([SpecialName] () =>
				{
					try
					{
						Check.A(A, B: false);
						_disp.Invoke(_Closure_0024__.A.A);
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
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.Workbooks.Count > 0)
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
			A(application.ActiveWorkbook, B: true);
			if (LibraryShapes.Count > 0)
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
				Pane.A();
			}
			else
			{
				Forms.InfoMessage(VH.A(82929));
			}
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, VH.A(83014));
		}
		else
		{
			Forms.WarningMessage(VH.A(83055));
		}
		application = null;
	}

	private static void A(Microsoft.Office.Interop.Excel.Workbook A, bool B)
	{
		List<ShapeItem> B2 = new List<ShapeItem>();
		try
		{
			IEnumerator enumerator = A.Worksheets.GetEnumerator();
			try
			{
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Worksheet worksheet = (Worksheet)enumerator.Current;
					try
					{
						enumerator2 = worksheet.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Check.A((Shape)enumerator2.Current, ref B2);
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
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
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0083;
					}
					continue;
					end_IL_0083:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			LibraryShapes = B2;
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
					switch (3)
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
		B2 = null;
	}

	private static void A(Shape A, ref List<ShapeItem> B)
	{
		if (A.Type == MsoShapeType.msoGroup)
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
			if (!Tagging.A(A))
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = A.GroupItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Check.A((Shape)enumerator.Current, ref B);
					}
					return;
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
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

	private static void B(Shape A, ref List<ShapeItem> B)
	{
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		ContentInfo? val = Tagging.A(A);
		if (!val.HasValue)
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
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				Tagging.A(A, libraryPath);
			}
			ManifestInfo? manifestInfo = Content.GetManifestInfo(libraryPath, value);
			if (!manifestInfo.HasValue)
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
				B.Add(new ShapeItem(A, value, manifestInfo.Value));
				return;
			}
		}
	}
}
