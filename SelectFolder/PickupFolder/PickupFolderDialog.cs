using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using System.Windows;

namespace PickupFolder
{
	public enum FDAP
	{
		BOTTOM = 0,
		TOP    = 1
	}
	/// <summary>
	/// Vista 形式のファイルダイアログによるフォルダ選択機能。
	/// 選択可能なフォルダは物理フォルダのみ
	/// </summary>
	public class PickupFolderDialog
	{
		/// <summary>
		/// 初期フォルダの指定(省略可)
		/// SelectedPath がセットされていない(or設定がおかしい)等の場合に利用される
		/// </summary>
		public string InitialFolder { get; set; }
		/// <summary>
		/// set = 前回設定していたパスの指定(省略可)
		/// get = ダイアログで選択したパス(Cancel終了の場合は設定されない)
		/// </summary>
		public string SelectedPath { get; set; }
		/// <summary>
		/// ダイアログのタイトル省略時は、「開く」
		/// </summary>
		public string Title { get; set; }
		/// <summary>
		/// プレースフォルダの追加(オプション)
		/// </summary>
		/// <param name="value"></param>
		public void AddPlace( string folder, FDAP fdap )
		{
			if( m_places == null )
			{
				m_places = new List<(string folder, FDAP fdap)>();
			}
			m_places.Add( ( folder, fdap ) );
		}
		#region WPF
		public bool ShowDialog()
		{
			return ShowDialog( Application.Current.MainWindow );
		}
		public bool ShowDialog( Window ownerWindow )
		{
			var hwndSrc = System.Windows.Interop.HwndSource.FromVisual( ownerWindow ) as System.Windows.Interop.HwndSource;
			return ShowDialog( hwndSrc != null ? hwndSrc.Handle : IntPtr.Zero );
		}
		#endregion
		#region InternalCode
		public bool ShowDialog( IntPtr ownerWindow )
		{
			//	オーナーウィンドウを正規化する
			ownerWindow = GetSafeOwnerWindow( ownerWindow );

			//	ダイアログインターフェースを構築
			IFileOpenDialog dlg = new FileOpenDialog() as IFileOpenDialog;  //	IUnknown::QueryInterfaceを使ってインターフェースを特定する
			try
			{
				//	フォルダ選択モードに切り替え
				dlg.SetOptions( FOS.FORCEFILESYSTEM | FOS.PICKFOLDERS );
				//	タイトル
				if( !string.IsNullOrWhiteSpace( Title ) )
				{
					dlg.SetTitle( Title );
				}
				//	ショートカット追加
				if( m_places.Count != 0 )
				{
					foreach( var place in m_places )
					{
						if( NativeMethods.SUCCEEDED( NativeMethods.SHCreateItemFromParsingName( place.folder, 
														IntPtr.Zero, typeof( IShellItem ).GUID, out var item ) ) )
						{
							dlg.AddPlace( item, place.fdap );
							Marshal.ReleaseComObject( item );
						}
					}
				}
				//	以前選択されていたフォルダを指定
				bool setFolder = false;
				if( !string.IsNullOrWhiteSpace( SelectedPath ) )
				{
					if( NativeMethods.SUCCEEDED( NativeMethods.SHCreateItemFromParsingName( SelectedPath,
													IntPtr.Zero, typeof(IShellItem).GUID, out var item ) ) )
					{
						dlg.SetFolder( item );
						Marshal.ReleaseComObject( item );
						setFolder = true;
					}
				}
				//	まだフォルダを設定していない場合は初期フォルダを設定する
				if( !setFolder && !string.IsNullOrWhiteSpace( InitialFolder ) )
				{
					if( NativeMethods.SUCCEEDED( NativeMethods.SHCreateItemFromParsingName( InitialFolder,
													IntPtr.Zero, typeof( IShellItem ).GUID, out var item ) ) )
					{
						dlg.SetFolder( item );
						Marshal.ReleaseComObject( item );
					}
				}
				//	ダイアログを表示
				var hRes = dlg.Show( ownerWindow );
				if( NativeMethods.SUCCEEDED( hRes ) )
				{
					var item = dlg.GetResult();
					SelectedPath = item.GetDisplayName( SIGDN.FILESYSPATH );
					Marshal.ReleaseComObject( item );
					return true;
				}
			}
			finally
			{
				Marshal.ReleaseComObject( dlg );
			}
			return false;
		}
		#endregion
		#region interop
		private static class NativeMethods
		{
			//	HWND サポート
			[DllImport( "user32.dll" )]
			[return: MarshalAs( UnmanagedType.Bool )]
			internal static extern bool IsWindow( IntPtr hWnd );
			[DllImport( "user32.dll" )]
			internal static extern IntPtr GetForegroundWindow();
			[DllImport( "user32.dll" )]
			internal static extern IntPtr GetParent( IntPtr hwnd );
			[DllImport( "user32.dll" )]
			internal static extern IntPtr GetLastActivePopup( IntPtr hwnd );
			//	Shell サポート(例外処理をしたくないので、HRESULT を受け取る)
			[DllImport( "shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true )]
			internal static extern int SHCreateItemFromParsingName(
				[In][MarshalAs( UnmanagedType.LPWStr )] string pszPath,
				[In] IntPtr pbc,
				[In][MarshalAs( UnmanagedType.LPStruct )] Guid riid,
				[Out][MarshalAs( UnmanagedType.Interface, IidParameterIndex = 2 )] out IShellItem ppv );
			//	HRESULT サポート
			internal static bool SUCCEEDED( int result ) => result >= 0;
			internal static bool FAILED( int result ) => result < 0;
		}
		#endregion
		#region COM Interop
		/// <summary>
		/// IShellItem シェル(エクスプローラ)がファイル等を仮想的にアイテムとして扱うためのインターフェース
		/// </summary>
		[
			ComImport,
			Guid( "43826D1E-E718-42EE-BC55-A1E261C37BFE" ),
			InterfaceType( ComInterfaceType.InterfaceIsIUnknown )
		]
		private interface IShellItem
		{
			void NotUsed(); // void BindToHandler();	名前まで、省略する例(通常はここまでしない)
			void GetParent(); // GetParent 省略
			/// <summary>
			/// このオブジェクトの文字列表記を取得
			/// </summary>
			/// <param name="sigdnName"></param>
			/// <returns>sigdnName に応じた文字列</returns>
			[return: MarshalAs( UnmanagedType.LPWStr )]
			string GetDisplayName( SIGDN sigdnName );
			void GetAttributes();  //  省略
			void Compare();  // Compare 省略
		}
		private enum SIGDN : uint // 今回使う識別子のみ移植
		{
			FILESYSPATH = 0x80058000,
		}
		//	ファイルダイアログ関連定義
		/// <summary>
		/// coclass FileOpenDialog
		/// </summary>
		[ComImport, Guid( "DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7" )] private class FileOpenDialog { }
		[
			ComImport,
			Guid( "42f85136-db7e-439c-85f1-e4075d135fc8" ),
			InterfaceType( ComInterfaceType.InterfaceIsIUnknown )
		]
		private interface IFileOpenDialog
		{
			//	IModalWindow
			[PreserveSig]
			int Show( IntPtr hwndParent );
			//	IFileDialog
			void SetFileTypes();
			void SetFileTypeIndex();
			void GetFileTypeIndex();
			void Advise();
			void Unadvise();
			void SetOptions( FOS fos );
			void GetOptions();
			void SetDefaultFolder();
			void SetFolder( IShellItem psi );
			void GetFolder();
			void GetCurrentSelection();
			void SetFileName();
			void GetFileName();
			void SetTitle( [MarshalAs( UnmanagedType.LPWStr )] string pszTitle );
			void SetOkButtonLabel();
			void SetFileNameLabel();
			IShellItem GetResult();
			void AddPlace( IShellItem item, FDAP fdap );
			void SetDefaultExtension();
			void Close();
			void SetClientGuid();
			void ClearClientData();
			void SetFilter();
			//	IFileOpenDialog
			void GetResults();
			void GetSelectedItems();
		}
		/// <summary>
		/// enum FOS(_FILEOPENDIALOGOPTIONSの略称)
		/// enum 名は、C/C++ 定義の略称名を利用。
		/// 今回使用するファイルシステム限定フラグ(フォルダの指定がメインなので)と
		/// フォルダ選択モードのフラグのみ転写。
		/// </summary>
		[Flags]
		private enum FOS
		{
			FORCEFILESYSTEM = 0x40,
			PICKFOLDERS = 0x20,
		}
		#endregion
		#region private members
		/// <summary>
		/// hwndOwner で指定されたウィンドウをオーナーウィンドウとして指定できるウィンドウに正規化する
		/// </summary>
		/// <param name="hwndOwner"></param>
		/// <returns></returns>
		internal static IntPtr GetSafeOwnerWindow( IntPtr hwndOwner )
		{
			//	無効なウィンドウを参照している場合の排除
			if( hwndOwner != IntPtr.Zero && !NativeMethods.IsWindow( hwndOwner ) )
			{
				hwndOwner = IntPtr.Zero;
			}
			//	オーナーウィンドウの基本を探す
			if( hwndOwner == IntPtr.Zero )
			{
				hwndOwner = NativeMethods.GetForegroundWindow();
			}
			//	トップレベルウィンドウを探す
			IntPtr hwndParent = hwndOwner;
			while( hwndParent != IntPtr.Zero )
			{
				hwndOwner = hwndParent;
				hwndParent = NativeMethods.GetParent( hwndOwner );
			}
			//	トップレベルウィンドウに所属する現在アクティブなポップアップ(自分も含む)を取得
			if( hwndOwner != IntPtr.Zero )
			{
				hwndOwner = NativeMethods.GetLastActivePopup( hwndOwner );
			}
			return hwndOwner;
		}
		private List<(string folder, FDAP fdap)> m_places;
		#endregion
	}
}
