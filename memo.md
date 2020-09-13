# .NET Lab 2020/09

2020/09/26 開催  
時間は30分  
無駄話できるほどの時間はなさそう。

# タイトル

「デスクトップの .NET アプリで IFileOpenDialog 版フォルダ選択ダイアログを使ってみよう」

Microsoft MVP for Developer Technologies とっちゃん(高萩俊行)


Vista からファイルダイアログが刷新されました。ファイル選択については .NET アプリでも新しいファイルダイアログを使うようになりましたが、フォルダ選択はなぜかいつまでたっても用意されません。

ということで何番煎じかわかりませんが、商利用でも耐えられる程度の耐性を持つ最小限実装を作ってみました。

最小限なので、COMの定義も最小限です。どこが省略出来て、どれが省略できないのか？なんてことを IFileOpenDialog の C#版呼び出しコードを例にお話ししてみたいと思います。

## 目次的な何か

- フォルダ選択ダイアログ
- COM ってなんだっけ？
- C# から使えるようにするために！
- まとめ

## フォルダ選択ダイアログ

- 今のWindowsには２種類ある
- 9x時代(確か98)に用意されたAPI
  - SHBrowseForFolder
  - System.Windows.Forms.FolderBrowserDialog
  - 一度は 新しいフォルダ選択を推奨とされていたのだけど...
- Vista から用意されたAPI
  - COM ベース
  - ファイル選択と同じUIに統一
  - インターフェースが統一されたことで操作系も統一された
- いまどきなら FolderPicker じゃ？
  - UWP アプリじゃないと使えないんです。

### .NET 標準ライブラリ内では？

- System.Windows.Forms.FolderBrowserDialog
  - 名前の通り、Forms 用
- WPF 用は、標準提供されていない。
  - Windows API Code Pack があったのだけど...
    - 気づいたら消えていました。
- ファイル選択ダイアログは新しいのに。。。
  - 統一感がないのは、お客様のそうｓ。。。

### 新しいフォルダ選択ダイアログが使いたいの！

- COM インターフェースにアクセスできるようにする必要がある
  - クライアントサイド側にもインターフェースの定義が必要
    - 通常は提供元が用意するんじゃない？
    - COM定義は用意されている(C/C++とIDLだけだけどねｗ)

## COMってなんだっけ？

- COM:Component Object Model の略。
- ABI(Application Binary Interface)規約の一種
- 基本原則
  - インターフェースを介した通信形式
  - 実行時に提供元から動的にオブジェクトを取得する
- COM インターフェースの形式は？
  - 関数テーブル(関数ポインタの構造体) 構造
    - C++ 仮想関数テーブル互換
  - マングリング問題(名前装飾)が発生しない
- インターフェースは Windows に依存していない。

### 復習。IUnknown のインターフェース定義

#### IDL の定義

```midl
[
  local,
  object,
  uuid(00000000-0000-0000-C000-000000000046),

  pointer_default(unique)
]
interface IUnknown
{
    typedef [unique] IUnknown *LPUNKNOWN;
    HRESULT QueryInterface( [in] REFIID riid,
        [out, iid_is(riid), annotation("_COM_Outptr_")] void **ppvObject );
    ULONG AddRef();
    ULONG Release();
}
```

#### C++ の場合(マクロ定義)

```
MIDL_INTERFACE("00000000-0000-0000-C000-000000000046")
IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE QueryInterface( REFIID riid, _COM_Outptr_  void **ppvObject ) = 0;
    virtual ULONG STDMETHODCALLTYPE AddRef(void) = 0;
    virtual ULONG STDMETHODCALLTYPE Release(void) = 0;
};
```

#### C の場合

```
typedef struct IUnknownVtbl
{
    HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
        IUnknown * This, /* [in] */ REFIID riid, /* [annotation][iid_is][out] */ _COM_Outptr_  void **ppvObject );
    ULONG ( STDMETHODCALLTYPE *AddRef )( IUnknown * This );
    ULONG ( STDMETHODCALLTYPE *Release )( IUnknown * This );
} IUnknownVtbl;

interface IUnknown
{
    CONST_VTBL struct IUnknownVtbl *lpVtbl;
};

#ifdef COBJMACROS
#define IUnknown_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 
#define IUnknown_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 
#define IUnknown_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 
#endif /* COBJMACROS */
```

#### 一部マクロ展開版(C/C++)

```
struct __declspec(uuid("00000000-0000-0000-C000-000000000046"))
       __declspec(novtable)
IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE QueryInterface( const IID & riid, void **ppvObject ) = 0;
    virtual ULONG STDMETHODCALLTYPE AddRef() = 0;
    virtual ULONG STDMETHODCALLTYPE Release() = 0;
};
```
```
typedef struct IUnknownVtbl
{
    HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
        IUnknown * This, const IID *riid, void **ppvObject );
    ULONG ( STDMETHODCALLTYPE *AddRef )( IUnknown * This );
    ULONG ( STDMETHODCALLTYPE *Release )( IUnknown * This );
} IUnknownVtbl;

struct IUnknown
{
    const struct IUnknownVtbl *lpVtbl;
};
```

### オブジェクトの構築に必要なもの

- CLSID(coclassとも呼ぶ) と呼ばれるオブジェクト識別子(GUID)で定義
- レジストリに実態構築のための情報を記述しておく
- CoCreateInstance APIで、オブジェクト(インスタンス)を構築する
- 実装詳細はクライアントからは見えない
  - インターフェースでのやり取りのみ

## 新しいファイルダイアログのインターフェース

- Windows Vista から新規に提供 COM インターフェース
- 実態は、新しいコモンダイアログに含まれている
  - comdlg32.dll が実装提供
  - 定義は Windows SDK で提供
    - C/C++ 向け
      - C#？ 知らない子ですね

### coclass(CLSID)

- CLSID_FileOpenDialog

IDL
```
[ uuid(DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7) ]
coclass FileOpenDialog
{
  interface IFileOpenDialog;
}
```

C++
```
class DECLSPEC_UUID("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")
FileOpenDialog;
```

C
```
EXTERN_C const CLSID CLSID_FileOpenDialog;
// 変数そのものは、スタティックライブラリをリンクする
```

## インターフェース(メソッド省略)

```
interface IUnknown{} // メソッド数 : 3
interface IModalWindow : public IUnknown{} // メソッド数 : 1
interface IFileDialog : public IModalWindow{} // メソッド数 : 23
interface IFileOpenDialog : public IFileDialog{} // メソッド数 : 2
```
- これら以外に、オブジェクトのやり取り用インターフェースも存在。
- イベント、カスタマイズなどのインターフェースも別途存在(今回は使わないので省略)

### C# で定義するにはどうすれば。。。

- IDL があるんだから、タイプライブラリもあるよね？
  - NOTE - this typelib is never registered anywhere
  - ビルドができないとは書かれていないけどね。
- IDL から頑張って書き起こすの？
  - その通りです！
- 25個もメソッド記述するの？
  - いいえ！そんなことはありません。

### COMってなんだっけ？(抜粋)

- COM:Component Object Model の略。
- ABI(Application Binary Interface)規約の一つ。
- インターフェースを介して通信
- 実行時に動的に接続

- COM インターフェースの形式
  - 関数テーブル(関数ポインタの構造体) 構造
    - C++ 仮想関数テーブル互換
  - マングリング問題(名前装飾)が発生しない


### 派生インターフェースの例

### IDL の定義

```
[
    uuid(b4db1657-70d7-485e-8e3e-6fcb5a5c1802),
    object,
    pointer_default(unique)
]
interface IModalWindow : IUnknown
{
    [local] HRESULT Show([in, unique, annotation("_In_opt_")] HWND hwndOwner);
    [call_as(Show)] HRESULT RemoteShow([in, unique] HWND hwndOwner);
}
```

### C++ の定義

```
MIDL_INTERFACE("b4db1657-70d7-485e-8e3e-6fcb5a5c1802")
IModalWindow : public IUnknown
{
public:
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Show( 
        /* [annotation][unique][in] */ 
        _In_opt_  HWND hwndOwner) = 0;
};
```

### C の定義

```
typedef struct IModalWindowVtbl
{
    HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
        __RPC__in IModalWindow * This,
        /* [in] */ __RPC__in REFIID riid,
        /* [annotation][iid_is][out] */ 
        _COM_Outptr_  void **ppvObject);
    ULONG ( STDMETHODCALLTYPE *AddRef )( 
        __RPC__in IModalWindow * This);
    ULONG ( STDMETHODCALLTYPE *Release )( 
        __RPC__in IModalWindow * This);

    /* [local] */ HRESULT ( STDMETHODCALLTYPE *Show )( 
        IModalWindow * This,
        /* [annotation][unique][in] */ 
        _In_opt_  HWND hwndOwner);
} IModalWindowVtbl;

interface IModalWindow
{
    CONST_VTBL struct IModalWindowVtbl *lpVtbl;
};
```

### C# の定義

```
[ComImport]
[Guid( "b4db1657-70d7-485e-8e3e-6fcb5a5c1802" )]
[InterfaceType( ComInterfaceType.InterfaceIsIUnknown )]
interface IModalWindow
{
    [PreserveSig]
    int Show( [In] IntPtr hwndParent );
}
```

### IFileDialog の定義(抜粋)

IDL

```
[
    uuid(42f85136-db7e-439c-85f1-e4075d135fc8),
    object,
    pointer_default(unique)
]
interface IFileDialog : IModalWindow
{
    HRESULT SetFileTypes(
        [in] UINT cFileTypes,
        [in, size_is(cFileTypes)] const COMDLG_FILTERSPEC *rgFilterSpec);
    HRESULT SetFileTypeIndex([in] UINT iFileType);
    HRESULT GetFileTypeIndex([out] UINT *piFileType);
    HRESULT Advise(
        [in] IFileDialogEvents *pfde,
        [out] DWORD *pdwCookie);
    HRESULT Unadvise([in] DWORD dwCookie);
    // 省略
}

```

C++

```
MIDL_INTERFACE("42f85136-db7e-439c-85f1-e4075d135fc8")
IFileDialog : public IModalWindow
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetFileTypes( 
        /* [in] */ UINT cFileTypes,
        /* [size_is][in] */ __RPC__in_ecount_full(cFileTypes) const COMDLG_FILTERSPEC *rgFilterSpec) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE SetFileTypeIndex( 
        /* [in] */ UINT iFileType) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE GetFileTypeIndex( 
        /* [out] */ __RPC__out UINT *piFileType) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE Advise( 
        /* [in] */ __RPC__in_opt IFileDialogEvents *pfde,
        /* [out] */ __RPC__out DWORD *pdwCookie) = 0;
        
    virtual HRESULT STDMETHODCALLTYPE Unadvise( 
        /* [in] */ DWORD dwCookie) = 0;
    // 省略
};
```

C

```
typedef struct IFileDialogVtbl
{
// IUnknown 部分
    HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
        __RPC__in IFileDialog * This,
        /* [in] */ __RPC__in REFIID riid,
        /* [annotation][iid_is][out] */ 
        _COM_Outptr_  void **ppvObject);
    ULONG ( STDMETHODCALLTYPE *AddRef )( 
        __RPC__in IFileDialog * This);
    ULONG ( STDMETHODCALLTYPE *Release )( 
        __RPC__in IFileDialog * This);
// IModalWindow 部分
    /* [local] */ HRESULT ( STDMETHODCALLTYPE *Show )( 
        IFileDialog * This,
        /* [annotation][unique][in] */ 
        _In_opt_  HWND hwndOwner);
// IFileDialog 部分
    HRESULT ( STDMETHODCALLTYPE *SetFileTypes )( 
        __RPC__in IFileDialog * This,
        /* [in] */ UINT cFileTypes,
        /* [size_is][in] */ __RPC__in_ecount_full(cFileTypes) const COMDLG_FILTERSPEC *rgFilterSpec);
    HRESULT ( STDMETHODCALLTYPE *SetFileTypeIndex )( 
        __RPC__in IFileDialog * This,
        /* [in] */ UINT iFileType);
    HRESULT ( STDMETHODCALLTYPE *GetFileTypeIndex )( 
        __RPC__in IFileDialog * This,
        /* [out] */ __RPC__out UINT *piFileType);
    HRESULT ( STDMETHODCALLTYPE *Advise )( 
        __RPC__in IFileDialog * This,
        /* [in] */ __RPC__in_opt IFileDialogEvents *pfde,
        /* [out] */ __RPC__out DWORD *pdwCookie);
    HRESULT ( STDMETHODCALLTYPE *Unadvise )( 
        __RPC__in IFileDialog * This,
        /* [in] */ DWORD dwCookie);
    /* 省略 */
};
```

## これらを踏まえてC#で定義すると...

C# パターン１(分離表記型 - .NET Core Source Browser では、明示的に new でメソッドを再定義しているけど、いらないはずでは？)
```
[ComImport]
[Guid( "42f85136-db7e-439c-85f1-e4075d135fc8" )]
[InterfaceType( ComInterfaceType.InterfaceIsIUnknown )] // IUnknown 部分
interface IFileDialog : IModalWindow
{
// IModalWindow 部分(.NET Core では明示的に記載しているのであえて再定義。いらないと思うんだけど。。。)
    [PreserveSig]
    new int Show( [In] IntPtr hwndParent );
// IFileDialog 部分
    void SetFileTypes( [In] uint cFileTypes, [in MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] COMDLG_FILTERSPEC[] rgFilterSpec );
    void SetFileTypeIndex( [In] uint iFileType );
    uint GetFileTypeIndex();
    uint Advise( [In] IFileDialogEvents pfde );
    void Unadvise( [in] uint dwCookie );
    // ...
}
```

C# パターン２(合成表記型)
```
[ComImport]
[Guid( "42f85136-db7e-439c-85f1-e4075d135fc8" )]
[InterfaceType( ComInterfaceType.InterfaceIsIUnknown )] // IUnknown 部分
interface IFileDialog
{
// IModalWindow 部分
    [PreserveSig]
    int Show( [In] IntPtr hwndParent );
// IFileDialog 部分
    void SetFileTypes( [In] uint cFileTypes, [in MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] COMDLG_FILTERSPEC[] rgFilterSpec );
    void SetFileTypeIndex( [In] uint iFileType );
    uint GetFileTypeIndex();
    uint Advise( [In] IFileDialogEvents pfde );
    void Unadvise( [in] uint dwCookie );
    // ...
}
```

C# パターン３(必要最低限。抜粋)
```
[ComImport]
[Guid( "42f85136-db7e-439c-85f1-e4075d135fc8" )]
[InterfaceType( ComInterfaceType.InterfaceIsIUnknown )] // IUnknown 部分
interface IFileDialog
{
// IModalWindow 部分
    [PreserveSig]
    int Show( [In] IntPtr hwndParent );
// IFileDialog 部分
    void SetFileTypes();
    void SetFileTypeIndex();
    void GetFileTypeIndex();
    void Advise();
    void Unadvise();
    // ...
}
```

### 必要な定義とは？

- **COM ABI として維持できる**必要**十分**な**内容を満たす**定義
  - バイナリレベルの情報が正当であること。
  - 正しい順番で関数ポインタが構造化されていること。
  - COMにとって互換となるパラメータと戻り値になっていること。
    - .NET の COMサポートには一定の特別な規約がある
    - HRESULT が必要な場合は、PreserveSigAttribute を付ける

### COM ABI として維持できる必要十分な内容を満たす定義

- RCW に関する基本的な知識
  - そもそもなんだっけ？では厳しい
  - out パラメータの扱い
  - PreserveSig 属性
- IUnkonown 以外のすべてのインターフェースのメソッドの一覧
  - 使わないメソッドの名前はいらない
  - 使わないメソッドの引数はいらない
  - キャストしないなら、直接IUnknownから派生しても問題ない

### IFileOpenDialog はどこにある？

- Windows SDK にあります！
  - C++ ワークロードだとデフォルトで入った気がするけど。。。
  - VSの個別コンポーネントからもインストール可能
- Visual Studioは持ってないんだけど？
  - 単体版をどうぞ！
    - 「Windows 10 SDK」で検索！

ShObjIdl_core.idl からコピーしてくるよ！

1. IFileOpenDialog の定義に移動
2. インターフェースの属性も含めてコピー
3. C# ソースの適当なところにペースト
4. 親インターフェースが IUnknown じゃなければ、その定義に移動
5. 2 に戻る

### 最低限のフォルダ選択機能を利用する際に使われるもの

- IModalWindow.Show()
  - 表示するのに必要
- IFileDialog.SetOptions()
  - フォルダ選択モードのために必要
- IFileDialog.GetResult()
  - 選択したフォルダ情報を取得するために必要
- IFileDialog.SetFolder()
  - 以前のフォルダを指定するために必要
- IFileDialog.SetTitle()
  - ダイアログキャプションを指定するために必要
- IFileDialog.AddPlace()
  - アプリケーション用フォルダショートカットの追加のために必要(オプション的存在)

### で、シュリンクすると。。。

実ソースを出す。

おまけコードの説明は今回は割愛かなぁ？
