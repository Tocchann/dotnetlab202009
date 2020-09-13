typedef struct IFileOpenDialogVtbl
{
    // IUnknown
    HRESULT( STDMETHODCALLTYPE* QueryInterface )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in REFIID riid,
        /* [annotation][iid_is][out] */ _COM_Outptr_  void** ppvObject);
    ULONG( STDMETHODCALLTYPE* AddRef )(
        __RPC__in IFileOpenDialog* This);
    ULONG( STDMETHODCALLTYPE* Release )(
        __RPC__in IFileOpenDialog* This);
    // IModalWindow
    /* [local] */ HRESULT( STDMETHODCALLTYPE* Show )(
        IFileOpenDialog* This,
        /* [annotation][unique][in] */ _In_opt_  HWND hwndOwner);
    // IFileDialog
    HRESULT( STDMETHODCALLTYPE* SetFileTypes )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ UINT cFileTypes,
        /* [size_is][in] */ __RPC__in_ecount_full( cFileTypes ) const COMDLG_FILTERSPEC* rgFilterSpec);
    HRESULT( STDMETHODCALLTYPE* SetFileTypeIndex )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ UINT iFileType);
    HRESULT( STDMETHODCALLTYPE* GetFileTypeIndex )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__out UINT* piFileType);
    HRESULT( STDMETHODCALLTYPE* Advise )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in_opt IFileDialogEvents* pfde,
        /* [out] */ __RPC__out DWORD* pdwCookie);
    HRESULT( STDMETHODCALLTYPE* Unadvise )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ DWORD dwCookie);
    HRESULT( STDMETHODCALLTYPE* SetOptions )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ FILEOPENDIALOGOPTIONS fos);
    HRESULT( STDMETHODCALLTYPE* GetOptions )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__out FILEOPENDIALOGOPTIONS* pfos);
    HRESULT( STDMETHODCALLTYPE* SetDefaultFolder )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in_opt IShellItem* psi);
    HRESULT( STDMETHODCALLTYPE* SetFolder )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in_opt IShellItem* psi);
    HRESULT( STDMETHODCALLTYPE* GetFolder )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__deref_out_opt IShellItem** ppsi);
    HRESULT( STDMETHODCALLTYPE* GetCurrentSelection )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__deref_out_opt IShellItem** ppsi);
    HRESULT( STDMETHODCALLTYPE* SetFileName )(
        __RPC__in IFileOpenDialog* This,
        /* [string][in] */ __RPC__in_string LPCWSTR pszName);
    HRESULT( STDMETHODCALLTYPE* GetFileName )(
        __RPC__in IFileOpenDialog* This,
        /* [string][out] */ __RPC__deref_out_opt_string LPWSTR* pszName);
    HRESULT( STDMETHODCALLTYPE* SetTitle )(
        __RPC__in IFileOpenDialog* This,
        /* [string][in] */ __RPC__in_string LPCWSTR pszTitle);
    HRESULT( STDMETHODCALLTYPE* SetOkButtonLabel )(
        __RPC__in IFileOpenDialog* This,
        /* [string][in] */ __RPC__in_string LPCWSTR pszText);
    HRESULT( STDMETHODCALLTYPE* SetFileNameLabel )(
        __RPC__in IFileOpenDialog* This,
        /* [string][in] */ __RPC__in_string LPCWSTR pszLabel);
    HRESULT( STDMETHODCALLTYPE* GetResult )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__deref_out_opt IShellItem** ppsi);
    HRESULT( STDMETHODCALLTYPE* AddPlace )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in_opt IShellItem* psi,
        /* [in] */ FDAP fdap);
    HRESULT( STDMETHODCALLTYPE* SetDefaultExtension )(
        __RPC__in IFileOpenDialog* This,
        /* [string][in] */ __RPC__in_string LPCWSTR pszDefaultExtension);
    HRESULT( STDMETHODCALLTYPE* Close )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ HRESULT hr);
    HRESULT( STDMETHODCALLTYPE* SetClientGuid )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in REFGUID guid);
    HRESULT( STDMETHODCALLTYPE* ClearClientData )(
        __RPC__in IFileOpenDialog* This);
    HRESULT( STDMETHODCALLTYPE* SetFilter )(
        __RPC__in IFileOpenDialog* This,
        /* [in] */ __RPC__in_opt IShellItemFilter* pFilter);
    // IFileOpenDialog
    HRESULT( STDMETHODCALLTYPE* GetResults )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__deref_out_opt IShellItemArray** ppenum);
    HRESULT( STDMETHODCALLTYPE* GetSelectedItems )(
        __RPC__in IFileOpenDialog* This,
        /* [out] */ __RPC__deref_out_opt IShellItemArray** ppsai);
    END_INTERFACE
} IFileOpenDialogVtbl;
interface IFileOpenDialog
{
    CONST_VTBL struct IFileOpenDialogVtbl* lpVtbl;
};
