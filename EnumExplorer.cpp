#include <wchar.h>
#include <locale.h>
#include <windows.h>
#include <shlobj.h>

int main(int argc, char *argv)
{
	setlocale(LC_CTYPE, ".ACP");

	HRESULT hr = CoInitialize(NULL);
	if (hr == S_OK)
	{
		IShellWindows *psw = NULL;
		hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&psw));
		if (hr == S_OK)
		{
			LONG n = 0;
			hr = psw->get_Count(&n);
			if (hr == S_OK)
			{
				IDispatch *pd = NULL;
				for (LONG i = 0; i < n; i++)
				{
					VARIANT vt;
					V_VT(&vt) = VT_I4;
					V_I4(&vt) = i;

					hr = psw->Item(vt, &pd);
					if (hr != S_OK)
					{
						continue;
					}

					IWebBrowserApp *pwb = NULL;
					hr = pd->QueryInterface(IID_PPV_ARGS(&pwb));
					pd->Release();
					if (hr != S_OK)
					{
						continue;
					}

					hr = pwb->get_Document(&pd);
					pwb->Release();
					if (hr != S_OK)
					{
						continue;
					}

					IShellFolderViewDual *psfv = NULL;
					hr = pd->QueryInterface(IID_PPV_ARGS(&psfv));
					pd->Release();
					if (hr != S_OK)
					{
						continue;
					}

					Folder *pf = NULL;
					hr = psfv->get_Folder(&pf);
					psfv->Release();
					if (hr != S_OK)
					{
						continue;
					}

					Folder2 *pf2 = NULL;
					hr = pf->QueryInterface(IID_PPV_ARGS(&pf2));
					pf->Release();
					if (hr != S_OK)
					{
						continue;
					}

					FolderItem *pfi = NULL;
					hr = pf2->get_Self(&pfi);
					pf2->Release();
					if (hr != S_OK)
					{
						continue;
					}

					BSTR pszPath = NULL;
					hr = pfi->get_Path(&pszPath);
					wprintf(L"[%s]\n", pszPath);

					pfi->Release();
				}
			}
			psw->Release();
		}
		CoUninitialize();
	}

	return 0;
}
