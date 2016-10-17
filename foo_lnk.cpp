#include "../SDK/foobar2000.h"
#include "shlobj.h"

#define COMPONENT_NAME    "Shell Link Resolver"
#define COMPONENT_VERSION "1.3.1"
#define COMPONENT_ABOUT   "Resolves shell links (.lnk)\n\nwritten by Holger Stenger"

DECLARE_COMPONENT_VERSION(COMPONENT_NAME, COMPONENT_VERSION, COMPONENT_ABOUT);

DECLARE_FILE_TYPE("Shell Links","*.LNK");

static bool initialized = false;

void uResolveLink(const char * linkfile, pfc::string_base & path, HWND hwnd = NULL)
{
	HRESULT hres;
	pfc::com_ptr_t<IShellLinkA> psla;
	pfc::com_ptr_t<IShellLinkW> pslw;

	path.reset(); // assume failure 
	pfc::string8 short_path;

	if (!initialized)
	{
		hres = ::CoInitialize(NULL);

		initialized = !!SUCCEEDED(hres);

		if (FAILED(hres)) throw exception_win32(hres);
	}

	DWORD flags = hwnd ? 0 : SLR_NO_UI;

	// Get a pointer to the IShellLink interface (Unicode version). 
	hres = ::CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, 
		IID_IShellLinkW, (LPVOID*)pslw.receive_ptr());

	if (SUCCEEDED(hres))
	{
		pfc::com_ptr_t<IPersistFile> ppf;

		// Get a pointer to the IPersistFile interface. 
		hres = pslw->QueryInterface(IID_IPersistFile, (void**)ppf.receive_ptr());

		if (FAILED(hres)) throw exception_win32(hres);

		// Load the shortcut. 
		hres = ppf->Load(pfc::stringcvt::string_wide_from_utf8(linkfile), STGM_READ); 

		if (FAILED(hres)) throw exception_win32(hres);

		// Resolve the link.
		hres = pslw->Resolve(hwnd, flags);

		if (FAILED(hres)) throw exception_win32(hres);

		WCHAR szGotPath[MAX_PATH];
		WIN32_FIND_DATAW wfd; 
		// Get the path to the link target.
		hres = pslw->GetPath(szGotPath, MAX_PATH, &wfd, SLGP_SHORTPATH); 

		if (FAILED(hres)) throw exception_win32(hres);

		short_path.set_string(pfc::stringcvt::string_utf8_from_wide(szGotPath));
	}
	else
	{
		// Get a pointer to the IShellLink interface (ANSI version).
		hres = ::CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
			IID_IShellLinkA, (LPVOID*)&psla);

		if (FAILED(hres)) throw exception_win32(hres);

		pfc::com_ptr_t<IPersistFile> ppf;

		// Get a pointer to the IPersistFile interface. 
		hres = psla->QueryInterface(IID_IPersistFile, (void**)ppf.receive_ptr());

		if (FAILED(hres)) throw exception_win32(hres);

		// Load the shortcut. 
		hres = ppf->Load(pfc::stringcvt::string_wide_from_utf8(linkfile), STGM_READ); 

		if (FAILED(hres)) throw exception_win32(hres);

		// Resolve the link.
		hres = psla->Resolve(hwnd, flags); 

		if (FAILED(hres)) throw exception_win32(hres);

		char szGotPath[MAX_PATH];
		WIN32_FIND_DATAA wfd;

		// Get the path to the link target.
		hres = psla->GetPath(szGotPath, MAX_PATH, &wfd, SLGP_SHORTPATH);

		if (FAILED(hres)) throw exception_win32(hres);
		
		short_path.set_string(pfc::stringcvt::string_utf8_from_ansi(szGotPath));
	}

	if (!uGetLongPathName(short_path, path))
	{
		path.set_string(short_path, short_path.length());
	}
}

class link_resolver_lnk : public link_resolver
{
public:
	//! Tests whether specified file is supported by this link_resolver service.
	//! @param p_path Path of file being queried.
	//! @param p_extension Extension of file being queried. This is provided for performance reasons, path already includes it.
	virtual bool is_our_path(const char * p_path, const char * p_extension)
	{
		return !pfc::stricmp_ascii(p_extension, "lnk");
	}

	//! Resolves a link file. Before this is called, path must be accepted by is_our_path().
	//! @param p_filehint Optional file interface to use. If null/empty, implementation should open file by itself.
	//! @param p_path Path of link file to resolve.
	//! @param p_out Receives path the link is pointing to.
	//! @param p_abort abort_callback object signaling user aborting the operation.
	virtual void resolve(service_ptr_t<file> p_filehint, const char * p_path, pfc::string_base & p_out, abort_callback & p_abort)
	{
		TRACK_CALL_TEXT("link_resolver_lnk::resolve");

		if (stricmp_utf8_partial("file://", p_path, 7))
		{
			throw exception_io_data(pfc::string_formatter() << "shortcut is not on local filesystem:\n" << p_path);
		}
		else
		{
			const char * linkfile = p_path+7;
			uResolveLink(linkfile, p_out);
		}
	}
};

static service_factory_single_t<link_resolver_lnk> foo_link_resolver;

class initquit_lnk : public initquit
{
public:
	virtual void on_init() {}

	virtual void on_quit()
	{
		if (initialized)
		{
			CoUninitialize();
		}
	}
};

static initquit_factory_t< initquit_lnk > foo_initquit;
