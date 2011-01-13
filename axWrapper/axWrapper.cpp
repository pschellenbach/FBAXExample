/**********************************************************\

  Auto-generated axWrapper.cpp

  This file contains the auto-generated main plugin object
  implementation for the axWrapper project

\**********************************************************/

#include "NpapiTypes.h"
#include "variant_list.h"
#include "axWrapperAPI.h"
#include "axWrapper.h"

#ifdef SubclassWindow
#undef SubclassWindow // this is a macro defined in Windowsx.h - it conflicts with ATL CWindow::SubclassWindow function!
#endif

// Define the event function info structures for any events we want
// to catch. These structures are referenced in the event sink map
// used by IDispEventSimpleImpl (see axWrapper.h BEGIN_SINK_MAP).
_ATL_FUNC_INFO efiClick = {CC_STDCALL, VT_EMPTY, 0, {NULL}};
_ATL_FUNC_INFO efiKeyPress = {CC_STDCALL, VT_EMPTY, 1, {VT_I2 | VT_BYREF}};

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapper::StaticInitialize()
///
/// @brief  Called from PluginFactory::globalPluginInitialize()
///
/// @see FB::FactoryBase::globalPluginInitialize
///////////////////////////////////////////////////////////////////////////////
void axWrapper::StaticInitialize()
{
    // Place one-time initialization stuff here; note that there isn't an absolute guarantee that
    // this will only execute once per process, just a guarantee that it won't execute again until
    // after StaticDeinitialize is called
}

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapper::StaticInitialize()
///
/// @brief  Called from PluginFactory::globalPluginDeinitialize()
///
/// @see FB::FactoryBase::globalPluginDeinitialize
///////////////////////////////////////////////////////////////////////////////
void axWrapper::StaticDeinitialize()
{
    // Place one-time deinitialization stuff here
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  axWrapper constructor.  Note that your API is not available
///         at this point, nor the window.  For best results wait to use
///         the JSAPI object until the onPluginReady method is called
///////////////////////////////////////////////////////////////////////////////
axWrapper::axWrapper() : m_axwin(this)
{
	// Enumerate the supported parameters
	m_supportedParamSet.insert("caption");
	m_supportedParamSet.insert("theme");
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  axWrapper destructor.
///////////////////////////////////////////////////////////////////////////////
axWrapper::~axWrapper()
{
}

void axWrapper::onPluginReady()
{
    // When this is called, the BrowserHost is attached, the JSAPI object is
    // created, and we are ready to interact with the page and such.  The
    // PluginWindow may or may not have already fire the AttachedEvent at
    // this point.
}

///////////////////////////////////////////////////////////////////////////////
/// @brief  Creates an instance of the JSAPI object that provides your main
///         Javascript interface.
///
/// Note that m_host is your BrowserHost and shared_ptr returns a
/// FB::PluginCorePtr, which can be used to provide a
/// boost::weak_ptr<axWrapper> for your JSAPI class.
///
/// Be very careful where you hold a shared_ptr to your plugin class from,
/// as it could prevent your plugin class from getting destroyed properly.
///////////////////////////////////////////////////////////////////////////////
FB::JSAPIPtr axWrapper::createJSAPI()
{
    // m_host is the BrowserHost
    //return FB::JSAPIPtr(new axWrapperAPI(FB::ptr_cast<axWrapper>(shared_ptr()), m_host));
    m_axwrapperapi = axWrapperAPIPtr(new axWrapperAPI(FB::ptr_cast<axWrapper>(shared_ptr()), m_host));
	 return m_axwrapperapi;
}

bool axWrapper::onMouseDown(FB::MouseDownEvent *evt, FB::PluginWindow *)
{
    //printf("Mouse down at: %d, %d\n", evt->m_x, evt->m_y);
    return false;
}

bool axWrapper::onMouseUp(FB::MouseUpEvent *evt, FB::PluginWindow *)
{
    //printf("Mouse up at: %d, %d\n", evt->m_x, evt->m_y);
    return false;
}

bool axWrapper::onMouseMove(FB::MouseMoveEvent *evt, FB::PluginWindow *)
{
    //printf("Mouse move at: %d, %d\n", evt->m_x, evt->m_y);
    return false;
}
bool axWrapper::onWindowsMessage(FB::WindowsEvent *evt, FB::PluginWindowWin *piw)
{
	bool rc = false;
	if(!m_axwin)
		return false;
	switch(evt->uMsg)
	{
		case WM_SIZE:
			// resize the ActiveX container window
			m_axwin.MoveWindow(0, 0, LOWORD(evt->lParam), HIWORD(evt->lParam));
			evt->lRes = 0;
			rc = true;
			break;
		case WM_MOUSEACTIVATE:
			// activate on mouse click
			evt->lRes = MA_ACTIVATE;
			rc = true;
			break;
		case WM_SETFOCUS:
			// forward focus to the ActiveX control container window
			m_axwin.SetFocus();
			evt->lRes = 0;
			rc = true;
			break;
	}
	return rc;
}

bool axWrapper::onWindowAttached(FB::AttachedEvent *evt, FB::PluginWindow *piw)
{
    // The window is attached; act appropriately
#ifdef WIN32
	try {

		/* Now that we have the plugin window, create the ActiveX container
		   window as a child of the plugin, then create the ActiveX control
			as a child of the container.
		*/
		
		FB::PluginWindowWin* pwnd = piw->get_as<FB::PluginWindowWin>();
		if(pwnd != NULL)
		{
			HWND hWnd = pwnd->getHWND();
			if(hWnd)
			{				
				// Create the ActiveX control container
				RECT rc;
				::GetClientRect(hWnd, &rc);
				m_axwin.Create(hWnd, &rc, 0, WS_VISIBLE|WS_CHILD);

				// Create an instance of the ActiveX control in the container. If the ActiveX
				// control requires a license key, change CreateControlEx to CreateControlLicEx
				// and add one more parameter - CComBSTR(AXCTLLICKEY) - to the argument list.
				CComPtr<IUnknown> spControl;
				HRESULT hr = m_axwin.CreateControlEx(AXCTLPROGID, NULL, NULL, &spControl, GUID_NULL, NULL);
				if(SUCCEEDED(hr) && (spControl != NULL))
				{
					// Get the control's default interface
					spControl.QueryInterface(&m_spaxctl);
					if(m_spaxctl)
					{
						// Connect the event sink
						hr = m_axwin.DispEventAdvise((IUnknown*)m_spaxctl);

						// Get the initialization parameters
						std::string caption;
						int theme = 0;

						try {
							caption = m_params["caption"].convert_cast<std::string>();
						} catch(...) {} // ignore missing param
						try {
							theme = m_params["theme"].convert_cast<long>();
						} catch(...) {} // ignore missing param

						// Set ActiveX control initial properties using initialization paramters
						InitializeAxControl(caption, theme);

					}
				}
			}
		}
	} catch(...) {
		//TODO: should we throw a FB exception here?
	}
#endif

    return false;
}

bool axWrapper::onWindowDetached(FB::DetachedEvent *evt, FB::PluginWindow *)
{
    // The window is about to be detached; act appropriately
	if(m_spaxctl)
	{
		// Disconnect the event sink
		m_axwin.DispEventUnadvise((IUnknown*)m_spaxctl);		
		// Kill reference to the ActiveX control - when the plugin
		// window is destroyed, the container & control will be
		// automatically destroyed.
		m_spaxctl = NULL;
	}
    return false;
}

bool axWrapper::InitializeAxControl(const std::string& caption, int theme)
{
	set_Caption(caption);
	set_Theme(theme);
	return true;
}

std::string axWrapper::get_Caption()
{
	if(m_spaxctl)
	{
		try {
			CComBSTR bstr;
			HRESULT hr = m_spaxctl->get_Caption(&bstr);
			if(SUCCEEDED(hr))
				return FB::wstring_to_utf8(std::wstring(bstr.m_str, bstr.Length()));
		}
		catch(...) {
		}
	}
	return std::string(); // punt
}

void axWrapper::set_Caption(const std::string& caption)
{
	if(m_spaxctl)
	{
		try {
			HRESULT hr = m_spaxctl->put_Caption(CComBSTR(FB::utf8_to_wstring(caption).c_str()));
		}
		catch(...) {
		}
	}
}

int axWrapper::get_Theme()
{
	if(m_spaxctl)
	{
		try {
			ThemeConstants theme;
			HRESULT hr = m_spaxctl->get_Theme(&theme);
			if(SUCCEEDED(hr))
				return (int)theme;
		}
		catch(...) {
		}
	}
	return 0; // punt
}

void axWrapper::FireClick(int duration)
{
	if(m_spaxctl)
	{
		try {
			HRESULT hr = m_spaxctl->FireClick(duration);
		}
		catch(...) {
		}
	}
}

void axWrapper::set_Theme(int theme)
{
	if(m_spaxctl)
	{
		try {
			HRESULT hr = m_spaxctl->put_Theme(static_cast<ThemeConstants>(theme));
		}
		catch(...) {
		}
	}
}

/**********************************************************\

  implementation for the ActiveX control container class

\**********************************************************/

// ActiveX control event handlers

void __stdcall axWrapperAxWin::onKeyPress(short* pKey)
{
	if(pKey)
	{
		try {
			m_pPlugin->getJSAPI()->FireEvent("onkeypress", FB::variant_list_of(*pKey));
		}
		catch(...)
		{
		}
	}
}

void __stdcall axWrapperAxWin::onClick()
{
	try {
		m_pPlugin->getJSAPI()->FireEvent("onclick", FB::VariantList());
	}
	catch(...)
	{
	}
}

