/**********************************************************\

  Auto-generated axWrapper.cpp

  This file contains the auto-generated main plugin object
  implementation for the axWrapper project

\**********************************************************/
#ifndef H_axWrapperPLUGIN
#define H_axWrapperPLUGIN

#include "PluginWindow.h"
#include "Win\PluginWindowWin.h"
#include "PluginEvents/MouseEvents.h"
#include "PluginEvents/AttachedEvent.h"
#include "PluginEvents/WindowsEvent.h"

#include "PluginCore.h"

/**********************************************************\
 This Firebreath example is a wrapper for the accompanying
 sample xpcmdbutton VB6 ActiveX user-control (FBExample.vbp
 project).
\**********************************************************/

// Forward declare
class axWrapper;
class axWrapperAPI;

#ifdef WIN32

#include <atlwin.h>


// Import the ActiveX control's typelib so we can easily call methods, etc.
// on the ActiveX control.
#import "PROGID:FBExampleCtl.xpcmdbutton" no_namespace, raw_interfaces_only

// Define the ProgID for the ActiveX control.
#define AXCTLPROGID L"FBExampleCtl.xpcmdbutton"

// Define the ActiveX control's default & event interfaces. You might
// want to use a type library view like OleView.exe to identify interface
// names. Or you can find them in the registry.
#define AXCTLDEFINTF _xpcmdbutton
#define AXCTLEVTINTF __xpcmdbutton

// If the ActiveX control you are wrapping requires a license key,
// define that key here. Check Microsoft knowledge base article
// Q151771 for information on obtaining the license key for an
// ActiveX control.
//#define AXCTLLICKEY L"xxxxxxxxxxxx"

// Event function info structures are defined in axWrapper.cpp
extern _ATL_FUNC_INFO efiClick;
extern _ATL_FUNC_INFO efiKeyPress;


/**********************************************************\

 Define the ActiveX control container class using ATL's CAxWindow
 template with event support from ATL's IDispEventSimpleImpl
 tempalte.

 Note: if the ActiveX control that you are wrapping requires a
 license key, use CAxWindow2 instead of CAxWindow.

 \**********************************************************/
class axWrapperAxWin : public CAxWindow, public IDispEventSimpleImpl<1, axWrapperAxWin, &__uuidof(AXCTLEVTINTF)>
{
public:

	axWrapperAxWin(axWrapper* pPlugin) : m_pPlugin(pPlugin) {}
	virtual ~axWrapperAxWin() {}

	// Declare the events from the ActiveX control that we want to catch
	// in our FB plugin. See ATL documentation for IDispEventSimpleImpl for
	// details.
	BEGIN_SINK_MAP(axWrapperAxWin)
		SINK_ENTRY_INFO(1, __uuidof(AXCTLEVTINTF), DISPID_CLICK, onClick, &efiClick)
		SINK_ENTRY_INFO(1, __uuidof(AXCTLEVTINTF), DISPID_KEYPRESS, onKeyPress, &efiKeyPress)
	END_SINK_MAP()

	// Declare the ActiveX control event handler functions
	void __stdcall onClick();
	void __stdcall onKeyPress(short* pKeyCode);

	// Back pointer to containing plugin object (this is not using boost
	// shared_ptr because this object's lifetime is controlled excusively
	// by the plugin object's lifetime).
	axWrapper* m_pPlugin; 

};

#endif

/**********************************************************\
  This is the main Firebreath plugin class
\**********************************************************/
class axWrapper : public FB::PluginCore
{
public:
    static void StaticInitialize();
    static void StaticDeinitialize();

public:
    axWrapper();
    virtual ~axWrapper();

public:
    void onPluginReady();
    virtual FB::JSAPIPtr createJSAPI();
    virtual bool IsWindowless() { return false; }

    BEGIN_PLUGIN_EVENT_MAP()
        EVENTTYPE_CASE(FB::WindowsEvent, onWindowsMessage, FB::PluginWindowWin)
        EVENTTYPE_CASE(FB::MouseDownEvent, onMouseDown, FB::PluginWindow)
        EVENTTYPE_CASE(FB::MouseUpEvent, onMouseUp, FB::PluginWindow)
        EVENTTYPE_CASE(FB::MouseMoveEvent, onMouseMove, FB::PluginWindow)
        EVENTTYPE_CASE(FB::MouseMoveEvent, onMouseMove, FB::PluginWindow)
        EVENTTYPE_CASE(FB::AttachedEvent, onWindowAttached, FB::PluginWindow)
        EVENTTYPE_CASE(FB::DetachedEvent, onWindowDetached, FB::PluginWindow)
    END_PLUGIN_EVENT_MAP()

    /** BEGIN EVENTDEF -- DON'T CHANGE THIS LINE **/
	 virtual bool onWindowsMessage(FB::WindowsEvent *evt, FB::PluginWindowWin *);
    virtual bool onMouseDown(FB::MouseDownEvent *evt, FB::PluginWindow *);
    virtual bool onMouseUp(FB::MouseUpEvent *evt, FB::PluginWindow *);
    virtual bool onMouseMove(FB::MouseMoveEvent *evt, FB::PluginWindow *);
    virtual bool onWindowAttached(FB::AttachedEvent *evt, FB::PluginWindow *);
    virtual bool onWindowDetached(FB::DetachedEvent *evt, FB::PluginWindow *);
    /** END EVENTDEF -- DON'T CHANGE THIS LINE **/
private:

	FB::BrowserHostPtr getHost() {return m_host;}

	// ActiveX control initialization
	bool InitializeAxControl(const std::string& caption, int theme);

public:

	// ActiveX control property accessors
	std::string get_Caption();
	void set_Caption(const std::string& caption);

	int get_Theme();
	void set_Theme(int theme);

	// ActiveX control methods
	void FireClick(int duration);

	// provide access to JSAPI functions
   typedef boost::shared_ptr<axWrapperAPI> axWrapperAPIPtr;
	axWrapperAPIPtr getJSAPI() {return m_axwrapperapi;}

private:

	axWrapperAPIPtr m_axwrapperapi; // easy access to JSAPI functions
	#ifdef WIN32
		axWrapperAxWin m_axwin; // ActiveX control container instance (CAxWindow)
		CComPtr<AXCTLDEFINTF> m_spaxctl; // ActiveX control instance
	#endif

};
typedef boost::shared_ptr<axWrapper> axWrapperPtr;
typedef boost::weak_ptr<axWrapper> axWrapperWeakPtr;


#endif

