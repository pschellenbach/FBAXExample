/**********************************************************\

  Auto-generated axWrapperAPI.cpp

\**********************************************************/

#include "JSObject.h"
#include "variant_list.h"
#include "DOM/Document.h"

#include "axWrapperAPI.h"

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapperAPI::axWrapperAPI(axWrapperPtr plugin, FB::BrowserHostPtr host)
///
/// @brief  Constructor for your JSAPI object.  You should register your methods, properties, and events
///         that should be accessible to Javascript from here.
///
/// @see FB::JSAPIAuto::registerMethod
/// @see FB::JSAPIAuto::registerProperty
/// @see FB::JSAPIAuto::registerEvent
///////////////////////////////////////////////////////////////////////////////
axWrapperAPI::axWrapperAPI(axWrapperPtr plugin, FB::BrowserHostPtr host) : m_plugin(plugin), m_host(host)
{
    // Read-write property
    registerProperty("caption",
                     make_property(this,
                        &axWrapperAPI::get_Caption,
                        &axWrapperAPI::set_Caption));

    // Read-only property
    registerProperty("version",
                     make_property(this,
                        &axWrapperAPI::get_version));
    
	 // Methods
	 registerMethod("fireclick",
							make_method(this,
								&axWrapperAPI::FireClick));
    // Events
    registerEvent("onclick");    
	 registerEvent("onkeypress");

}

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapperAPI::~axWrapperAPI()
///
/// @brief  Destructor.  Remember that this object will not be released until
///         the browser is done with it; this will almost definitely be after
///         the plugin is released.
///////////////////////////////////////////////////////////////////////////////
axWrapperAPI::~axWrapperAPI()
{
}

///////////////////////////////////////////////////////////////////////////////
/// @fn axWrapperPtr axWrapperAPI::getPlugin()
///
/// @brief  Gets a reference to the plugin that was passed in when the object
///         was created.  If the plugin has already been released then this
///         will throw a FB::script_error that will be translated into a
///         javascript exception in the page.
///////////////////////////////////////////////////////////////////////////////
axWrapperPtr axWrapperAPI::getPlugin()
{
    axWrapperPtr plugin(m_plugin.lock());
    if (!plugin) {
        throw FB::script_error("The plugin is invalid");
    }
    return plugin;
}



// Read/Write property Caption
std::string axWrapperAPI::get_Caption()
{
    return getPlugin()->get_Caption();
}
void axWrapperAPI::set_Caption(const std::string& val)
{
    getPlugin()->set_Caption(val);
}

// Read-only property version
std::string axWrapperAPI::get_version()
{
    return "CURRENT_VERSION";
}

// Method
void axWrapperAPI::FireClick(int duration)
{
	getPlugin()->FireClick(duration);
}

