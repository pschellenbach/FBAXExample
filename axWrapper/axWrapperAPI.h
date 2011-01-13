/**********************************************************\

  Auto-generated axWrapperAPI.h

\**********************************************************/

#include <string>
#include <sstream>
#include <boost/weak_ptr.hpp>
#include "JSAPIAuto.h"
#include "BrowserHost.h"
#include "axWrapper.h"

#ifndef H_axWrapperAPI
#define H_axWrapperAPI

class axWrapperAPI : public FB::JSAPIAuto
{
public:
    friend class axWrapperAxWin; // so ActiveX control container can fire events on JSAPI

    axWrapperAPI(axWrapperPtr plugin, FB::BrowserHostPtr host);
    virtual ~axWrapperAPI();

    axWrapperPtr getPlugin();

    // Read/Write property ${PROPERTY.ident}
    std::string get_Caption();
    void set_Caption(const std::string& val);

    // Read-only property ${PROPERTY.ident}
    std::string get_version();

	 // Methods
	 void FireClick(int duration);

private:
    axWrapperWeakPtr m_plugin;
    FB::BrowserHostPtr m_host;

};

#endif // H_axWrapperAPI

