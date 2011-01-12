Example Firebreath plugin wrapper for ActiveX control
-----------------------------------------------------

Created by Peter Schellenbach (pjs@asent.com) January 2011,
for educational purposes. This is free software. Use at
your own risk.

This example includes a VB6 ActiveX control project used
to create FBExampleCtl.ocx. The UserControl in this project
acts like a command button. It has public properties,
methods and events which are manipulated by the Firebreath
plugin. The VB6 project and ocx are in the fbexamplectrl
directory.

The plugin project is in the axWrapper directory, and
was initially created with Firebreath 1.4a2. The plugin
uses ATL CAxWindow as an ActiveX control container, and
IDispEventSimpleImpl as an event sink for the wrapped
ActiveX control.

The technique used in this example should be fairly
straightforward to apply to other ActiveX controls.

This example plugin was tested with IE8, FireFox 3.6,
Chrome 8.0.552.224 and Safari 5.0.3. It works as expected
in all tested browsers, except that focus is not handled
properly in most browsers (IE works best).

A sample html page is included to illustrate basic
JavaScript functions to get/set plugin properties,
call methods and handle events.
