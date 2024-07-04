let natives = {};

if( process.platform == "win32" ){
    const { FocusAssist } = require("./build/Release/focus.node");
    const WinMedia = require("./build/Release/winmedia.node");

    natives["FocusAssist"] = FocusAssist;
    natives["WinMedia"] = WinMedia;
}

module.exports = natives;