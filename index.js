let natives = {};

if( process.platform == "win32" ){
    const { FocusAssist } = require("./build/Release/focus.node");
    natives["FocusAssist"] = FocusAssist;
}

module.exports = natives;