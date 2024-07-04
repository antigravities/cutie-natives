let natives = {};

if( process.platform == "win32" ){
    const { FocusAssist } = require("./prebuilds/focus.node");
    natives["FocusAssist"] = FocusAssist;
}

module.exports = natives;