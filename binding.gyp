{
    "targets": [
      {
        "target_name": "focus",
        "sources": [ "focus.cpp" ],
        "include_dirs": [
          "<!@(node -p \"require('node-addon-api').include\")"
        ],
        "dependencies": [
          "<!(node -p \"require('node-addon-api').gyp\")"
        ]
      }
    ]
  }