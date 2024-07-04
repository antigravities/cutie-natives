#define NAPI_CPP_EXCEPTIONS 1

#include "util.h"
#include <string>
#include <Windows.h>
#include <stdexcept>

std::string WideStringToNarrowString(const std::wstring& wideString) {
    // Calculate the length of the resulting narrow string
    int narrowStringLength = WideCharToMultiByte(CP_UTF8, 0, wideString.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (narrowStringLength == 0) {
        throw std::runtime_error("Failed to convert wide string to narrow string.");
    }

    // Allocate a buffer for the narrow string
    std::string narrowString(narrowStringLength, 0);

    // Perform the conversion
    int result = WideCharToMultiByte(CP_UTF8, 0, wideString.c_str(), -1, &narrowString[0], narrowStringLength, nullptr, nullptr);
    if (result == 0) {
        throw std::runtime_error("Failed to convert wide string to narrow string.");
    }

    // The result includes the null terminator, so we remove it
    narrowString.pop_back();

    return narrowString;
}