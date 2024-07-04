#ifndef CUTIE_UTIL
#define CUTIE_UTIL

#include <string>

std::string WideStringToNarrowString(const std::wstring& wideString);
std::wstring NarrowStringToWideString(const std::string& narrowString);

#endif