#pragma once

#include <string>
#include <vector>

void Xlsx2DwgProp_Command();
bool Xlsx2DwgProp_GetDescriptionForKey(const std::string& key, std::string& outDescription);
bool Xlsx2DwgProp_GetTrackedFileStatus(std::string& fileNameOnly, std::string& fullPath, bool& hashMismatch);
int Xlsx2DwgProp_GetHashCheckMinutes();
bool Xlsx2DwgProp_GetGroupForKey(const std::string& key, std::string& outGroup);
void Xlsx2DwgProp_GetGroups(std::vector<std::string>& outGroups);
