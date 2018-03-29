#include "stdafx.h"
#include <initguid.h>
#include <MAPIX.h>
#include <string>

void DisplayUsage()
{
	wprintf(L"FixContab - MAPI Profile Contab Fix\n");
	wprintf(L"   Puts Contab first in Address Book load order to fix MAPI crashes.\n");
	wprintf(L"\n");
	wprintf(L"Usage:  FixContab profile\n");
	wprintf(L"\n");
	wprintf(L"Options:\n");
	wprintf(L"        profile - Name of profile to modify.\n");
}

int main(int argc, char* argv[])
{
	std::string profile;
	if (argc == 2)
	{
		profile = argv[1];
	}
	else
	{
		DisplayUsage();
	}

	wprintf(L"Modifying profile %hs", profile.c_str());

	return 0;
}