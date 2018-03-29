#include "stdafx.h"
#include <initguid.h>
#include <MAPIX.h>
#include <string>

void DisplayUsage()
{
	printf("FixContab - MAPI Profile Contab Fix\n");
	printf("   Puts Contab first in Address Book load order to fix MAPI crashes.\n");
	printf("\n");
	printf("Usage:  FixContab profile\n");
	printf("\n");
	printf("Options:\n");
	printf("        profile - Name of profile to modify.\n");
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

	printf("Modifying profile %s", profile.c_str());

	return 0;
}