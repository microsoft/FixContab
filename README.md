# FixContab
Simple program to reorder MAPI providers so the Contact Address Book is first.

# The Problem
Suppose you've created a profile programatically. Perhaps you followed Dave's steps given [here](https://blogs.msdn.microsoft.com/dvespa/2015/10/28/how-to-configure-an-outlook-2016-profile-using-mfcmapi) or [here](https://blogs.msdn.microsoft.com/dvespa/2017/06/19/create-outlook-2013-profile-mapi-http/). Those steps focus (rightly) on getting the profile to connect to Exchange, since that's the hard part. One thing you might not have considered, is adding the Contact Address Book (aka contab). That's fine - when you use the profile the first time, MAPI will add the provider for you!
And this is where we hit a problem: MAPI assumes that contab will be the first address book provider loaded. When it's not, it gets all flustered, and you have crashes in programs other than Outlook. For instance, using Send To -> Mail Recipient explorer might crash with the error "FIXMAPI 1.0 MAPI Repair Tool has stopped working". This issue [spawned](https://social.technet.microsoft.com/Forums/en-US/ba2dfcf0-d5a8-4530-8b0c-c8ac06789e9b/fixmapi-10-mapi-repair-tool-has-stopped-working-outlook-2016-32bit-version-all-send-to-email-mapi?forum=outlook) [many](https://community.spiceworks.com/topic/2024964-fixmapi-1-0-mapi-repair-tool-has-stopped-working-in-office-2016-c2r?page=1#entry-7157347) [forum reports](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_install-mso_win10-mso_o365b/fixmapi-10-mapi-repair-tool-has-stopped-working/e8eeeee3-0fe8-4c82-b94b-cd3db29013d7).

If you're running Office 365 (aka Click To Run aka c2r), then we fixed this in build [16.0.8420.1000](https://support.office.com/en-us/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7?ui=en-US&rs=en-US&ad=US). But if you're running Office 2016 MSI, or still have Outlook 2013, you're out of luck. The fix was too complicated to backport without risking further breaks.

# The Solution
As noted in the forum reports, deleting and recreating the profile in Outlook solved the problem. This is because in new profile creation, Outlook always adds contab first.

But maybe you don't want to recreate the profile. Maybe you want to use MFCMAPI to fix it? These steps work, but are super complicated:
1. Start [MFCMAPI](https://github.com/stephenegriffin/mfcmapi/releases) with Outlook closed
2. Profile/Show Profiles
3. Double-click profile
4. Locate and click on the row with PR_DISPLAY_NAME of Outlook Address Book and PR_SERVICE_NAME of CONTAB.
5. In the bottom pane, look for PR_AB_PROVIDERS.
6. Open this property and copy out the Binary section. It should be 16 bytes/32 characters. For instance, this is what I have:  
69DD16E10828474F9F31C1F39A27111D
7. Actions/Open Profile Section
8. Paste in the following guid and check the “Byte swapped” checkbox:  
9207F3E0A3B11019908B08002B2A56C2
9. A new set of properties should load in the bottom pane, including one named PR_AB_PROVIDERS
10. Open this property and copy out the Binary section. The count of bytes of this section should be a multiple of 16, for a multiple of 32 characters. For instance, in a broken profile I built, it looked like this:  
EE9B4F3DEAC0BC4CA3BBCA0F842FED3E69DD16E10828474F9F31C1F39A27111D
11. Divide this string in to 32 character chunks, like so:  
EE9B4F3DEAC0BC4CA3BBCA0F842FED3E  
69DD16E10828474F9F31C1F39A27111D
12. See that one of those chunks is the string we copied from the PR_AB_PROVIDERS property for Contab:  
69DD16E10828474F9F31C1F39A27111D
13. We want this entry to be first in the list, so move it to the top:  
69DD16E10828474F9F31C1F39A27111D  
EE9B4F3DEAC0BC4CA3BBCA0F842FED3E
14. And merge the list back to a single string:  
69DD16E10828474F9F31C1F39A27111DEE9B4F3DEAC0BC4CA3BBCA0F842FED3E
15. Back in the property we were looking at in step 10, replace the original Binary section with the new string.
16. Close the dialog (this will save) and close MFCMAPI.
17. Start Outlook.
