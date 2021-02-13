/*
	This file is part of Speech Plugin for Notepad++
	Copyright (C)2008 Jim Xochellis <dxoch at users.sourceforge.net>

	This program is free software; you can redistribute it and/or
	modify it under the terms of the GNU General Public License
	as published by the Free Software Foundation; either version 2
	of the License, or (at your option) any later version.

	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.

	You should have received a copy of the GNU General Public License
	along with this program; if not, write to the Free Software
	Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/

#ifndef _CRT_SECURE_NO_DEPRECATE
#define _CRT_SECURE_NO_DEPRECATE
#endif

#include "PluginInterface.h"
#include <shlobj.h>

#include <objbase.h>
#include <sapi.h>

#include <string>
#include <cassert>

const TCHAR PLUGIN_NAME[] = __TEXT("Speech");
//const char sectionName[] = "Speech Extesion";
const int nbFunc = 6;

NppData nppData;
bool gCanSpeek = false;
FuncItem funcItem[nbFunc];

void initMenu(void);
void Cleanup();

void SpeakDocument();
void SpeakSelection();
void StopSpeech();
void PauseSpeech();
void ResumeSpeech();

//---------------------------------------------------------

BOOL APIENTRY DllMain( HANDLE /*hModule*/, 
                       DWORD  reasonForCall, 
                       LPVOID /*lpReserved*/ )
{
	switch (reasonForCall)
	{
		case DLL_PROCESS_ATTACH:
		{
			funcItem[0]._pFunc = SpeakSelection;
			funcItem[1]._pFunc = SpeakDocument;
			funcItem[2]._pFunc = StopSpeech;
			funcItem[3]._pFunc = StopSpeech;
			funcItem[4]._pFunc = PauseSpeech;
			funcItem[5]._pFunc = ResumeSpeech;

			lstrcpy(funcItem[0]._itemName, __TEXT("Speak Selection"));
			lstrcpy(funcItem[1]._itemName, __TEXT("Speak Document"));
			lstrcpy(funcItem[2]._itemName, __TEXT("Stop Speech"));
			lstrcpy(funcItem[3]._itemName, __TEXT("-"));
			lstrcpy(funcItem[4]._itemName, __TEXT("Pause Speech"));
			lstrcpy(funcItem[5]._itemName, __TEXT("Resume Speech"));

			funcItem[0]._init2Check = false;
			funcItem[1]._init2Check = false;
			funcItem[2]._init2Check = false;
			funcItem[3]._init2Check = false;
			funcItem[4]._init2Check = false;
			funcItem[5]._init2Check = false;

			funcItem[0]._pShKey = NULL;
			funcItem[1]._pShKey = NULL;
			funcItem[2]._pShKey = NULL;
			funcItem[3]._pShKey = NULL;
			funcItem[4]._pShKey = NULL;
			funcItem[5]._pShKey = NULL;

			gCanSpeek = SUCCEEDED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED));
		}
		break;

		case DLL_PROCESS_DETACH:
			Cleanup();
			break;

		case DLL_THREAD_ATTACH:
			break;

		case DLL_THREAD_DETACH:
			break;
	}

	return TRUE;
}

extern "C" __declspec(dllexport) void setInfo(NppData notpadPlusData)
{
	nppData = notpadPlusData;
}

extern "C" __declspec(dllexport) const TCHAR * getName()
{
	return PLUGIN_NAME;
}

extern "C" __declspec(dllexport) FuncItem * getFuncsArray(int *nbF)
{
	*nbF = nbFunc;
	return funcItem;
}

HWND getCurrentHScintilla(int which)
{
	return (which == 0)?nppData._scintillaMainHandle:nppData._scintillaSecondHandle;
};

extern "C" __declspec(dllexport) void beNotified(SCNotification * /*notifyCode*/)
{
	//switch (notifyCode->nmhdr.code) { }
}

// Here you can process the Npp Messages 
// I will make the messages accessible little by little, according to the need of plugin development.
// Please let me know if you need to access to some messages :
// http://sourceforge.net/forum/forum.php?forum_id=482781
//
extern "C" __declspec(dllexport) LRESULT messageProc(UINT Message, WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	if (Message == WM_CREATE)
	{
		initMenu();
	}

	return TRUE;
}

#ifdef UNICODE
extern "C" __declspec(dllexport) BOOL isUnicode()
{
	return TRUE;
}
#endif

void initMenu(void)
{
	HMENU hMenu = ::GetMenu(nppData._nppHandle);

	if (hMenu != NULL)
		::ModifyMenu(hMenu, funcItem[3]._cmdID, MF_BYCOMMAND | MF_SEPARATOR, 0, 0);
}

//======================================================================

struct VoiceStruct
{
	VoiceStruct()
	{
		mVoice = NULL;
		mIsPaused= false;
	}

	ISpVoice *Voice() { return mVoice; }

	void Init(ISpVoice *pVoice)
	{
		this->Stop();

		mVoice = pVoice;
		mIsPaused= false;
	}

	void Stop()
	{
		if (mVoice != NULL)
		{
			mVoice->Release();
			mVoice = NULL;
			mIsPaused= false;
		}
	}

	void Pause()
	{
		if ( (mVoice != NULL) && ! mIsPaused )
		{
			mVoice->Pause();
			mIsPaused= true;
		}
	}

	void Resume()
	{
		if ( (mVoice != NULL) && mIsPaused )
		{
			mVoice->Resume();
			mIsPaused= false;
		}
	}

protected:
	ISpVoice *mVoice;
	bool mIsPaused;
};

VoiceStruct gVoice;

void SpeekText(std::wstring txt);
HWND GetCurrentEditHandle();

//---------------------------------------------------------

void SpeakDocument()
{
	HWND editHandle = GetCurrentEditHandle();

	assert(editHandle != NULL);

	if (editHandle != NULL)
	{
		DWORD txtLen = (long)SendMessage(editHandle, SCI_GETTEXTLENGTH, 0, 0);

		if ( txtLen > 0 )
		{		
			std::string buf(txtLen+1, 0);

			SendMessage(editHandle, SCI_GETTEXT, buf.size(), (LPARAM)&buf[0]);

			std::wstring wbuf(buf.begin(), buf.end());
			SpeekText(wbuf);
		}
	}
}

void SpeakSelection()
{
	HWND editHandle = GetCurrentEditHandle();

	assert(editHandle != NULL);

	if (editHandle != NULL)
	{
		struct Sci_TextRange tr;

		tr.chrg.cpMin = (long)SendMessage(editHandle, SCI_GETSELECTIONSTART, 0, 0);
		tr.chrg.cpMax = (long)SendMessage(editHandle, SCI_GETSELECTIONEND, 0, 0);

		if( tr.chrg.cpMax > 0 && (tr.chrg.cpMax > tr.chrg.cpMin))
		{
			std::string buf(tr.chrg.cpMax - tr.chrg.cpMin + 1, 0);
			
			tr.lpstrText = &buf[0];
			SendMessage(editHandle, SCI_GETTEXTRANGE, 0, (LPARAM)&tr);

			std::wstring wbuf(buf.begin(), buf.end());
			SpeekText(wbuf);
		}
		else
		{
			MessageBox(nppData._nppHandle, __TEXT("Please select some text first"), __TEXT("No selection"), 0);
		}
	}
}

void StopSpeech()
{
	gVoice.Stop();
}

void PauseSpeech()
{
	gVoice.Pause();
}

void ResumeSpeech()
{
	gVoice.Resume();
}

//---------------------------------------------------------

void SpeekText(std::wstring txt)
{
	if ( ! txt.empty() )
	{
		if ( gCanSpeek )
		{
			ISpVoice *pVoice = NULL;

			HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&pVoice);

			if ( SUCCEEDED(hr) )
			{
				gVoice.Init(pVoice);
				hr = gVoice.Voice()->Speak(txt.c_str(), SPF_ASYNC|SPF_PURGEBEFORESPEAK, NULL);
			}
		}
	}
}

HWND GetCurrentEditHandle()
{
	int currentEdit = 0;

	SendMessage(nppData._nppHandle, NPPM_GETCURRENTSCINTILLA, 0, (LPARAM)&currentEdit);

	return (currentEdit == 0) ? nppData._scintillaMainHandle : nppData._scintillaSecondHandle;
}

void Cleanup()
{
	if ( gCanSpeek )
	{
		//Stopping speech here leaves ghost threads!
		//StopSpeech();

		CoUninitialize();
		gCanSpeek = false;
	}
}
