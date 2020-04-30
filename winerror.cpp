#include "winerror.h"

std::tstring WinError::GetString(DWORD error, HMODULE hModule)
{
    TCHAR* lpMsgBuf = nullptr;

    DWORD Result = FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | (hModule ? FORMAT_MESSAGE_FROM_HMODULE : FORMAT_MESSAGE_FROM_SYSTEM),
        hModule,
        error,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
        (LPTSTR) &lpMsgBuf,
        0,
        NULL
    );

    if (Result == 0 && hModule)
        Result = FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
            NULL,
            error,
            MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
            (LPTSTR) &lpMsgBuf,
            0,
            NULL
        );

    if (lpMsgBuf)
    {
        std::tstring Message(lpMsgBuf);
        LocalFree(lpMsgBuf);
        return Message;
    }
    else
        return std::tstring();
}
